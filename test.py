r"""
solidworks_export.py
=====================

یک اسکریپت واحد برای دو کار:

  1) drawing   : تبدیل دسته‌ای فایل‌های نقشه‌ی SolidWorks (.SLDDRW) به DWG
  2) assembly  : باز کردن یک اسمبلی (.SLDASM)، پیمایش همه‌ی قطعات (حتی داخل
                 زیراسمبلی‌ها)، و Export گرفتن الگوی بازشده (Flat Pattern) هر
                 قطعه‌ی ورق‌فلزی به یک فایل DWG جدا (یک فایل به ازای هر Instance)

نیازمندی‌ها
-----------
    pip install pywin32 psutil
    - SolidWorks باید روی همین ویندوز نصب و لایسنس معتبر داشته باشد.

چرا Dispatch ساده به‌جای gencache.EnsureDispatch؟
--------------------------------------------------
    روی بعضی سیستم‌ها gencache.EnsureDispatch با خطای COM از نوع
    "Element not found" / "can not automate the makepy process" شکست می‌خورد،
    چون نمی‌تواند TypeInfo را از COM سالیدورکس بخواند. برای اینکه این اسکریپت
    مستقل از آن مشکل کار کند، از win32com.client.Dispatch (late-binding) و
    ثابت‌های API هاردکد (طبق مستندات رسمی، در کلاس SWConstants) استفاده می‌شود.

نحوه‌ی اجرا
-----------
    # ۱) تبدیل دسته‌ای همه‌ی فایل‌های .slddrw یک پوشه به DWG:
    python solidworks_export.py drawing --input "D:\Drawings" --output "D:\DWG_Out" --recursive

    # ۲) خروجی گرفتن الگوی بازشده‌ی قطعات ورق‌فلزی یک اسمبلی:
    python solidworks_export.py assembly --assembly "D:\Models\Rooftop 20T.SLDASM" --output "D:\DWG_Out"

    برای دیدن همه‌ی گزینه‌های هر زیر-دستور:
    python solidworks_export.py drawing --help
    python solidworks_export.py assembly --help

خروجی
------
    - فایل‌های DWG در پوشه‌ی خروجی
    - یک فایل لاگ کامل (conversion.log)
    - یک گزارش CSV نهایی (report.csv) شامل وضعیت هر فایل/قطعه
"""

from __future__ import annotations

import argparse
import csv
import logging
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    import psutil
except ImportError:
    psutil = None  # فقط برای kill کردن فرآیند گیرکرده لازم است؛ اختیاری

try:
    import win32com.client
except ImportError:
    print("پکیج pywin32 نصب نیست. اجرا کنید: pip install pywin32", file=sys.stderr)
    raise


# --------------------------------------------------------------------------- #
# ثابت‌های SolidWorks API (هاردکد، بدون نیاز به gencache)
# --------------------------------------------------------------------------- #

class SWConstants:
    swDocNONE = 0
    swDocPART = 1
    swDocASSEMBLY = 2
    swDocDRAWING = 3

    swOpenDocOptions_Silent = 1

    swSaveAsCurrentVersion = 0
    swSaveAsOptions_Silent = 1

    swExportToDWG_ExportSheetMetal = 1

    swSuppressFeature = 0
    swUnSuppressFeature = 1
    swThisConfiguration = 1


SW_PROCESS_NAME = "SLDWORKS.exe"
SHEET_METAL_FEATURE_TYPES = ("SheetMetal", "SMBaseFlange", "FlatPattern")
INVALID_FILENAME_CHARS = r'<>:"/\|?*'

logger = logging.getLogger("solidworks_export")


def setup_logging(log_path: Path, verbose: bool = False) -> None:
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if verbose else logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)


def sanitize_filename(name: str) -> str:
    for ch in INVALID_FILENAME_CHARS:
        name = name.replace(ch, "_")
    return name.strip()


# --------------------------------------------------------------------------- #
# نتیجه‌ی هر عملیات، برای گزارش نهایی (مشترک بین هر دو حالت)
# --------------------------------------------------------------------------- #

@dataclass
class ItemResult:
    name: str                 # نام فایل .slddrw یا نام کامپوننت اسمبلی
    source: str = ""          # مسیر فایل مبدا
    target: str = ""          # مسیر فایل DWG خروجی
    success: bool = False
    skipped_reason: str = ""  # مثلا NOT_SHEET_METAL, SKIPPED_EXISTS, MODEL_NOT_RESOLVED
    message: str = ""
    attempts: int = 1
    duration_sec: float = 0.0


@dataclass
class RunSummary:
    results: list = field(default_factory=list)

    def add(self, r: ItemResult) -> None:
        self.results.append(r)

    @property
    def ok(self):
        return [r for r in self.results if r.success]

    @property
    def skipped(self):
        return [r for r in self.results if r.skipped_reason and not r.success]

    @property
    def failed(self):
        return [r for r in self.results if not r.success and not r.skipped_reason]

    def write_csv(self, path: Path) -> None:
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["name", "source", "target", "success", "skipped_reason",
                        "attempts", "duration_sec", "message"])
            for r in self.results:
                w.writerow([r.name, r.source, r.target, r.success, r.skipped_reason,
                            r.attempts, f"{r.duration_sec:.2f}", r.message])


# --------------------------------------------------------------------------- #
# نشست مشترک SolidWorks (اتصال/قطع اتصال/راه‌اندازی مجدد)
# --------------------------------------------------------------------------- #

class SolidWorksSession:
    def __init__(self, visible: bool = False):
        self.visible = visible
        self.app = None
        self.const = SWConstants()

    def __enter__(self) -> "SolidWorksSession":
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()

    def connect(self) -> None:
        logger.info("در حال اتصال به SolidWorks (اگر باز نباشد، اجرا می‌شود؛ کمی صبر کنید)...")
        self.app = win32com.client.Dispatch("SldWorks.Application")
        self.app.Visible = self.visible
        logger.info("اتصال برقرار شد.")

    def disconnect(self) -> None:
        if self.app is not None:
            try:
                self.app.ExitApp()
            except Exception as e:
                logger.warning("بستن عادی SolidWorks ناموفق بود: %s", e)
            self.app = None
        self._kill_leftover()

    @staticmethod
    def _kill_leftover() -> None:
        if psutil is None:
            return
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == SW_PROCESS_NAME.lower():
                    p.kill()
                    logger.warning("پردازه‌ی باقیمانده‌ی %s به‌زور بسته شد.", SW_PROCESS_NAME)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

    def restart(self) -> None:
        logger.warning("راه‌اندازی مجدد SolidWorks...")
        self.disconnect()
        time.sleep(3)
        self.connect()


def run_with_timeout(worker_fn, timeout: float, on_timeout_result):
    """
    worker_fn را در یک ترد جدا اجرا می‌کند و اگر بیش از timeout طول بکشد،
    پردازه‌ی SolidWorks را به‌زور می‌بندد. برای جلوگیری از هنگ کامل روی
    فایل‌های خراب استفاده می‌شود.
    """
    box = {}

    def run():
        box["result"] = worker_fn()

    t = threading.Thread(target=run, daemon=True)
    t.start()
    t.join(timeout)

    if t.is_alive():
        logger.error("تایم‌اوت (%ss)؛ SolidWorks kill می‌شود.", timeout)
        SolidWorksSession._kill_leftover()
        return on_timeout_result

    return box.get("result", on_timeout_result)


# =========================================================================== #
# حالت ۱: تبدیل دسته‌ای SLDDRW -> DWG
# =========================================================================== #

def collect_drawing_files(input_dir: Path, recursive: bool) -> list:
    files = []
    it = input_dir.rglob("*") if recursive else input_dir.glob("*")
    for f in it:
        if f.suffix.lower() == ".slddrw":
            files.append(f)
    return sorted(set(files))


def drawing_target_path(src: Path, input_dir: Path, output_dir: Path) -> Path:
    rel = src.relative_to(input_dir).with_suffix(".dwg")
    return output_dir / rel


def convert_drawing_one(session: SolidWorksSession, src: Path, dst: Path) -> ItemResult:
    c = session.const
    app = session.app
    t0 = time.time()
    result = ItemResult(name=src.name, source=str(src), target=str(dst))

    dst.parent.mkdir(parents=True, exist_ok=True)
    model = None
    try:
        model = app.OpenDoc6(str(src), c.swDocDRAWING, c.swOpenDocOptions_Silent, "", 0, 0)
        if model is None:
            raise RuntimeError("SolidWorks نتوانست فایل را باز کند (فایل خراب یا نسخه ناسازگار؟).")

        try:
            model.ForceRebuild3(True)
        except Exception:
            pass

        ok = model.Extension.SaveAs3(
            str(dst), c.swSaveAsCurrentVersion, c.swSaveAsOptions_Silent, None, None, None
        )
        if not ok:
            raise RuntimeError(f"SaveAs3 شکست خورد (کد بازگشتی: {ok}).")
        if not dst.exists():
            raise RuntimeError("فایل DWG پس از SaveAs3 روی دیسک پیدا نشد.")

        result.success = True
        result.message = "OK"
        logger.info("✔ تبدیل شد: %s → %s", src.name, dst.name)

    except Exception as e:
        result.message = str(e)
        logger.error("✘ ناموفق: %s → %s | خطا: %s", src.name, dst.name, e)

    finally:
        if model is not None:
            try:
                app.CloseDoc(model.GetTitle())
            except Exception:
                pass
        result.duration_sec = time.time() - t0

    return result


def run_drawing_batch(args: argparse.Namespace) -> RunSummary:
    input_dir = Path(args.input).resolve()
    output_dir = Path(args.output).resolve()
    if not input_dir.exists():
        raise SystemExit(f"پوشه‌ی ورودی پیدا نشد: {input_dir}")

    files = collect_drawing_files(input_dir, args.recursive)
    logger.info("تعداد فایل .slddrw پیدا شده: %d", len(files))

    summary = RunSummary()
    if args.dry_run:
        for f in files:
            logger.info("[dry-run] %s -> %s", f, drawing_target_path(f, input_dir, output_dir))
        return summary
    if not files:
        logger.warning("هیچ فایل .slddrw پیدا نشد.")
        return summary

    with SolidWorksSession(visible=args.visible) as session:
        for idx, src in enumerate(files, start=1):
            dst = drawing_target_path(src, input_dir, output_dir)
            logger.info("[%d/%d] %s", idx, len(files), src.name)

            if dst.exists() and not args.overwrite:
                logger.info("رد شد (از قبل وجود دارد): %s", dst)
                summary.add(ItemResult(name=src.name, source=str(src), target=str(dst),
                                        success=True, skipped_reason="SKIPPED_EXISTS"))
                continue

            last = None
            for attempt in range(1, args.retries + 2):
                last = run_with_timeout(
                    lambda: convert_drawing_one(session, src, dst),
                    args.timeout,
                    ItemResult(name=src.name, source=str(src), target=str(dst), message="TIMEOUT"),
                )
                last.attempts = attempt
                if last.success:
                    break
                logger.warning("تلاش %d برای %s ناموفق بود؛ راه‌اندازی مجدد...", attempt, src.name)
                try:
                    session.restart()
                except Exception as e:
                    logger.error("راه‌اندازی مجدد ناموفق بود: %s", e)
                    break
            summary.add(last)

    return summary


# =========================================================================== #
# حالت ۲: Export الگوی بازشده‌ی قطعات ورق‌فلزی یک اسمبلی
# =========================================================================== #

def open_assembly(session: SolidWorksSession, path: Path):
    c = session.const
    model = session.app.OpenDoc6(str(path), c.swDocASSEMBLY, c.swOpenDocOptions_Silent, "", 0, 0)
    if model is None:
        raise RuntimeError(f"باز کردن اسمبلی ناموفق بود: {path}")
    try:
        model.Extension.ResolveAllLightWeightComponents(True)
    except Exception:
        pass
    try:
        model.ForceRebuild3(True)
    except Exception:
        pass
    return model


def close_assembly(session: SolidWorksSession, model) -> None:
    if model is not None:
        try:
            session.app.CloseDoc(model.GetTitle())
        except Exception:
            pass


def iter_leaf_components(assembly_model, include_suppressed: bool = False):
    root = assembly_model.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    if root is None:
        raise RuntimeError("ریشه‌ی اسمبلی پیدا نشد؛ فایل باز شده اسمبلی نیست؟")

    stack = list(root.GetChildren() or [])
    while stack:
        comp = stack.pop(0)
        try:
            is_suppressed = comp.IsSuppressed()
        except Exception:
            is_suppressed = False
        if is_suppressed and not include_suppressed:
            logger.debug("رد شد (Suppressed): %s", comp.Name2)
            continue

        children = comp.GetChildren() or []
        if children:
            stack.extend(children)
        else:
            yield comp


def is_sheet_metal(part_model) -> bool:
    try:
        feat = part_model.FirstFeature()
        while feat is not None:
            if feat.GetTypeName2() in SHEET_METAL_FEATURE_TYPES:
                return True
            feat = feat.GetNextFeature()
    except Exception as e:
        logger.debug("بررسی Sheet Metal با خطا مواجه شد: %s", e)
    return False


def find_flat_pattern_feature(part_model):
    try:
        feat = part_model.FirstFeature()
        while feat is not None:
            if feat.GetTypeName2() == "FlatPattern":
                return feat
            feat = feat.GetNextFeature()
    except Exception:
        pass
    return None


def export_component(session: SolidWorksSession, comp, output_dir: Path) -> ItemResult:
    c = session.const
    name = comp.Name2
    result = ItemResult(name=name)
    t0 = time.time()

    part_model = None
    flat_feat = None
    flat_was_suppressed = False

    try:
        part_model = comp.GetModelDoc2()
        if part_model is None:
            result.skipped_reason = "MODEL_NOT_RESOLVED"
            result.message = "مدل قطعه بارگذاری نشد (Lightweight/Missing Reference)."
            return result

        result.source = part_model.GetPathName()

        try:
            ref_config = comp.ReferencedConfiguration
            active_config = part_model.ConfigurationManager.ActiveConfiguration.Name
            if ref_config and ref_config != active_config:
                part_model.ShowConfiguration2(ref_config)
        except Exception as e:
            logger.debug("تنظیم Configuration برای %s ممکن نشد: %s", name, e)

        if not is_sheet_metal(part_model):
            result.skipped_reason = "NOT_SHEET_METAL"
            return result

        flat_feat = find_flat_pattern_feature(part_model)
        if flat_feat is not None:
            try:
                flat_was_suppressed = flat_feat.IsSuppressed()
                if flat_was_suppressed:
                    flat_feat.SetSuppression2(c.swUnSuppressFeature, c.swThisConfiguration, None)
                    part_model.EditRebuild3()
            except Exception as e:
                logger.debug("فعال‌سازی Flat Pattern برای %s ممکن نشد: %s", name, e)

        target = output_dir / f"{sanitize_filename(name)}.dwg"
        target.parent.mkdir(parents=True, exist_ok=True)
        result.target = str(target)

        try:
            part_model.ClearSelection2(True)
            part_model.Extension.SelectAll()
        except Exception:
            pass

        export_ok = part_model.ExportToDWG2(
            str(target), result.source, c.swExportToDWG_ExportSheetMetal,
            True, None, False, False, 0, None,
        )
        if not export_ok:
            raise RuntimeError("ExportToDWG2 مقدار False برگرداند.")
        if not target.exists():
            raise RuntimeError("فایل DWG بعد از Export روی دیسک پیدا نشد.")

        result.success = True
        result.message = "OK"
        logger.info("✔ خروجی گرفته شد: %s → %s", name, target.name)

    except Exception as e:
        result.message = str(e)
        logger.error("✘ ناموفق برای %s: %s", name, e)

    finally:
        if flat_feat is not None and flat_was_suppressed:
            try:
                flat_feat.SetSuppression2(c.swSuppressFeature, c.swThisConfiguration, None)
            except Exception:
                pass
        result.duration_sec = time.time() - t0

    return result


def run_assembly_batch(args: argparse.Namespace) -> RunSummary:
    assembly_path = Path(args.assembly).resolve()
    output_dir = Path(args.output).resolve()
    if not assembly_path.exists():
        raise SystemExit(f"فایل اسمبلی پیدا نشد: {assembly_path}")

    summary = RunSummary()

    with SolidWorksSession(visible=args.visible) as session:
        model = open_assembly(session, assembly_path)
        try:
            components = list(iter_leaf_components(model, include_suppressed=args.include_suppressed))
            logger.info("تعداد قطعات (Instance) پیدا شده: %d", len(components))

            if args.dry_run:
                for comp in components:
                    pm = comp.GetModelDoc2()
                    sm = is_sheet_metal(pm) if pm else False
                    logger.info("[dry-run] %s | sheet_metal=%s | resolved=%s", comp.Name2, sm, pm is not None)
                return summary

            for idx, comp in enumerate(components, start=1):
                logger.info("[%d/%d] %s", idx, len(components), comp.Name2)
                last = None
                for attempt in range(1, args.retries + 2):
                    last = run_with_timeout(
                        lambda: export_component(session, comp, output_dir),
                        args.timeout,
                        ItemResult(name=comp.Name2, message="TIMEOUT"),
                    )
                    last.attempts = attempt
                    if last.success or last.skipped_reason:
                        break
                    logger.warning("تلاش %d برای %s ناموفق بود؛ راه‌اندازی مجدد...", attempt, comp.Name2)
                    try:
                        session.restart()
                        model = open_assembly(session, assembly_path)
                    except Exception as e:
                        logger.error("راه‌اندازی مجدد ناموفق بود: %s", e)
                        break
                summary.add(last)
        finally:
            close_assembly(session, model)

    return summary


# =========================================================================== #
# CLI
# =========================================================================== #

def add_common_args(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("--output", "-o", required=True, help="پوشه‌ی خروجی برای فایل‌های DWG و گزارش‌ها")
    sp.add_argument("--visible", action="store_true", help="نمایش پنجره‌ی SolidWorks حین اجرا")
    sp.add_argument("--timeout", type=float, default=180.0, help="حداکثر زمان (ثانیه) مجاز برای هر فایل/قطعه")
    sp.add_argument("--retries", type=int, default=1, help="تعداد تلاش مجدد برای موارد ناموفق")
    sp.add_argument("--dry-run", action="store_true", help="فقط لیست کن، چیزی خروجی نده")
    sp.add_argument("--verbose", "-v", action="store_true", help="لاگ کامل‌تر در کنسول")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="ابزار یکپارچه‌ی Export از SolidWorks به DWG (نقشه‌ی تکی/دسته‌ای یا اسمبلی)."
    )
    sub = p.add_subparsers(dest="command", required=True)

    p_draw = sub.add_parser("drawing", help="تبدیل دسته‌ای فایل‌های .slddrw یک پوشه به DWG")
    p_draw.add_argument("--input", "-i", required=True, help="پوشه‌ی حاوی فایل‌های .slddrw")
    p_draw.add_argument("--recursive", "-r", action="store_true", help="پیمایش زیرپوشه‌ها هم انجام شود")
    p_draw.add_argument("--overwrite", action="store_true", help="فایل DWG موجود را بازنویسی کن")
    add_common_args(p_draw)

    p_asm = sub.add_parser("assembly", help="Export الگوی بازشده‌ی قطعات ورق‌فلزی یک اسمبلی")
    p_asm.add_argument("--assembly", "-a", required=True, help="مسیر فایل .SLDASM")
    p_asm.add_argument("--include-suppressed", action="store_true", help="قطعات Suppressed هم بررسی شوند")
    add_common_args(p_asm)

    return p.parse_args()


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    setup_logging(output_dir / "conversion.log", verbose=args.verbose)

    logger.info("=" * 70)
    logger.info("شروع اجرا | حالت: %s", args.command)
    logger.info("خروجی: %s", args.output)
    logger.info("=" * 70)

    start = time.time()
    try:
        if args.command == "drawing":
            summary = run_drawing_batch(args)
        else:
            summary = run_assembly_batch(args)
    except KeyboardInterrupt:
        logger.warning("توسط کاربر متوقف شد.")
        return 130

    if args.dry_run:
        return 0

    report_path = output_dir / "report.csv"
    summary.write_csv(report_path)

    elapsed = time.time() - start
    logger.info("=" * 70)
    logger.info(
        "پایان. موفق: %d | رد شده: %d | ناموفق: %d | زمان کل: %.1f ثانیه",
        len(summary.ok), len(summary.skipped), len(summary.failed), elapsed,
    )
    logger.info("گزارش کامل: %s", report_path)
    logger.info("=" * 70)

    return 0 if not summary.failed else 1


if __name__ == "__main__":
    sys.exit(main())