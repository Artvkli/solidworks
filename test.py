r"""
assembly_to_flat_dwg.py
========================

اتصال به SolidWorks، باز کردن یک اسمبلی (.SLDASM)، پیمایش همه‌ی قطعات آن (حتی داخل
زیراسمبلی‌ها)، و خروجی گرفتن الگوی باز شده (Flat Pattern) هر قطعه‌ی ورق‌فلزی
به فایل DWG جداگانه -- دقیقاً همان نوع خروجی که برای برش لیزر/پانچ لازم است
(مثل فایل نمونه‌ای که فرستادی: PartName_1.5mm-GAL.dwg).

منطق طبق توافق:
    - قطعاتی که Sheet Metal نیستند (پیچ، پروفیل ماشین‌کاری‌شده و ...) → نادیده گرفته می‌شوند.
    - قطعات تکراری (مثلاً ۴ پیچ/براکت یکسان) → به ازای هر Instance یک فایل جدا
      ساخته می‌شود (چون SolidWorks خودش به هر Instance نامی مثل "Bracket-1",
      "Bracket-2" می‌دهد، همین نام‌ها مستقیماً برای نام‌گذاری فایل خروجی استفاده می‌شود).

نیازمندی‌ها
-----------
    pip install pywin32 psutil
    - SolidWorks باید روی همین ویندوز نصب باشد.

اجرا
----
    python assembly_to_flat_dwg.py --assembly "D:\Models\AHU22000.SLDASM" --output "D:\DWG_Out"

    گزینه‌های مهم:
      --skip-suppressed / --include-suppressed   قطعات Suppressed نادیده گرفته شوند یا نه (پیش‌فرض: نادیده)
      --visible            نمایش پنجره‌ی SolidWorks حین اجرا (دیباگ)
      --timeout SEC        حداکثر زمان مجاز برای هر قطعه قبل از kill کردن SolidWorks
      --retries N          تعداد تلاش مجدد برای قطعات ناموفق
      --dry-run            فقط لیست قطعات را نشان بده (سالید فلزی/غیرفلزی)، چیزی خروجی نده

خروجی
------
    - یک فایل DWG برای هر Instance از هر قطعه‌ی ورق‌فلزی، در پوشه‌ی خروجی
    - assembly_conversion.log و assembly_report.csv در همان پوشه
"""

from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    import psutil
except ImportError:
    psutil = None

try:
    import win32com.client
    from win32com.client import gencache
except ImportError:
    print("پکیج pywin32 نصب نیست. اجرا کنید: pip install pywin32", file=sys.stderr)
    raise


SW_PROCESS_NAME = "SLDWORKS.exe"
SHEET_METAL_FEATURE_TYPES = ("SheetMetal", "SMBaseFlange", "FlatPattern")
INVALID_FILENAME_CHARS = r'<>:"/\|?*'

logger = logging.getLogger("assembly2flatdwg")


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


def resolve_const(constants_module, candidate_names: list, default: Optional[int] = None) -> Optional[int]:
    for name in candidate_names:
        if hasattr(constants_module, name):
            return getattr(constants_module, name)
    if default is not None:
        logger.debug("ثابت‌های %s پیدا نشد؛ از مقدار پیش‌فرض %s استفاده می‌شود.", candidate_names, default)
    return default


@dataclass
class PartResult:
    component: str
    part_file: str = ""
    target: str = ""
    sheet_metal: bool = False
    exported: bool = False
    skipped_reason: str = ""
    message: str = ""
    attempts: int = 1
    duration_sec: float = 0.0


@dataclass
class RunSummary:
    results: list = field(default_factory=list)

    def add(self, r: PartResult) -> None:
        self.results.append(r)

    @property
    def exported(self):
        return [r for r in self.results if r.exported]

    @property
    def failed(self):
        return [r for r in self.results if not r.exported and not r.skipped_reason]

    @property
    def skipped(self):
        return [r for r in self.results if r.skipped_reason]

    def write_csv(self, path: Path) -> None:
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["component", "part_file", "target_dwg", "sheet_metal",
                        "exported", "skipped_reason", "attempts", "duration_sec", "message"])
            for r in self.results:
                w.writerow([r.component, r.part_file, r.target, r.sheet_metal, r.exported,
                            r.skipped_reason, r.attempts, f"{r.duration_sec:.2f}", r.message])


# --------------------------------------------------------------------------- #
# نشست SolidWorks
# --------------------------------------------------------------------------- #

class SolidWorksAssemblySession:
    def __init__(self, visible: bool = False):
        self.visible = visible
        self.app = None
        self.const = None
        self.assembly_model = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()

    def connect(self) -> None:
        logger.info("در حال اتصال به SolidWorks...")
        self.app = gencache.EnsureDispatch("SldWorks.Application")
        self.const = win32com.client.constants
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

    # --- باز کردن اسمبلی ---------------------------------------------------- #

    def open_assembly(self, path: Path):
        c = self.const
        doc_type = resolve_const(c, ["swDocASSEMBLY"], default=2)
        open_opts = resolve_const(c, ["swOpenDocOptions_Silent"], default=1)
        model = self.app.OpenDoc6(str(path), doc_type, open_opts, "", 0, 0)
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
        self.assembly_model = model
        return model

    def close_assembly(self) -> None:
        if self.assembly_model is not None:
            try:
                self.app.CloseDoc(self.assembly_model.GetTitle())
            except Exception:
                pass
            self.assembly_model = None

    # --- پیمایش قطعات (Leaf Components) -------------------------------------- #

    def iter_leaf_components(self, include_suppressed: bool = False):
        root = self.assembly_model.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
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
                stack.extend(children)  # زیراسمبلی → برو داخل‌تر
            else:
                yield comp

    # --- بررسی ورق‌فلزی بودن قطعه --------------------------------------------- #

    @staticmethod
    def is_sheet_metal(part_model) -> bool:
        try:
            feat = part_model.FirstFeature()
            while feat is not None:
                type_name = feat.GetTypeName2()
                if type_name in SHEET_METAL_FEATURE_TYPES:
                    return True
                feat = feat.GetNextFeature()
        except Exception as e:
            logger.debug("بررسی Sheet Metal با خطا مواجه شد: %s", e)
        return False

    @staticmethod
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

    # --- خروجی گرفتن یک قطعه ------------------------------------------------- #

    def export_component(self, comp, output_dir: Path) -> PartResult:
        c = self.const
        name = comp.Name2  # مثلا "Panel_Side-1" ؛ عدد انتهایی مشخصه‌ی Instance است
        result = PartResult(component=name)
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

            result.part_file = part_model.GetPathName()

            # اگر Instance از یک Configuration خاص استفاده می‌کند، همان را فعال کن
            try:
                ref_config = comp.ReferencedConfiguration
                active_config = part_model.ConfigurationManager.ActiveConfiguration.Name
                if ref_config and ref_config != active_config:
                    part_model.ShowConfiguration2(ref_config)
            except Exception as e:
                logger.debug("تنظیم Configuration برای %s ممکن نشد: %s", name, e)

            if not self.is_sheet_metal(part_model):
                result.sheet_metal = False
                result.skipped_reason = "NOT_SHEET_METAL"
                return result

            result.sheet_metal = True

            # الگوی بازشده (Flat Pattern) معمولاً به‌صورت پیش‌فرض Suppressed است؛
            # موقتاً آن را فعال می‌کنیم، بعد از Export به حالت اول برمی‌گردانیم.
            flat_feat = self.find_flat_pattern_feature(part_model)
            if flat_feat is not None:
                try:
                    flat_was_suppressed = flat_feat.IsSuppressed()
                    if flat_was_suppressed:
                        unsuppress = resolve_const(c, ["swUnSuppressFeature"], default=0)
                        flat_feat.SetSuppression2(unsuppress, 1, None)
                        part_model.EditRebuild3()
                except Exception as e:
                    logger.debug("فعال‌سازی Flat Pattern برای %s ممکن نشد: %s", name, e)

            target = output_dir / f"{sanitize_filename(name)}.dwg"
            target.parent.mkdir(parents=True, exist_ok=True)
            result.target = str(target)

            action = resolve_const(c, ["swExportToDWG_ExportSheetMetal"], default=1)

            # برای قطعات چندبدنه (Multi-body) گاهی لازم است بدنه‌ها انتخاب شده باشند
            try:
                part_model.ClearSelection2(True)
                part_model.Extension.SelectAll()
            except Exception:
                pass

            export_ok = part_model.ExportToDWG2(
                str(target),          # FilePath
                result.part_file,     # ModelName
                action,                # Action = Export Sheet Metal
                True,                  # ExportToSingleFile
                None,                  # Alignment
                False,                 # IsXDirFlipped
                False,                 # IsYDirFlipped
                0,                     # SheetMetalOptions
                None,                  # Views
            )

            if not export_ok:
                raise RuntimeError("ExportToDWG2 مقدار False برگرداند.")
            if not target.exists():
                raise RuntimeError("فایل DWG بعد از Export روی دیسک پیدا نشد.")

            result.exported = True
            result.message = "OK"
            logger.info("✔ خروجی گرفته شد: %s → %s", name, target.name)

        except Exception as e:
            result.message = str(e)
            logger.error("✘ ناموفق برای %s: %s", name, e)

        finally:
            # اگر خودمان Flat Pattern را فعال کردیم، به حالت اولش برگردانیم
            if flat_feat is not None and flat_was_suppressed:
                try:
                    suppress = resolve_const(c, ["swSuppressFeature"], default=1)
                    flat_feat.SetSuppression2(suppress, 1, None)
                except Exception:
                    pass
            result.duration_sec = time.time() - t0

        return result


# --------------------------------------------------------------------------- #
# اجرای دسته‌ای با timeout / retry (مشابه اسکریپت قبلی)
# --------------------------------------------------------------------------- #

def convert_with_timeout(session: SolidWorksAssemblySession, comp, output_dir: Path, timeout: float) -> PartResult:
    import threading

    box = {}

    def worker():
        box["result"] = session.export_component(comp, output_dir)

    name = comp.Name2
    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(timeout)

    if t.is_alive():
        logger.error("تایم‌اوت (%ss) برای %s؛ SolidWorks kill می‌شود.", timeout, name)
        SolidWorksAssemblySession._kill_leftover()
        return PartResult(component=name, message=f"TIMEOUT after {timeout}s")

    return box.get("result", PartResult(component=name, message="UNKNOWN_ERROR"))


def run(args: argparse.Namespace) -> RunSummary:
    assembly_path = Path(args.assembly).resolve()
    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not assembly_path.exists():
        raise SystemExit(f"فایل اسمبلی پیدا نشد: {assembly_path}")

    summary = RunSummary()

    with SolidWorksAssemblySession(visible=args.visible) as session:
        session.open_assembly(assembly_path)
        try:
            components = list(session.iter_leaf_components(include_suppressed=args.include_suppressed))
            logger.info("تعداد قطعات (Instance) پیدا شده: %d", len(components))

            if args.dry_run:
                for comp in components:
                    part_model = comp.GetModelDoc2()
                    is_sm = session.is_sheet_metal(part_model) if part_model else False
                    logger.info("[dry-run] %s | sheet_metal=%s | resolved=%s",
                                comp.Name2, is_sm, part_model is not None)
                return summary

            for idx, comp in enumerate(components, start=1):
                logger.info("[%d/%d] %s", idx, len(components), comp.Name2)
                last = None
                for attempt in range(1, args.retries + 2):
                    last = convert_with_timeout(session, comp, output_dir, args.timeout)
                    last.attempts = attempt
                    if last.exported or last.skipped_reason:
                        break
                    logger.warning("تلاش %d برای %s ناموفق بود؛ راه‌اندازی مجدد SolidWorks...", attempt, comp.Name2)
                    try:
                        session.restart()
                        session.open_assembly(assembly_path)
                    except Exception as e:
                        logger.error("راه‌اندازی مجدد ناموفق بود: %s", e)
                        break
                summary.add(last)
        finally:
            session.close_assembly()

    return summary


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="خروجی گرفتن الگوی بازشده‌ی (Flat Pattern) قطعات ورق‌فلزی یک اسمبلی SolidWorks به DWG."
    )
    p.add_argument("--assembly", "-a", required=True, help="مسیر فایل .SLDASM")
    p.add_argument("--output", "-o", required=True, help="پوشه‌ی خروجی برای فایل‌های DWG")
    p.add_argument("--include-suppressed", action="store_true", help="قطعات Suppressed هم بررسی شوند")
    p.add_argument("--visible", action="store_true", help="نمایش پنجره‌ی SolidWorks حین اجرا")
    p.add_argument("--timeout", type=float, default=180.0, help="حداکثر زمان (ثانیه) مجاز برای هر قطعه")
    p.add_argument("--retries", type=int, default=1, help="تعداد تلاش مجدد برای قطعات ناموفق")
    p.add_argument("--dry-run", action="store_true", help="فقط لیست قطعات و نوعشان را نشان بده")
    p.add_argument("--verbose", "-v", action="store_true", help="لاگ کامل‌تر در کنسول")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    setup_logging(output_dir / "assembly_conversion.log", verbose=args.verbose)

    logger.info("=" * 70)
    logger.info("شروع خروجی‌گیری Flat Pattern از اسمبلی")
    logger.info("اسمبلی : %s", args.assembly)
    logger.info("خروجی  : %s", args.output)
    logger.info("=" * 70)

    start = time.time()
    try:
        summary = run(args)
    except KeyboardInterrupt:
        logger.warning("توسط کاربر متوقف شد.")
        return 130

    if args.dry_run:
        return 0

    report_path = output_dir / "assembly_report.csv"
    summary.write_csv(report_path)

    elapsed = time.time() - start
    logger.info("=" * 70)
    logger.info(
        "پایان. خروجی گرفته‌شده: %d | رد شده (غیر ورق‌فلزی/...): %d | ناموفق: %d | زمان کل: %.1f ثانیه",
        len(summary.exported), len(summary.skipped), len(summary.failed), elapsed,
    )
    logger.info("گزارش کامل: %s", report_path)
    logger.info("=" * 70)

    return 0 if not summary.failed else 1


if __name__ == "__main__":
    sys.exit(main())