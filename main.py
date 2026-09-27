r"""
assembly_to_flat_dwg.py
========================

اتصال به SolidWorks، باز کردن یک اسمبلی (.SLDASM)، پیمایش همه‌ی قطعات آن (حتی داخل
زیراسمبلی‌ها)، خروجی گرفتن الگوی باز شده (Flat Pattern) هر قطعه‌ی ورق‌فلزی، و در
نهایت **گروه‌بندی بر اساس ضخامت ورق** و ترکیب همه‌ی قطعات هم‌ضخامت در یک فایل DXF
واحد -- یعنی برای هر اسمبلی، به تعداد ضخامت‌های متفاوت موجود در آن (مثلاً 1mm،
1.5mm، 2mm و ...) یک فایل DXF جداگانه ساخته می‌شود که همه‌ی Flat Pattern های همان
ضخامت داخلش کنار هم چیده شده‌اند (دقیقاً همان چیزی که برای نصب روی دستگاه برش
لیزر/پانچ لازم است).

منطق طبق توافق:
    - قطعاتی که Sheet Metal نیستند (پیچ، پروفیل ماشین‌کاری‌شده و ...) → نادیده گرفته می‌شوند.
    - قطعات تکراری (مثلاً ۴ پیچ/براکت یکسان) → به ازای هر Instance یک Flat Pattern
      جدا گرفته می‌شود (چون SolidWorks خودش به هر Instance نامی مثل "Bracket-1",
      "Bracket-2" می‌دهد) و همه‌ی این‌ها اگر هم‌ضخامت باشند، در یک فایل جمع می‌شوند.
    - ابتدا Flat Pattern هر قطعه در یک فایل DXF موقت (temp) گرفته می‌شود، سپس
      قطعات بر اساس ضخامت (گرد شده تا --group-decimals رقم اعشار) دسته‌بندی و با
      کتابخانه‌ی ezdxf در یک فایل DXF نهایی برای هر گروهِ ضخامت ادغام می‌شوند.
      فایل‌های موقت در پایان پاک می‌شوند (مگر با --keep-temp).

نیازمندی‌ها
-----------
    pip install pywin32 psutil ezdxf
    - SolidWorks باید روی همین ویندوز نصب باشد.
    - ادغام فایل‌ها فقط برای خروجی DXF پشتیبانی می‌شود (نه DWG باینری).

اجرا
----
    python assembly_to_flat_dwg.py --assembly "D:\Models\AHU22000.SLDASM" --output "D:\DWG_Out"

    گزینه‌های مهم:
      --skip-suppressed / --include-suppressed   قطعات Suppressed نادیده گرفته شوند یا نه (پیش‌فرض: نادیده)
      --visible            نمایش پنجره‌ی SolidWorks حین اجرا (دیباگ)
      --timeout SEC        حداکثر زمان مجاز برای هر قطعه قبل از kill کردن SolidWorks
      --retries N          تعداد تلاش مجدد برای قطعات ناموفق
      --dry-run            فقط لیست قطعات و ضخامتشان را نشان بده، چیزی خروجی نده
      --group-decimals N   تعداد رقم اعشار برای گرد کردن ضخامت هنگام گروه‌بندی (پیش‌فرض: 2)
      --gap MM             فاصله‌ی بین قطعات داخل فایل ادغام‌شده، بر حسب میلی‌متر (پیش‌فرض: 20)
      --keep-temp          فایل‌های DXF موقت هر قطعه پاک نشوند (برای دیباگ)
      --no-group           برگشت به رفتار قدیمی: یک فایل جدا به ازای هر قطعه (بدون گروه‌بندی ضخامت)
      --with-bend-lines    خطوط خم (Bend Lines) هم در خروجی گنجانده شود (پیش‌فرض: فقط هندسه)

خروجی
------
    - در حالت گروه‌بندی (پیش‌فرض): یک فایل DXF برای هر ضخامت موجود در اسمبلی،
      با نام "{نام اسمبلی}_{ضخامت}mm.dxf"، شامل Flat Pattern همه‌ی قطعات هم‌ضخامت.
    - در حالت --no-group: یک فایل جدا برای هر Instance از هر قطعه‌ی ورق‌فلزی.
    - assembly_conversion.log و assembly_report.csv در همان پوشه‌ی خروجی
"""

from __future__ import annotations

import argparse
import csv
import logging
import shutil
import sys
import tempfile
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

try:
    import ezdxf
    from ezdxf import bbox as ezdxf_bbox
    from ezdxf.addons.importer import Importer as EzdxfImporter
except ImportError:
    ezdxf = None


SW_PROCESS_NAME = "SLDWORKS.exe"

# برای تشخیص «ورق‌فلزی بودن قطعه» - وجود هر یک از این تایپ‌های Feature کافی است.
SHEET_METAL_FEATURE_TYPES = ("SheetMetal", "SMBaseFlange", "FlatPattern")

# برای خواندن مقدار ضخامت (Thickness) - این‌ها همان Feature پایه‌ی Sheet Metal هستند
# (نه خودِ FlatPattern، چون تعریف FlatPattern مقدار Thickness ندارد).
THICKNESS_FEATURE_TYPES = ("SheetMetal", "SMBaseFlange", "SM3D", "BaseFlange")

INVALID_FILENAME_CHARS = r'<>:"/\|?*'

DEFAULT_GROUP_DECIMALS = 2
DEFAULT_GAP_MM = 20.0

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


def format_thickness_label(thickness_mm: float, decimals: int) -> str:
    """0.8 -> '0.8' | 1.0 -> '1' | 1.50 -> '1.5' (بدون صفرهای اضافه در انتها)."""
    rounded = round(thickness_mm, decimals)
    text = f"{rounded:.{decimals}f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


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
    thickness_mm: Optional[float] = None
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
            w.writerow(["component", "part_file", "thickness_mm", "target_file", "sheet_metal",
                        "exported", "skipped_reason", "attempts", "duration_sec", "message"])
            for r in self.results:
                thickness_text = "" if r.thickness_mm is None else f"{r.thickness_mm:.3f}"
                w.writerow([r.component, r.part_file, thickness_text, r.target, r.sheet_metal, r.exported,
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

    @staticmethod
    def get_thickness_mm(part_model) -> Optional[float]:
        """ضخامت ورق را بر حسب میلی‌متر برمی‌گرداند؛ اگر پیدا نشد None."""
        try:
            feat = part_model.FirstFeature()
            while feat is not None:
                type_name = feat.GetTypeName2()
                if type_name in THICKNESS_FEATURE_TYPES:
                    try:
                        definition = feat.GetDefinition()
                        thickness_m = definition.Thickness
                        if thickness_m:
                            return float(thickness_m) * 1000.0
                    except Exception as e:
                        logger.debug("خواندن Thickness از Feature '%s' ممکن نشد: %s", type_name, e)
                feat = feat.GetNextFeature()
        except Exception as e:
            logger.debug("پیمایش Feature برای پیدا کردن ضخامت با خطا مواجه شد: %s", e)
        return None

    # --- خروجی گرفتن یک قطعه ------------------------------------------------- #

    def export_component(
        self,
        comp,
        target_path: Path,
        sheet_metal_options: int = 0,
    ) -> PartResult:
        """Flat Pattern یک Instance را به یک فایل DXF/DWG مستقل (target_path) خروجی می‌گیرد.

        برای گروه‌بندی بر اساس ضخامت، target_path معمولاً یک مسیر موقت است که
        بعداً توسط merge_dxf_group با بقیه‌ی قطعات هم‌ضخامت ادغام می‌شود.
        """
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

            # خواندن ضخامت ورق (برای گروه‌بندی بعدی لازم است)
            result.thickness_mm = self.get_thickness_mm(part_model)

            if result.thickness_mm is None:
                result.skipped_reason = "THICKNESS_NOT_FOUND"
                result.message = "ضخامت ورق پیدا نشد؛ قطعه در هیچ گروهی قرار نمی‌گیرد."
                logger.warning("ضخامت پیدا نشد برای %s؛ رد می‌شود.", name)
                return result

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

            target_path.parent.mkdir(parents=True, exist_ok=True)
            result.target = str(target_path)

            action = resolve_const(c, ["swExportToDWG_ExportSheetMetal"], default=1)

            # برای قطعات چندبدنه (Multi-body) گاهی لازم است بدنه‌ها انتخاب شده باشند
            try:
                part_model.ClearSelection2(True)
                part_model.Extension.SelectAll()
            except Exception:
                pass

            export_ok = part_model.ExportToDWG2(
                str(target_path),      # FilePath
                result.part_file,      # ModelName
                action,                 # Action = Export Sheet Metal
                True,                   # ExportToSingleFile
                None,                   # Alignment
                False,                  # IsXDirFlipped
                False,                  # IsYDirFlipped
                sheet_metal_options,    # SheetMetalOptions (bit1=geometry, bit3=bend lines)
                None,                   # Views
            )

            if not export_ok:
                raise RuntimeError("ExportToDWG2 مقدار False برگرداند.")
            if not target_path.exists():
                raise RuntimeError("فایل خروجی بعد از Export روی دیسک پیدا نشد.")

            result.exported = True
            result.message = "OK"
            logger.info(
                "✔ Flat Pattern گرفته شد: %s (ضخامت %.3f mm) → %s",
                name, result.thickness_mm, target_path.name,
            )

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
# ادغام چند فایل DXF (هم‌ضخامت) در یک فایل واحد
# --------------------------------------------------------------------------- #

def merge_dxf_group(
    parts: list[tuple[str, Path]],
    output_path: Path,
    gap_mm: float = DEFAULT_GAP_MM,
    add_labels: bool = True,
) -> None:
    """چند فایل DXF (خروجی Flat Pattern هر قطعه) را در یک فایل DXF واحد ادغام می‌کند.

    هر قطعه به‌صورت یک Block مستقل import می‌شود و سپس در یک چیدمان شبکه‌ای
    (grid) کنار هم، بدون هم‌پوشانی، در فایل نهایی درج می‌شود. اگر add_labels
    فعال باشد، نام هر Instance زیر آن نوشته می‌شود.

    parts: لیستی از (برچسب/نام قطعه، مسیر فایل DXF موقت آن قطعه)
    """
    if ezdxf is None:
        raise RuntimeError(
            "کتابخانه‌ی ezdxf نصب نیست. اجرا کنید: pip install ezdxf"
        )

    if not parts:
        raise ValueError("لیست قطعات برای ادغام خالی است.")

    target_doc = ezdxf.new(dxfversion="R2010")
    target_msp = target_doc.modelspace()

    max_row_width_mm = 3000.0  # بعد از این عرض، سطر جدید شروع می‌شود
    cursor_x = 0.0
    cursor_y = 0.0
    row_height = 0.0

    for index, (label, dxf_path) in enumerate(parts):
        try:
            source_doc = ezdxf.readfile(str(dxf_path))
        except Exception as e:
            logger.warning("خواندن فایل موقت DXF ممکن نشد (%s): %s", dxf_path, e)
            continue

        source_msp = source_doc.modelspace()

        extents = ezdxf_bbox.extents(source_msp, fast=True)
        if extents.has_data:
            width = extents.size.x
            height = extents.size.y
            min_x, min_y = extents.extmin.x, extents.extmin.y
        else:
            width = height = min_x = min_y = 0.0

        # اگر با اضافه شدن این قطعه از عرض مجاز سطر رد شویم، برو به سطر بعد
        if cursor_x > 0 and (cursor_x + width) > max_row_width_mm:
            cursor_x = 0.0
            cursor_y += row_height + gap_mm
            row_height = 0.0

        offset_x = cursor_x - min_x
        offset_y = cursor_y - min_y

        block_name = f"PART_{index:04d}"
        block = target_doc.blocks.new(name=block_name)

        importer = EzdxfImporter(source_doc, target_doc)
        importer.import_modelspace(target_layout=block)
        importer.finalize()

        target_msp.add_blockref(block_name, insert=(offset_x, offset_y))

        if add_labels:
            label_height = max(3.0, min(15.0, height * 0.06 if height else 5.0))
            target_msp.add_text(
                label,
                dxfattribs={"height": label_height, "layer": "PART_LABELS"},
            ).set_placement((offset_x, offset_y - label_height * 2.0))

        cursor_x += width + gap_mm
        row_height = max(row_height, height)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    target_doc.saveas(str(output_path))


# --------------------------------------------------------------------------- #
# اجرای دسته‌ای با timeout / retry (مشابه اسکریپت قبلی)
# --------------------------------------------------------------------------- #

def convert_with_timeout(
    session: SolidWorksAssemblySession,
    comp,
    target_path: Path,
    sheet_metal_options: int,
    timeout: float,
) -> PartResult:
    import threading

    box = {}

    def worker():
        box["result"] = session.export_component(comp, target_path, sheet_metal_options)

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

    group_by_thickness = not args.no_group
    sheet_metal_options = 5 if args.with_bend_lines else 1

    if group_by_thickness and ezdxf is None:
        raise SystemExit(
            "برای گروه‌بندی بر اساس ضخامت باید ezdxf نصب باشد: pip install ezdxf\n"
            "یا برای بازگشت به حالت قدیمی (یک فایل به ازای هر قطعه) از --no-group استفاده کنید."
        )

    if not assembly_path.exists():
        raise SystemExit(f"فایل اسمبلی پیدا نشد: {assembly_path}")

    summary = RunSummary()

    # در حالت گروه‌بندی، ابتدا هر قطعه در یک پوشه‌ی موقت DXP می‌گیرد؛
    # در حالت --no-group مستقیماً در output_dir خروجی گرفته می‌شود.
    temp_dir: Optional[Path] = None
    if group_by_thickness:
        temp_dir = Path(tempfile.mkdtemp(prefix="flat_pattern_tmp_"))

    try:
        with SolidWorksAssemblySession(visible=args.visible) as session:
            session.open_assembly(assembly_path)
            try:
                components = list(session.iter_leaf_components(include_suppressed=args.include_suppressed))
                logger.info("تعداد قطعات (Instance) پیدا شده: %d", len(components))

                if args.dry_run:
                    for comp in components:
                        part_model = comp.GetModelDoc2()
                        is_sm = session.is_sheet_metal(part_model) if part_model else False
                        thickness = (
                            session.get_thickness_mm(part_model)
                            if (part_model is not None and is_sm)
                            else None
                        )
                        thickness_text = f"{thickness:.3f} mm" if thickness is not None else "-"
                        logger.info(
                            "[dry-run] %s | sheet_metal=%s | thickness=%s | resolved=%s",
                            comp.Name2, is_sm, thickness_text, part_model is not None,
                        )
                    return summary

                # ------------------------------------------------------- #
                # فاز ۱: خروجی گرفتن Flat Pattern هر قطعه (تکی)
                # ------------------------------------------------------- #

                for idx, comp in enumerate(components, start=1):
                    name = comp.Name2
                    logger.info("[%d/%d] %s", idx, len(components), name)

                    if group_by_thickness:
                        target_path = temp_dir / f"{sanitize_filename(name)}_{idx:04d}.dxf"
                    else:
                        target_path = output_dir / f"{sanitize_filename(name)}.dxf"

                    last = None
                    for attempt in range(1, args.retries + 2):
                        last = convert_with_timeout(
                            session, comp, target_path, sheet_metal_options, args.timeout
                        )
                        last.attempts = attempt
                        if last.exported or last.skipped_reason:
                            break
                        logger.warning(
                            "تلاش %d برای %s ناموفق بود؛ راه‌اندازی مجدد SolidWorks...",
                            attempt, name,
                        )
                        try:
                            session.restart()
                            session.open_assembly(assembly_path)
                        except Exception as e:
                            logger.error("راه‌اندازی مجدد ناموفق بود: %s", e)
                            break
                    summary.add(last)
            finally:
                session.close_assembly()

        # ------------------------------------------------------------- #
        # فاز ۲: گروه‌بندی بر اساس ضخامت و ادغام در یک فایل برای هر گروه
        # ------------------------------------------------------------- #

        if group_by_thickness:
            groups: dict[float, list[tuple[PartResult, Path]]] = {}

            for r in summary.results:
                if not r.exported or r.thickness_mm is None:
                    continue
                key = round(r.thickness_mm, args.group_decimals)
                groups.setdefault(key, []).append((r, Path(r.target)))

            logger.info("تعداد گروه‌های ضخامت پیدا شده: %d", len(groups))

            for thickness_key in sorted(groups.keys()):
                entries = groups[thickness_key]
                label = format_thickness_label(thickness_key, args.group_decimals)
                group_output = output_dir / f"{assembly_path.stem}_{label}mm.dxf"

                parts_for_merge = [(r.component, path) for r, path in entries]

                try:
                    merge_dxf_group(parts_for_merge, group_output, gap_mm=args.gap)
                    for r, _ in entries:
                        r.target = str(group_output)
                        r.message = f"ادغام شد در گروه ضخامت {label}mm ({len(entries)} قطعه)."
                    logger.info(
                        "✔ گروه ضخامت %smm → %s (%d قطعه)",
                        label, group_output.name, len(entries),
                    )
                except Exception as e:
                    logger.error("✘ ادغام گروه ضخامت %smm ناموفق بود: %s", label, e)
                    for r, _ in entries:
                        r.exported = False
                        r.message = f"ادغام گروه ناموفق بود: {e}"

    finally:
        if temp_dir is not None and not args.keep_temp:
            shutil.rmtree(temp_dir, ignore_errors=True)
        elif temp_dir is not None:
            logger.info("فایل‌های موقت نگه داشته شدند در: %s", temp_dir)

    return summary


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=(
            "خروجی گرفتن الگوی بازشده‌ی (Flat Pattern) قطعات ورق‌فلزی یک اسمبلی SolidWorks، "
            "با گروه‌بندی بر اساس ضخامت در یک فایل DXF مشترک برای هر ضخامت."
        )
    )
    p.add_argument("--assembly", "-a", required=True, help="مسیر فایل .SLDASM")
    p.add_argument("--output", "-o", required=True, help="پوشه‌ی خروجی برای فایل‌های DXF")
    p.add_argument("--include-suppressed", action="store_true", help="قطعات Suppressed هم بررسی شوند")
    p.add_argument("--visible", action="store_true", help="نمایش پنجره‌ی SolidWorks حین اجرا")
    p.add_argument("--timeout", type=float, default=180.0, help="حداکثر زمان (ثانیه) مجاز برای هر قطعه")
    p.add_argument("--retries", type=int, default=1, help="تعداد تلاش مجدد برای قطعات ناموفق")
    p.add_argument("--dry-run", action="store_true", help="فقط لیست قطعات، نوع و ضخامتشان را نشان بده")
    p.add_argument("--verbose", "-v", action="store_true", help="لاگ کامل‌تر در کنسول")

    p.add_argument(
        "--no-group", action="store_true",
        help="گروه‌بندی بر اساس ضخامت غیرفعال شود؛ یک فایل جدا برای هر قطعه ساخته شود (رفتار قدیمی).",
    )
    p.add_argument(
        "--group-decimals", type=int, default=DEFAULT_GROUP_DECIMALS,
        help="تعداد رقم اعشار برای گرد کردن ضخامت هنگام گروه‌بندی (پیش‌فرض: %(default)s)",
    )
    p.add_argument(
        "--gap", type=float, default=DEFAULT_GAP_MM,
        help="فاصله‌ی بین قطعات داخل فایل ادغام‌شده، بر حسب میلی‌متر (پیش‌فرض: %(default)s)",
    )
    p.add_argument(
        "--keep-temp", action="store_true",
        help="فایل‌های DXF موقت هر قطعه (پیش از ادغام) پاک نشوند؛ برای دیباگ مفید است.",
    )
    p.add_argument(
        "--with-bend-lines", action="store_true",
        help="خطوط خم (Bend Lines) هم در خروجی گنجانده شود (پیش‌فرض: فقط هندسه‌ی بیرونی).",
    )

    return p.parse_args()


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    setup_logging(output_dir / "assembly_conversion.log", verbose=args.verbose)

    logger.info("=" * 70)
    logger.info("شروع خروجی‌گیری Flat Pattern از اسمبلی")
    logger.info("اسمبلی      : %s", args.assembly)
    logger.info("خروجی       : %s", args.output)
    logger.info(
        "حالت خروجی  : %s",
        "یک فایل DXF جدا برای هر ضخامت (--no-group غیرفعال است)"
        if not args.no_group else
        "یک فایل جدا برای هر قطعه (--no-group فعال است)",
    )
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