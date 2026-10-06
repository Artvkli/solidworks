from pathlib import Path
import gc
import json
import re
import sys
import traceback

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is required. Install: pip install pywin32")
    raise

from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.dwg_exporter import DwgExporter
from solidworks.pdf_drawings import PdfDrawingExporter
from solidworks.thickness_dwg import ThicknessDwgBuilder
from solidworks.sheet_metal import SheetMetalDetector


GROUP_BY_THICKNESS = True
REPEAT_BY_QUANTITY = False
ODA_CONVERTER_PATH = None
PART_GAP = 20.0
MAX_ROW_WIDTH = 6000.0


# ============================================================
# PDF SETTINGS
# ============================================================

EXPORT_PDF = True

# False = اجرای عادی کل برنامه
# True  = فقط PDF از DXFهای قبلی ساخته می‌شود
PDF_ONLY = True

PDF_PAPER = "A4"
PDF_ORIENTATION = "portrait"
PDF_COMPANY = "TECHNICAL DEPARTMENT"
PDF_DRAWN_BY = "HE.A"
PDF_CHECKED_BY = "R.GH"
PDF_APPROVED_BY = "R.GH"
PDF_MATERIAL = ""


def connect_to_solidworks():
    print("\nConnecting to SolidWorks...")
    try:
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
        print("Connected successfully.")
        return sw_app
    except Exception as exc:
        print(f"Could not connect to SolidWorks: {exc}")
        return None


def select_assembly(input_folder):
    assemblies = find_assemblies(input_folder)
    if not assemblies:
        print(f"\nNo .SLDASM files found in:\n{input_folder}")
        return None

    print("\nAssemblies found:")
    for i, assembly in enumerate(assemblies, 1):
        print(f"  {i}. {assembly.name}")

    selected = assemblies[0]
    print(f"Selected: {selected}")
    return selected


def print_summary(components, unique_parts):
    parts = sum(1 for c in components if c.is_part)
    assemblies = sum(1 for c in components if c.is_assembly)

    print("\n" + "=" * 60)
    print("ASSEMBLY SUMMARY")
    print("=" * 60)
    print(f"Total components: {len(components)}")
    print(f"Part instances:   {parts}")
    print(f"Sub-assemblies:   {assemblies}")
    print(f"Unique part/config: {len(unique_parts)}")


def save_run_summary(
    output_folder,
    assembly_path,
    scan_result,
    drawing_result,
    pdf_result=None,
):
    data = {
        "assembly": str(assembly_path),

        "scan": {
            "unique_parts": scan_result.get("unique_parts", 0),
            "sheet_metal": scan_result.get("sheet_metal", 0),
            "not_sheet_metal": scan_result.get("not_sheet_metal", 0),
            "failed": scan_result.get("failed", 0),
            "total_quantity": scan_result.get("total_quantity", 0),
            "failed_parts": scan_result.get("failed_parts", []),
        },

        "drawings": {
            "created": len(drawing_result.get("files", []))
            if drawing_result else 0,

            "failed_parts": drawing_result.get("failed_parts", [])
            if drawing_result else [],

            "error": drawing_result.get("error")
            if drawing_result else None,

            "files": [
                {
                    "thickness": item.get("thickness"),
                    "unique_parts": item.get("unique_parts"),
                    "quantity": item.get("quantity"),
                    "path": str(item.get("path")),
                    "is_dwg": item.get("is_dwg", False),
                }
                for item in (
                    drawing_result.get("files", [])
                    if drawing_result
                    else []
                )
            ],
        },

        "pdf": {
            "created": pdf_result.get("created", 0)
            if pdf_result else 0,

            "failed": pdf_result.get("failed", 0)
            if pdf_result else 0,

            "failed_parts": pdf_result.get("failed_parts", [])
            if pdf_result else [],
        },
    }

    path = Path(output_folder) / "run_summary.json"

    try:
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"Run summary saved: {path}")

    except Exception as exc:
        print(f"Could not save run summary: {exc}")


def build_pdf_exported_parts_from_dxf(dxf_folder):
    """
    PDF-only / Resume mode.

    DXFهای موجود در _parts_dxf را پیدا می‌کند و با همان
    ساختار مورد نیاز PdfDrawingExporter برمی‌گرداند.
    """

    dxf_folder = Path(dxf_folder)

    exported_parts = []

    dxf_files = sorted(
        dxf_folder.glob("*.dxf"),
        key=lambda p: p.name.lower(),
    )

    for dxf_path in dxf_files:
        if not dxf_path.is_file():
            continue

        if dxf_path.stat().st_size <= 0:
            continue

        item = {
            "name": dxf_path.stem,
            "thickness": 0,
            "quantity": 1,
        }

        exported_parts.append(
            {
                "item": item,
                "dxf": dxf_path,
            }
        )

    return exported_parts


def run_pdf_only(output_folder):
    """
    فقط PDF را از DXFهای قبلی تولید می‌کند.
    هیچ اتصال یا Export مجددی به SolidWorks انجام نمی‌شود.
    """

    pdf_source_folder = (
        Path(output_folder)
        / ("by_thickness" if GROUP_BY_THICKNESS else "combined")
        / "_parts_dxf"
    )

    print("\n" + "=" * 60)
    print("PDF RESUME MODE")
    print("=" * 60)

    print(f"DXF source folder:")
    print(pdf_source_folder)

    if not pdf_source_folder.exists():
        print("\nERROR: DXF folder does not exist.")
        return 1

    exported_parts = build_pdf_exported_parts_from_dxf(
        pdf_source_folder
    )

    if not exported_parts:
        print("\nNo usable DXF files found.")
        return 1

    print(f"\nExisting DXF files: {len(exported_parts)}")
    print("SolidWorks export will be skipped.")

    pdf_output_folder = Path(output_folder) / "pdf"

    try:
        pdf_exporter = PdfDrawingExporter(
            paper=PDF_PAPER,
            orientation=PDF_ORIENTATION,
            company=PDF_COMPANY,
            drawn_by=PDF_DRAWN_BY,
            checked_by=PDF_CHECKED_BY,
            approved_by=PDF_APPROVED_BY,
            material=PDF_MATERIAL,
        )

        pdf_result = pdf_exporter.export_all(
            exported_parts,
            pdf_output_folder,
        )

        print("\n" + "=" * 60)
        print("PDF RESUME FINISHED")
        print("=" * 60)

        print(
            f"PDF drawings created: "
            f"{pdf_result.get('created', 0)}"
        )

        print(
            f"PDF failures: "
            f"{pdf_result.get('failed', 0)}"
        )

        print(f"Output: {pdf_output_folder}")

        return 0

    except Exception as exc:
        print("\n" + "=" * 60)
        print("PDF EXPORT FAILED")
        print("=" * 60)

        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()

        return 1


def main():
    project_folder = Path(__file__).resolve().parent
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"

    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)

    print("\n" + "=" * 60)
    print("SOLIDWORKS SHEET METAL -> DXF/DWG/PDF EXPORTER")
    print("=" * 60)

    # ============================================================
    # PDF ONLY MODE
    #
    # اگر PDF_ONLY=True باشد:
    # - SolidWorks باز نمی‌شود
    # - اسمبلی دوباره Scan نمی‌شود
    # - DXF دوباره Export نمی‌شود
    # - DWG دوباره ساخته نمی‌شود
    # - فقط DXFهای موجود به PDF تبدیل می‌شوند
    # ============================================================

    if PDF_ONLY and EXPORT_PDF:
        return run_pdf_only(output_folder)

    # ============================================================
    # ORIGINAL WORKING SOLIDWORKS FLOW
    # ============================================================

    sw_app = connect_to_solidworks()
    if sw_app is None:
        return 1

    assembly_path = select_assembly(input_folder)
    if assembly_path is None:
        return 1

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_path):
        print("Could not open assembly. Stopping.")
        return 1

    components = assembly.get_components()

    if not components:
        print("No components found. Stopping.")
        return 1

    unique_parts = assembly.get_unique_part_quantities(components)

    print_summary(components, unique_parts)

    detector = SheetMetalDetector(
        sw_app,
        assembly_path=assembly_path,
        read_thickness=GROUP_BY_THICKNESS,
    )

    scan_result = detector.scan_and_export(
        unique_parts,
        output_folder,
    )

    drawing_result = None

    exporter = DwgExporter(
        sw_app,
        assembly_path=assembly_path,
        reconnect=connect_to_solidworks,
    )

    builder = ThicknessDwgBuilder(
        exporter,
        gap=PART_GAP,
        max_row_width=MAX_ROW_WIDTH,
        repeat_by_quantity=REPEAT_BY_QUANTITY,
        oda_path=ODA_CONVERTER_PATH,
        group_by_thickness=GROUP_BY_THICKNESS,
    )

    try:
        drawing_result = builder.build(
            scan_result["sheet_metal_parts"],
            output_folder
            / (
                "by_thickness"
                if GROUP_BY_THICKNESS
                else "combined"
            ),
            base_name=assembly_path.stem,
        )

    except Exception as exc:
        print("\nDRAWING EXPORT FAILED")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()

        drawing_result = {
            "files": [],
            "failed_parts": [],
            "error": str(exc),
        }

    finally:
        exporter.activate_assembly()

        gc.collect()

        try:
            pythoncom.CoFreeUnusedLibraries()
        except Exception:
            pass

    # ============================================================
    # PDF EXPORT
    # ============================================================

    pdf_result = None

    if EXPORT_PDF:

        exported_parts = []

        # مسیر DXFهای تکی که ThicknessDwgBuilder قبلاً ساخته است
        dxf_folder = (
            output_folder
            / (
                "by_thickness"
                if GROUP_BY_THICKNESS
                else "combined"
            )
            / "_parts_dxf"
        )

        if dxf_folder.exists():

            exported_parts = (
                build_pdf_exported_parts_from_dxf(
                    dxf_folder
                )
            )

        if not exported_parts:

            print(
                "\nNo flat pattern DXFs were exported, "
                "so no PDF drawings were made."
            )

        else:

            try:
                pdf_exporter = PdfDrawingExporter(
                    paper=PDF_PAPER,
                    orientation=PDF_ORIENTATION,
                    company=PDF_COMPANY,
                    drawn_by=PDF_DRAWN_BY,
                    checked_by=PDF_CHECKED_BY,
                    approved_by=PDF_APPROVED_BY,
                    material=PDF_MATERIAL,
                )

                pdf_output_folder = output_folder / "pdf"

                pdf_result = pdf_exporter.export_all(
                    exported_parts,
                    pdf_output_folder,
                )

            except Exception as exc:

                print("\nPDF EXPORT FAILED")
                print(f"{type(exc).__name__}: {exc}")
                traceback.print_exc()

    save_run_summary(
        output_folder,
        assembly_path,
        scan_result,
        drawing_result,
        pdf_result,
    )

    print("\n" + "=" * 60)
    print("PROCESS FINISHED")
    print("=" * 60)

    print(
        f"Sheet metal (unique): "
        f"{scan_result['sheet_metal']}"
    )

    print(
        f"Sheet metal (quantity): "
        f"{scan_result['total_quantity']}"
    )

    print(
        f"Not sheet metal: "
        f"{scan_result['not_sheet_metal']}"
    )

    print(
        f"Scan failures: "
        f"{scan_result['failed']}"
    )

    if drawing_result:

        print(
            f"Drawings created: "
            f"{len(drawing_result['files'])}"
        )

        print(
            f"Export failures: "
            f"{len(drawing_result['failed_parts'])}"
        )

    if pdf_result:

        print(
            f"PDF drawings created: "
            f"{pdf_result['created']} "
            f"(failures: {pdf_result['failed']})"
        )

    print(f"Output: {output_folder}")

    return 0


if __name__ == "__main__":

    pythoncom.CoInitialize()

    try:
        sys.exit(main())

    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(130)

    except Exception as exc:
        print("\nUNEXPECTED ERROR")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()
        sys.exit(1)

    finally:
        pythoncom.CoUninitialize()

