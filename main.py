
from pathlib import Path
import gc
import json
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
from solidworks.thickness_dwg import ThicknessDwgBuilder
from solidworks.sheet_metal import SheetMetalDetector

# ============================================================
# OLD SOLIDWORKS PDF BUILDER
# ============================================================
# from solidworks.pdf_builder import SolidWorksPdfBuilder

# ============================================================
# NEW PDF BUILDER
# ============================================================

from solidworks.pdf_builder import PdfDrawingExporter


# ============================================================
# SETTINGS
# ============================================================

GROUP_BY_THICKNESS = True
REPEAT_BY_QUANTITY = False

# ODA File Converter:
# None = auto-detect installed ODA File Converter
ODA_CONVERTER_PATH = None

# Current DXF/DWG layout settings
PART_GAP = 20.0
MAX_ROW_WIDTH = 6000.0

# PDF output
PDF_OUTPUT_FOLDER_NAME = "pdf_by_thickness"

# ============================================================
# OLD SOLIDWORKS PDF SETTINGS
# ============================================================
#
# فعلاً حذف نشده‌اند.
# فقط برای تست PDF جدید استفاده نمی‌شوند.
#

PDF_DRAWING_TEMPLATE = None
PDF_BEND_TABLE_TEMPLATE = None
PDF_KEEP_DRAWINGS = False


# ============================================================
# SOLIDWORKS CONNECTION
# ============================================================

def connect_to_solidworks():

    print("\nConnecting to SolidWorks...")

    try:

        sw_app = win32com.client.Dispatch(
            "SldWorks.Application"
        )

        sw_app.Visible = True

        print("Connected successfully.")

        return sw_app

    except Exception as exc:

        print(
            f"Could not connect to SolidWorks: {exc}"
        )

        return None


# ============================================================
# ASSEMBLY SELECTION
# ============================================================

def select_assembly(input_folder):

    assemblies = find_assemblies(
        input_folder
    )

    if not assemblies:

        print(
            f"\nNo .SLDASM files found in:\n"
            f"{input_folder}"
        )

        return None

    print("\nAssemblies found:")

    for i, assembly in enumerate(
        assemblies,
        1,
    ):

        print(
            f"  {i}. {assembly.name}"
        )

    # Current behavior:
    # automatically select the first assembly.

    selected = assemblies[0]

    print(
        f"Selected: {selected}"
    )

    return selected


# ============================================================
# SUMMARY
# ============================================================

def print_summary(
    components,
    unique_parts,
):

    parts = sum(
        1
        for c in components
        if c.is_part
    )

    assemblies = sum(
        1
        for c in components
        if c.is_assembly
    )

    print(
        "\n" + "=" * 60
    )

    print(
        "ASSEMBLY SUMMARY"
    )

    print(
        "=" * 60
    )

    print(
        f"Total components:  {len(components)}"
    )

    print(
        f"Part instances:    {parts}"
    )

    print(
        f"Sub-assemblies:     {assemblies}"
    )

    print(
        f"Unique part/config: {len(unique_parts)}"
    )


# ============================================================
# OLD RUN SUMMARY
# ============================================================
#
# فعلاً نگه داشته شده.
# در تست PDF جدید استفاده نمی‌شود.
#

def save_run_summary(
    output_folder,
    assembly_path,
    scan_result,
    drawing_result,
    pdf_result=None,
):

    data = {

        "assembly": str(
            assembly_path
        ),

        "scan": {

            "unique_parts":
                scan_result.get(
                    "unique_parts",
                    0,
                ),

            "sheet_metal":
                scan_result.get(
                    "sheet_metal",
                    0,
                ),

            "not_sheet_metal":
                scan_result.get(
                    "not_sheet_metal",
                    0,
                ),

            "failed":
                scan_result.get(
                    "failed",
                    0,
                ),

            "total_quantity":
                scan_result.get(
                    "total_quantity",
                    0,
                ),

            "failed_parts":
                scan_result.get(
                    "failed_parts",
                    [],
                ),
        },

        "drawings": {

            "created":
                (
                    len(
                        drawing_result.get(
                            "files",
                            [],
                        )
                    )
                    if drawing_result
                    else 0
                ),

            "failed_parts":
                (
                    drawing_result.get(
                        "failed_parts",
                        [],
                    )
                    if drawing_result
                    else []
                ),

            "error":
                (
                    drawing_result.get(
                        "error"
                    )
                    if drawing_result
                    else None
                ),
        },

        "pdfs": {

            "created":
                (
                    len(
                        pdf_result.get(
                            "files",
                            [],
                        )
                    )
                    if pdf_result
                    else 0
                ),

            "failed_parts":
                (
                    pdf_result.get(
                        "failed_parts",
                        [],
                    )
                    if pdf_result
                    else []
                ),

            "error":
                (
                    pdf_result.get(
                        "error"
                    )
                    if pdf_result
                    else None
                ),
        },
    }

    path = (
        Path(output_folder)
        / "run_summary.json"
    )

    try:

        path.write_text(
            json.dumps(
                data,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        print(
            f"Run summary saved: {path}"
        )

    except Exception as exc:

        print(
            f"Could not save run summary: {exc}"
        )


# ============================================================
# MAIN
# ============================================================

def main():

    project_folder = (
        Path(__file__).resolve().parent
    )

    input_folder = (
        project_folder / "input"
    )

    output_folder = (
        project_folder / "output"
    )

    input_folder.mkdir(
        exist_ok=True
    )

    output_folder.mkdir(
        exist_ok=True
    )

    print(
        "\n" + "=" * 60
    )

    print(
        "MANDEGAR SYSTEM"
    )

    print(
        "SOLIDWORKS SHEET METAL -> "
        "DXF / DWG / PDF EXPORTER"
    )

    print(
        "=" * 60
    )

    # --------------------------------------------------------
    # 1. CONNECT TO SOLIDWORKS
    # --------------------------------------------------------

    sw_app = connect_to_solidworks()

    if sw_app is None:
        return 1

    # --------------------------------------------------------
    # 2. FIND ASSEMBLY
    # --------------------------------------------------------

    assembly_path = select_assembly(
        input_folder
    )

    if assembly_path is None:
        return 1

    # --------------------------------------------------------
    # 3. OPEN ASSEMBLY
    # --------------------------------------------------------

    assembly = SolidWorksAssembly(
        sw_app
    )

    if not assembly.open(
        assembly_path
    ):

        print(
            "Could not open assembly. "
            "Stopping."
        )

        return 1

    # --------------------------------------------------------
    # 4. GET COMPONENTS
    # --------------------------------------------------------

    components = (
        assembly.get_components()
    )

    if not components:

        print(
            "No components found. "
            "Stopping."
        )

        return 1

    # --------------------------------------------------------
    # 5. UNIQUE PARTS + QUANTITIES
    # --------------------------------------------------------

    unique_parts = (
        assembly.get_unique_part_quantities(
            components
        )
    )

    print_summary(
        components,
        unique_parts,
    )

    # --------------------------------------------------------
    # 6. SHEET-METAL SCAN
    # --------------------------------------------------------

    detector = SheetMetalDetector(
        sw_app,
        assembly_path=assembly_path,
        read_thickness=GROUP_BY_THICKNESS,
    )

    scan_result = detector.scan_and_export(
        unique_parts,
        output_folder,
    )

    # --------------------------------------------------------
    # 7. CREATE COMMON EXPORTER
    # --------------------------------------------------------

    exporter = DwgExporter(
        sw_app,
        assembly_path=assembly_path,
        reconnect=connect_to_solidworks,
    )

    # --------------------------------------------------------
    # 8. CURRENT DXF / DWG BUILDER
    # --------------------------------------------------------
    #
    # فعلاً حذف نشده.
    #
    # فقط برای اینکه DXFهای Flat Pattern ساخته شوند
    # این بخش را نگه می‌داریم.
    #
    # --------------------------------------------------------

    builder = ThicknessDwgBuilder(
        exporter,
        gap=PART_GAP,
        max_row_width=MAX_ROW_WIDTH,
        repeat_by_quantity=REPEAT_BY_QUANTITY,
        oda_path=ODA_CONVERTER_PATH,
        group_by_thickness=GROUP_BY_THICKNESS,
    )

    drawing_result = None
    pdf_result = None

    # ========================================================
    # 9. OLD DWG / DXF COMBINED EXPORT
    # ========================================================
    #
    # در تست فعلی این بخش اجرا نمی‌شود.
    #
    # اگر بعداً خواستی فعالش کنی، فقط کامنت‌ها را بردار.
    #
    # ========================================================

    # try:
    #
    #     drawing_result = builder.build(
    #         scan_result["sheet_metal_parts"],
    #         output_folder / (
    #             "by_thickness"
    #             if GROUP_BY_THICKNESS
    #             else "combined"
    #         ),
    #         base_name=assembly_path.stem,
    #     )
    #
    # except Exception as exc:
    #
    #     print("\nDRAWING EXPORT FAILED")
    #     print(f"{type(exc).__name__}: {exc}")
    #     traceback.print_exc()
    #
    #     drawing_result = {
    #         "files": [],
    #         "failed_parts": [],
    #         "error": str(exc),
    #     }


    # ========================================================
    # 10. GENERATE DXF FLAT PATTERNS FOR PDF TEST
    # ========================================================
    #
    # IMPORTANT:
    #
    # PdfDrawingExporter به DXF احتیاج دارد.
    #
    # بنابراین فعلاً ThicknessDwgBuilder را اجرا می‌کنیم
    # تا DXFهای لازم تولید شوند.
    #
    # این بخش خروجی DWG نهایی را برای کاربر نمی‌سازد؛
    # هدف آن فقط تهیه DXF برای PDF Builder جدید است.
    #
    # ========================================================

    exported_parts = []

    try:

        print()
        print(
            "=" * 60
        )

        print(
            "PREPARING DXFs FOR PDF TEST"
        )

        print(
            "=" * 60
        )

        drawing_result = builder.build(
            scan_result["sheet_metal_parts"],
            output_folder / (
                "by_thickness"
                if GROUP_BY_THICKNESS
                else "combined"
            ),
            base_name=assembly_path.stem,
        )

        # ----------------------------------------------------
        # IMPORTANT
        #
        # اگر build() خروجی را دقیقاً به شکل:
        #
        # [
        #     {
        #         "item": {...},
        #         "dxf": Path(...)
        #     }
        # ]
        #
        # برگرداند، مستقیماً استفاده می‌شود.
        # ----------------------------------------------------

        exported_parts = (
            drawing_result.get(
                "files",
                []
            )
        )

    except Exception as exc:

        print(
            "\nDXF PREPARATION FAILED"
        )

        print(
            f"{type(exc).__name__}: {exc}"
        )

        traceback.print_exc()

        return 1


    # ========================================================
    # 11. OLD SOLIDWORKS PDF BUILDER
    # ========================================================
    #
    # کاملاً نگه داشته شده ولی اجرا نمی‌شود.
    #
    # ========================================================

    # pdf_builder = SolidWorksPdfBuilder(
    #     sw_app,
    #     exporter=exporter,
    #     drawing_template=PDF_DRAWING_TEMPLATE,
    #     bend_table_template=PDF_BEND_TABLE_TEMPLATE,
    #     keep_drawings=PDF_KEEP_DRAWINGS,
    #     add_dimensions=True,
    #     add_bend_table=True,
    #     add_qty_note=True,
    # )

    # try:
    #
    #     pdf_result = pdf_builder.build(
    #         scan_result["sheet_metal_parts"],
    #         output_folder / PDF_OUTPUT_FOLDER_NAME,
    #     )
    #
    # except Exception as exc:
    #
    #     print("\nPDF EXPORT FAILED")
    #     print(f"{type(exc).__name__}: {exc}")
    #     traceback.print_exc()
    #
    #     pdf_result = {
    #         "files": [],
    #         "failed_parts": [],
    #         "error": str(exc),
    #     }


    # ========================================================
    # 12. NEW PDF DRAWING EXPORTER
    # ========================================================
    #
    # این قسمت PDF Builder جدید تو است.
    #
    # ========================================================

    pdf_builder = PdfDrawingExporter(

        paper="A4",

        orientation="auto",

        company="TECHNICAL DEPARTMENT",

        drawn_by="",

        checked_by="",

        approved_by="",

        material="",

        deburr_note=(
            "DEBURR AND BREAK SHARP EDGES"
        ),

        date_text="",
    )

    # --------------------------------------------------------
    # 13. EXPORT ONE PDF PER PART
    # --------------------------------------------------------

    try:

        pdf_output_folder = (
            output_folder
            / PDF_OUTPUT_FOLDER_NAME
        )

        print()
        print(
            "=" * 60
        )

        print(
            "NEW PDF DRAWING EXPORT"
        )

        print(
            "=" * 60
        )

        pdf_result = (
            pdf_builder.export_all(
                exported_parts,
                pdf_output_folder,
            )
        )

    except Exception as exc:

        print(
            "\nNEW PDF EXPORT FAILED"
        )

        print(
            f"{type(exc).__name__}: {exc}"
        )

        traceback.print_exc()

        pdf_result = {
            "created": 0,
            "failed": 1,
            "failed_parts": [
                (
                    "PDF EXPORT",
                    str(exc),
                )
            ],
            "files": [],
            "report_path": None,
        }


    # --------------------------------------------------------
    # 14. CLEANUP
    # --------------------------------------------------------

    finally:

        try:

            exporter.activate_assembly()

        except Exception:

            pass

        gc.collect()

        try:

            pythoncom.CoFreeUnusedLibraries()

        except Exception:

            pass


    # ========================================================
    # 15. OLD RUN SUMMARY
    # ========================================================
    #
    # فعلاً غیرفعال است تا تست فقط روی PDF جدید باشد.
    #
    # ========================================================

    # save_run_summary(
    #     output_folder,
    #     assembly_path,
    #     scan_result,
    #     drawing_result,
    #     pdf_result,
    # )


    # ========================================================
    # 16. FINAL SUMMARY
    # ========================================================

    print()
    print(
        "=" * 60
    )

    print(
        "PDF TEST FINISHED"
    )

    print(
        "=" * 60
    )

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

    if pdf_result:

        print(
            f"PDFs created: "
            f"{pdf_result['created']}"
        )

        print(
            f"PDF failures: "
            f"{pdf_result['failed']}"
        )

        if pdf_result.get(
            "report_path"
        ):

            print(
                f"PDF report: "
                f"{pdf_result['report_path']}"
            )

    print(
        f"Output: "
        f"{output_folder}"
    )

    return 0


# ============================================================
# PROGRAM ENTRY
# ============================================================

if __name__ == "__main__":

    pythoncom.CoInitialize()

    try:

        sys.exit(
            main()
        )

    except KeyboardInterrupt:

        print(
            "\nInterrupted by user."
        )

        sys.exit(130)

    except Exception as exc:

        print(
            "\nUNEXPECTED ERROR"
        )

        print(
            f"{type(exc).__name__}: {exc}"
        )

        traceback.print_exc()

        sys.exit(1)

    finally:

        pythoncom.CoUninitialize()

