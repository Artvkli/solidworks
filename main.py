# -*- coding: utf-8 -*-

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
from solidworks.pdf_builder import SolidWorksPdfBuilder


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

# If you have a custom SolidWorks Drawing Template (.drwdot),
# put its full path here.
#
# Example:
# PDF_DRAWING_TEMPLATE = r"C:\SolidWorks\Templates\AHU_A4.drwdot"
#
# None = use SolidWorks default drawing template.
PDF_DRAWING_TEMPLATE = None

# Optional bend-table template (.sldbndtbt).
# None = automatically search common SolidWorks installation folders.
PDF_BEND_TABLE_TEMPLATE = None

# False = temporary SolidWorks drawings are closed after PDF export.
# True  = keep the .slddrw files open/save them if supported by builder.
PDF_KEEP_DRAWINGS = False


# ============================================================
# SOLIDWORKS CONNECTION
# ============================================================


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


# ============================================================
# ASSEMBLY SELECTION
# ============================================================


def select_assembly(input_folder):
    assemblies = find_assemblies(input_folder)

    if not assemblies:
        print(f"\nNo .SLDASM files found in:\n{input_folder}")
        return None

    print("\nAssemblies found:")

    for i, assembly in enumerate(assemblies, 1):
        print(f"  {i}. {assembly.name}")

    # Current behavior: automatically select the first assembly.
    selected = assemblies[0]

    print(f"Selected: {selected}")

    return selected


# ============================================================
# SUMMARY
# ============================================================


def print_summary(components, unique_parts):
    parts = sum(1 for c in components if c.is_part)
    assemblies = sum(1 for c in components if c.is_assembly)

    print("\n" + "=" * 60)
    print("ASSEMBLY SUMMARY")
    print("=" * 60)
    print(f"Total components:  {len(components)}")
    print(f"Part instances:    {parts}")
    print(f"Sub-assemblies:     {assemblies}")
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
            "created": (len(drawing_result.get("files", [])) if drawing_result else 0),
            "failed_parts": (
                drawing_result.get("failed_parts", []) if drawing_result else []
            ),
            "error": (drawing_result.get("error") if drawing_result else None),
            "files": [
                {
                    "thickness": item.get("thickness"),
                    "unique_parts": item.get("unique_parts"),
                    "quantity": item.get("quantity"),
                    "path": str(item.get("path")),
                    "is_dwg": item.get("is_dwg", False),
                }
                for item in (drawing_result.get("files", []) if drawing_result else [])
            ],
        },
        "pdfs": {
            "created": (len(pdf_result.get("files", [])) if pdf_result else 0),
            "failed_parts": (pdf_result.get("failed_parts", []) if pdf_result else []),
            "error": (pdf_result.get("error") if pdf_result else None),
            "files": [
                {
                    "name": item.get("name"),
                    "thickness": item.get("thickness"),
                    "quantity": item.get("quantity"),
                    "path": str(item.get("path")),
                }
                for item in (pdf_result.get("files", []) if pdf_result else [])
            ],
        },
    }

    path = Path(output_folder) / "run_summary.json"

    try:
        path.write_text(
            json.dumps(
                data,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        print(f"Run summary saved: {path}")

    except Exception as exc:
        print(f"Could not save run summary: {exc}")


# ============================================================
# MAIN
# ============================================================


def main():
    project_folder = Path(__file__).resolve().parent

    input_folder = project_folder / "input"
    output_folder = project_folder / "output"

    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)

    print("\n" + "=" * 60)
    print("MANDEGAR SYSTEM")
    print("SOLIDWORKS SHEET METAL -> DXF / DWG / PDF EXPORTER")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. CONNECT TO SOLIDWORKS
    # --------------------------------------------------------

    sw_app = connect_to_solidworks()

    if sw_app is None:
        return 1

    # --------------------------------------------------------
    # 2. FIND ASSEMBLY
    # --------------------------------------------------------

    assembly_path = select_assembly(input_folder)

    if assembly_path is None:
        return 1

    # --------------------------------------------------------
    # 3. OPEN ASSEMBLY
    # --------------------------------------------------------

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_path):
        print("Could not open assembly. Stopping.")
        return 1

    # --------------------------------------------------------
    # 4. GET COMPONENTS
    # --------------------------------------------------------

    components = assembly.get_components()

    if not components:
        print("No components found. Stopping.")
        return 1

    # --------------------------------------------------------
    # 5. UNIQUE PARTS + QUANTITIES
    # --------------------------------------------------------

    unique_parts = assembly.get_unique_part_quantities(components)

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

    builder = ThicknessDwgBuilder(
        exporter,
        gap=PART_GAP,
        max_row_width=MAX_ROW_WIDTH,
        repeat_by_quantity=REPEAT_BY_QUANTITY,
        oda_path=ODA_CONVERTER_PATH,
        group_by_thickness=GROUP_BY_THICKNESS,
    )

    # --------------------------------------------------------
    # 9. PDF BUILDER
    # --------------------------------------------------------

    pdf_builder = SolidWorksPdfBuilder(
        sw_app,
        exporter=exporter,
        drawing_template=PDF_DRAWING_TEMPLATE,
        bend_table_template=PDF_BEND_TABLE_TEMPLATE,
        keep_drawings=PDF_KEEP_DRAWINGS,
        add_dimensions=True,
        add_bend_table=True,
        add_qty_note=True,
    )

    drawing_result = None
    pdf_result = None

    # --------------------------------------------------------
    # 10. EXPORT CURRENT DWG / DXF
    # --------------------------------------------------------

    try:
        drawing_result = builder.build(
            scan_result["sheet_metal_parts"],
            output_folder / ("by_thickness" if GROUP_BY_THICKNESS else "combined"),
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

    # --------------------------------------------------------
    # 11. EXPORT ONE PDF PER PART
    # --------------------------------------------------------

    try:
        pdf_result = pdf_builder.build(
            scan_result["sheet_metal_parts"],
            output_folder / PDF_OUTPUT_FOLDER_NAME,
        )

    except Exception as exc:
        print("\nPDF EXPORT FAILED")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()

        pdf_result = {
            "files": [],
            "failed_parts": [],
            "error": str(exc),
        }

    # --------------------------------------------------------
    # 12. CLEANUP
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

    # --------------------------------------------------------
    # 13. SAVE RUN SUMMARY
    # --------------------------------------------------------

    save_run_summary(
        output_folder,
        assembly_path,
        scan_result,
        drawing_result,
        pdf_result,
    )

    # --------------------------------------------------------
    # 14. FINAL SUMMARY
    # --------------------------------------------------------

    print("\n" + "=" * 60)
    print("PROCESS FINISHED")
    print("=" * 60)

    print(f"Sheet metal (unique): {scan_result['sheet_metal']}")

    print(f"Sheet metal (quantity): {scan_result['total_quantity']}")

    print(f"Not sheet metal: {scan_result['not_sheet_metal']}")

    print(f"Scan failures: {scan_result['failed']}")

    if drawing_result:
        print(f"Drawings created: {len(drawing_result['files'])}")

        print(f"Drawing failures: {len(drawing_result['failed_parts'])}")

    if pdf_result:
        print(f"PDFs created: {len(pdf_result['files'])}")

        print(f"PDF failures: {len(pdf_result['failed_parts'])}")

    print(f"Output: {output_folder}")

    return 0


# ============================================================
# PROGRAM ENTRY
# ============================================================

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
