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
from solidworks.bend_table import BendTableReader
from solidworks.bend_table_pdf import BendTablePdfWriter
from solidworks.dwg_exporter import DwgExporter
from solidworks.thickness_dwg import ThicknessDwgBuilder
from solidworks.sheet_metal import SheetMetalDetector

GROUP_BY_THICKNESS = True
REPEAT_BY_QUANTITY = False
ODA_CONVERTER_PATH = None
PART_GAP = 20.0
MAX_ROW_WIDTH = 6000.0

# ---- bend tables (Tag / Direction / Angle / Inner Radius) written to PDF ----
# The DXF/DWG output is NOT changed by this. PDFs go to output/bend_tables_pdf/
EXPORT_BEND_TABLE_PDF = True
COMBINE_BEND_TABLE_PDF = True  # also write ONE PDF with all parts (one page per part)
BEND_TABLE_LIMIT = None  # e.g. 3 = test on the first 3 parts only, None = all parts
DRAWING_TEMPLATE_PATH = (
    None  # full path of a .drwdot file, only if it is not found automatically
)
BEND_TABLE_TEMPLATE_PATH = (
    None  # full path of a .sldbndtbt file, only if you want a specific one
)


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


def save_run_summary(output_folder, assembly_path, scan_result, drawing_result):
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
            "created": len(drawing_result.get("files", [])) if drawing_result else 0,
            "failed_parts": drawing_result.get("failed_parts", [])
            if drawing_result
            else [],
            "error": drawing_result.get("error") if drawing_result else None,
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
    }
    path = Path(output_folder) / "run_summary.json"
    try:
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"Run summary saved: {path}")
    except Exception as exc:
        print(f"Could not save run summary: {exc}")


def main():
    project_folder = Path(__file__).resolve().parent
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"
    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)

    print("\n" + "=" * 60)
    print("SOLIDWORKS SHEET METAL -> DXF/DWG EXPORTER")
    print("=" * 60)

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
        sw_app, assembly_path=assembly_path, read_thickness=GROUP_BY_THICKNESS
    )
    scan_result = detector.scan_and_export(unique_parts, output_folder)

    # ---- bend tables: read from SolidWorks, then written to PDF (DXF/DWG untouched) ----
    if EXPORT_BEND_TABLE_PDF:
        try:
            reader = BendTableReader(
                sw_app,
                assembly_path=assembly_path,
                drawing_template=DRAWING_TEMPLATE_PATH,
                bend_template=BEND_TABLE_TEMPLATE_PATH,
            )
            reader.read_all(
                scan_result["sheet_metal_parts"], output_folder, limit=BEND_TABLE_LIMIT
            )
            BendTablePdfWriter().write_all(
                scan_result["sheet_metal_parts"],
                output_folder / "bend_tables_pdf",
                combined_name=f"{assembly_path.stem}_bend_tables"
                if COMBINE_BEND_TABLE_PDF
                else None,
            )
        except Exception as exc:
            print(
                f"\nBEND TABLES FAILED (continuing without them): {type(exc).__name__}: {exc}"
            )
            traceback.print_exc()

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
            output_folder / ("by_thickness" if GROUP_BY_THICKNESS else "combined"),
            base_name=assembly_path.stem,
        )
    except Exception as exc:
        print("\nDRAWING EXPORT FAILED")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()
        drawing_result = {"files": [], "failed_parts": [], "error": str(exc)}
    finally:
        exporter.activate_assembly()
        gc.collect()
        try:
            pythoncom.CoFreeUnusedLibraries()
        except Exception:
            pass

    save_run_summary(output_folder, assembly_path, scan_result, drawing_result)

    print("\n" + "=" * 60)
    print("PROCESS FINISHED")
    print("=" * 60)
    print(f"Sheet metal (unique): {scan_result['sheet_metal']}")
    print(f"Sheet metal (quantity): {scan_result['total_quantity']}")
    print(f"Not sheet metal: {scan_result['not_sheet_metal']}")
    print(f"Scan failures: {scan_result['failed']}")
    if drawing_result:
        print(f"Drawings created: {len(drawing_result['files'])}")
        print(f"Export failures: {len(drawing_result['failed_parts'])}")
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
