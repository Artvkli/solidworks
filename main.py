# main.py

from pathlib import Path

import pythoncom
import win32com.client

from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.dwg_exporter import DwgExporter
from solidworks.thickness_dwg import ThicknessDwgBuilder
from solidworks.sheet_metal import SheetMetalDetector


# =========================================================
# SETTINGS
# =========================================================

# Group the sheet metal parts by thickness (one drawing per thickness)?
#   False -> ONE drawing with all sheet metal parts (thickness is not read)
#   True  -> one drawing per thickness (needs the thickness to be readable)
GROUP_BY_THICKNESS = True

# Put every part as many times as its quantity into the drawing?
#   False -> each part once, with a label "xN" (recommended for checking)
#   True  -> N copies of each part (for cutting)
REPEAT_BY_QUANTITY = False

# Full path of ODAFileConverter.exe (only needed if it is not found automatically)
ODA_CONVERTER_PATH = None

# Gap between parts and maximum row width (in the DXF units, normally mm)
PART_GAP = 20.0
MAX_ROW_WIDTH = 6000.0


# =========================================================
# CONNECT TO SOLIDWORKS
# =========================================================


def connect_to_solidworks():
    print()
    print("=" * 60)
    print("SOLIDWORKS SHEET METAL SCANNER")
    print("=" * 60)

    print()
    print("Connecting to SolidWorks...")

    try:
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True

        print("Connected successfully.")
        return sw_app

    except Exception as e:
        print(f"Could not connect to SolidWorks:\n{e}")
        return None


# =========================================================
# SELECT ASSEMBLY
# =========================================================


def select_assembly(input_folder):
    assemblies = find_assemblies(input_folder)

    if not assemblies:
        print()
        print(f"No .SLDASM files found in:\n{input_folder}")
        return None

    print()
    print("Assemblies found:")

    for index, assembly in enumerate(assemblies, start=1):
        print(f"{index}. {assembly.name}")

    # Use the first assembly automatically
    selected = assemblies[0]

    print()
    print("Selected:")
    print(selected)

    return selected


# =========================================================
# PRINT COMPONENT SUMMARY
# =========================================================


def print_summary(components, unique_parts):
    parts = 0
    sub_assemblies = 0
    unknown = 0

    for component in components:
        if component.is_part:
            parts += 1
        elif component.is_assembly:
            sub_assemblies += 1
        else:
            unknown += 1

    print()
    print("=" * 60)
    print("ASSEMBLY SUMMARY")
    print("=" * 60)

    print(f"Total components: {len(components)}")
    print(f"Part instances: {parts}")
    print(f"Sub-assemblies: {sub_assemblies}")
    print(f"Unknown: {unknown}")
    print(f"Unique parts: {len(unique_parts)}")


# =========================================================
# MAIN
# =========================================================


def main():
    project_folder = Path(__file__).resolve().parent
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"

    input_folder.mkdir(parents=True, exist_ok=True)
    output_folder.mkdir(parents=True, exist_ok=True)

    # -----------------------------------------------------
    # Connect
    # -----------------------------------------------------

    sw_app = connect_to_solidworks()
    if sw_app is None:
        return

    # -----------------------------------------------------
    # Find assembly
    # -----------------------------------------------------

    assembly_path = select_assembly(input_folder)
    if assembly_path is None:
        return

    # -----------------------------------------------------
    # Open assembly
    # -----------------------------------------------------

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_path):
        print()
        print("Could not open the assembly. Stopping.")
        return

    # -----------------------------------------------------
    # Components + unique parts
    # -----------------------------------------------------

    components = assembly.get_components()
    if not components:
        print()
        print("No components found. Stopping.")
        return

    unique_parts = assembly.get_unique_part_quantities(components)

    print_summary(components, unique_parts)

    # -----------------------------------------------------
    # Sheet metal scan + CSV export
    # -----------------------------------------------------

    detector = SheetMetalDetector(
        sw_app, assembly_path=assembly_path, read_thickness=GROUP_BY_THICKNESS
    )

    try:
        result = detector.scan_and_export(unique_parts, output_folder)

    except Exception as e:
        print()
        print("=" * 60)
        print("SCAN FAILED")
        print("=" * 60)
        print(e)
        return

    # -----------------------------------------------------
    # One drawing per thickness
    # -----------------------------------------------------

    drawing_result = None
    exporter = DwgExporter(sw_app, assembly_path=assembly_path)
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
            result["sheet_metal_parts"],
            output_folder / ("by_thickness" if GROUP_BY_THICKNESS else "combined"),
            base_name=assembly_path.stem,
        )

    except Exception as e:
        print()
        print("=" * 60)
        print("DRAWING EXPORT FAILED")
        print("=" * 60)
        print(e)

    finally:
        exporter.activate_assembly(assembly.model)
        detector.close_opened_models()

    # -----------------------------------------------------
    # Final result
    # -----------------------------------------------------

    print()
    print("=" * 60)
    print("PROCESS FINISHED")
    print("=" * 60)

    print(f"Sheet metal (unique parts): {result['sheet_metal']}")
    print(f"Sheet metal (total quantity): {result['total_quantity']}")
    print(f"Not sheet metal: {result['not_sheet_metal']}")
    print(f"Failed: {result['failed']}")

    if drawing_result is not None:
        print(f"Drawings created: {len(drawing_result['files'])}")
        print(f"Parts not exported: {len(drawing_result['failed_parts'])}")

    print()
    print("Output folder:")
    print(output_folder)

    print()


# =========================================================
# ENTRY POINT
# =========================================================

if __name__ == "__main__":
    pythoncom.CoInitialize()

    try:
        main()

    except KeyboardInterrupt:
        print()
        print("Process interrupted by user.")

    except Exception as e:
        print()
        print("=" * 60)
        print("UNEXPECTED ERROR")
        print("=" * 60)
        print(e)

    finally:
        pythoncom.CoUninitialize()
