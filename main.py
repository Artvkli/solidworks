# main.py

from pathlib import Path

import pythoncom
import win32com.client

from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.dwg_exporter import DwgExporter
from solidworks.sheet_metal import SheetMetalDetector


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

    detector = SheetMetalDetector(sw_app, assembly_path=assembly_path)

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
    # DWG export (flat patterns)
    # -----------------------------------------------------

    dwg_result = None
    exporter = DwgExporter(sw_app, assembly_path=assembly_path)

    try:
        dwg_result = exporter.export_all(
            result["sheet_metal_parts"], output_folder / "DWG"
        )

    except Exception as e:
        print()
        print("=" * 60)
        print("DWG EXPORT FAILED")
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

    if dwg_result is not None:
        print(f"DWG exported: {dwg_result['exported']}")
        print(f"DWG failed: {dwg_result['failed']}")

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
