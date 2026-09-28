# main.py

from pathlib import Path
import pythoncom
import win32com.client

from solidworks.assembly import SolidWorksAssembly, find_assemblies
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

    # -----------------------------------------------------
    # Currently use the first assembly automatically
    # -----------------------------------------------------

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

    # -----------------------------------------------------
    # Project folders
    # -----------------------------------------------------

    project_folder = Path(__file__).resolve().parent

    input_folder = project_folder / "input"

    output_folder = project_folder / "output"

    # Create folders if necessary
    input_folder.mkdir(parents=True, exist_ok=True)

    output_folder.mkdir(parents=True, exist_ok=True)

    # -----------------------------------------------------
    # Connect to SolidWorks
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

    try:
        assembly.open(assembly_path)

    except Exception as e:
        print()
        print(f"Assembly open failed:\n{e}")

        return

    # -----------------------------------------------------
    # Get all components
    # -----------------------------------------------------

    try:
        components = assembly.get_components()

    except Exception as e:
        print()
        print(f"Component scan failed:\n{e}")

        return

    # -----------------------------------------------------
    # Group unique parts
    # -----------------------------------------------------

    try:
        unique_parts = assembly.get_unique_part_quantities(components)

    except Exception as e:
        print()
        print(f"Unique part grouping failed:\n{e}")

        return

    # -----------------------------------------------------
    # Summary
    # -----------------------------------------------------

    print_summary(components, unique_parts)

    # -----------------------------------------------------
    # Sheet Metal Detector
    #
    # IMPORTANT:
    # Pass assembly_path so the detector can search for
    # missing/old component paths next to the assembly.
    # -----------------------------------------------------

    detector = SheetMetalDetector(sw_app, assembly_path=assembly_path)

    # -----------------------------------------------------
    # Scan + Export DWG
    # -----------------------------------------------------

    try:
        result = detector.scan_and_export(components, unique_parts, output_folder)

    except Exception as e:
        print()
        print("=" * 60)
        print("SCAN FAILED")
        print("=" * 60)

        print(e)

        return

    # -----------------------------------------------------
    # Final result
    # -----------------------------------------------------

    print()
    print("=" * 60)
    print("PROCESS FINISHED")
    print("=" * 60)

    print(f"Exported DWG: {result['exported']}")

    print(f"Sheet Metal: {result['sheet_metal']}")

    print(f"Not Sheet Metal: {result['not_sheet_metal']}")

    print(f"Failed: {result['failed']}")

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
