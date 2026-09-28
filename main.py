
from pathlib import Path

import pythoncom
import win32com.client

from assembly import (
    SolidWorksAssembly,
    find_assemblies,
)

from sheet_metal import (
    SheetMetalDetector,
)


LINE = "=" * 60


# =========================================================
# SOLIDWORKS CONNECTION
# =========================================================

def connect_to_solidworks():

    print()
    print(
        "Connecting to SolidWorks..."
    )

    try:

        sw_app = (
            win32com.client.Dispatch(
                "SldWorks.Application"
            )
        )

        sw_app.Visible = True

        print(
            "Connected successfully."
        )

        return sw_app

    except Exception as e:

        print()
        print(
            "Could not connect to SolidWorks:"
        )

        print(e)

        return None


# =========================================================
# FIND ASSEMBLY
# =========================================================

def select_assembly(
    input_folder
):

    assemblies = (
        find_assemblies(
            input_folder
        )
    )

    if not assemblies:

        print()
        print(
            "No SLDASM files found."
        )

        print()
        print(
            "Put your assembly inside:"
        )

        print(
            input_folder
        )

        return None

    print()
    print(
        "Assemblies found:"
    )

    for index, file in enumerate(
        assemblies,
        start=1
    ):

        print(
            f"{index}. {file.name}"
        )

    # ---------------------------------------------
    # Currently use first assembly
    # ---------------------------------------------

    selected = assemblies[0]

    print()
    print(
        "Selected:"
    )

    print(
        selected
    )

    return selected


# =========================================================
# PRINT COMPONENT SUMMARY
# =========================================================

def print_summary(
    components,
    unique_parts
):

    part_count = 0
    assembly_count = 0

    for component in components:

        if component.is_part:

            part_count += 1

        elif component.is_assembly:

            assembly_count += 1

    print()
    print(LINE)
    print("ASSEMBLY SUMMARY")
    print(LINE)

    print()
    print(
        "All components:",
        len(components)
    )

    print(
        "PART components:",
        part_count
    )

    print(
        "SUB-ASSEMBLIES:",
        assembly_count
    )

    print(
        "Unique PART files:",
        len(unique_parts)
    )


# =========================================================
# MAIN
# =========================================================

def main():

    print()
    print(LINE)
    print(
        "SOLIDWORKS SHEET METAL SCANNER"
    )
    print(LINE)

    # -----------------------------------------------------
    # FOLDERS
    # -----------------------------------------------------

    project_folder = (
        Path(__file__).parent
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

    # -----------------------------------------------------
    # CONNECT
    # -----------------------------------------------------

    sw_app = (
        connect_to_solidworks()
    )

    if sw_app is None:
        return

    # -----------------------------------------------------
    # FIND ASSEMBLY
    # -----------------------------------------------------

    assembly_file = (
        select_assembly(
            input_folder
        )
    )

    if assembly_file is None:
        return

    # -----------------------------------------------------
    # OPEN ASSEMBLY
    # -----------------------------------------------------

    assembly = (
        SolidWorksAssembly(
            sw_app
        )
    )

    if not assembly.open(
        assembly_file
    ):

        print()
        print(
            "Assembly could not be opened."
        )

        return

    # -----------------------------------------------------
    # READ COMPONENTS
    # -----------------------------------------------------

    components = (
        assembly.get_components()
    )

    if not components:

        print()
        print(
            "No components found."
        )

        return

    # -----------------------------------------------------
    # UNIQUE PARTS
    # -----------------------------------------------------

    unique_parts = (
        assembly.get_unique_part_quantities(
            components
        )
    )

    # -----------------------------------------------------
    # SUMMARY
    # -----------------------------------------------------

    print_summary(
        components,
        unique_parts
    )

    # -----------------------------------------------------
    # SHEET METAL DETECTOR
    # -----------------------------------------------------

    detector = (
        SheetMetalDetector(
            sw_app
        )
    )

    # -----------------------------------------------------
    # SCAN + EXPORT
    # -----------------------------------------------------

    result = (
        detector.scan_and_export(
            components,
            unique_parts,
            output_folder
        )
    )

    # -----------------------------------------------------
    # DONE
    # -----------------------------------------------------

    print()
    print()
    print(LINE)
    print(
        "PROCESS FINISHED"
    )
    print(LINE)

    print()
    print(
        "Exported DWG:",
        len(
            result["exported"]
        )
    )

    print(
        "Sheet Metal:",
        len(
            result["sheet_metal"]
        )
    )

    print(
        "Not Sheet Metal:",
        len(
            result["non_sheet_metal"]
        )
    )

    print(
        "Failed:",
        len(
            result["failed"]
        )
    )

    print()
    print(
        "Output folder:"
    )

    print(
        output_folder
    )


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    pythoncom.CoInitialize()

    try:

        main()

    finally:

        pythoncom.CoUninitialize()

