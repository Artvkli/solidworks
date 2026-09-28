from pathlib import Path

from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.sheet_metal import SheetMetalDetector


def format_thickness(thickness_mm):
    """
    Convert thickness to a clean folder/file name.

    Examples:
        1.500 -> 1.5mm
        2.000 -> 2mm
        15.000 -> 15mm
    """
    value = f"{thickness_mm:.3f}".rstrip("0").rstrip(".")
    return f"{value}mm"


def main():

    print("=" * 60)
    print("SolidWorks Automation")
    print("=" * 60)

    # --------------------------------------------------
    # Project paths
    # --------------------------------------------------

    project_folder = Path(__file__).parent

    input_folder = project_folder / "input"
    output_folder = project_folder / "output"
    temp_folder = output_folder / "_temp"

    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)
    temp_folder.mkdir(exist_ok=True)

    # --------------------------------------------------
    # Connect to SolidWorks
    # --------------------------------------------------

    connection = SolidWorksConnection()

    if not connection.connect():
        print()
        print("Program stopped.")
        return

    sw_app = connection.get_application()

    # --------------------------------------------------
    # Find assemblies
    # --------------------------------------------------

    assemblies = find_assemblies(input_folder)

    if not assemblies:
        print()
        print("No SLDASM files found.")
        print()
        print("Please put a SolidWorks Assembly inside:")
        print(input_folder)
        return

    print()
    print("Assemblies found:")

    for index, assembly_file in enumerate(assemblies, start=1):
        print(f"{index}. {assembly_file.name}")

    # --------------------------------------------------
    # Open first assembly
    # --------------------------------------------------

    assembly_file = assemblies[0]

    print()
    print("Opening assembly:")
    print(assembly_file.name)

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_file):
        return

    # --------------------------------------------------
    # Get components
    # --------------------------------------------------

    components = assembly.get_components()

    print()
    print("Total components:", len(components))

    # --------------------------------------------------
    # Sheet Metal Detector
    # --------------------------------------------------

    detector = SheetMetalDetector(sw_app)

    print()
    print("=" * 60)
    print("Collecting Sheet Metal Parts")
    print("=" * 60)

    # --------------------------------------------------
    # Collect unique Sheet Metal parts
    # --------------------------------------------------

    sheet_metal_data = []

    for item in assembly.get_unique_part_quantities():

        component = next(
            (
                c
                for c in components
                if c.is_part
                and not c.suppressed
                and str(c.path).lower()
                == str(item["path"]).lower()
            ),
            None,
        )

        if component is None:
            continue

        # ----------------------------------------------
        # Get thickness
        # ----------------------------------------------

        thickness = detector.get_thickness(component)

        if thickness is None:
            continue

        thickness_mm = thickness * 1000

        # ----------------------------------------------
        # Store information
        # ----------------------------------------------

        data = {
            "name": item["name"],
            "path": item["path"],
            "quantity": item["quantity"],
            "thickness": thickness_mm,
        }

        sheet_metal_data.append(data)

        print(
            f"{item['name']} | "
            f"{thickness_mm:.3f} mm | "
            f"QTY: {item['quantity']}"
        )

    # --------------------------------------------------
    # Check result
    # --------------------------------------------------

    if not sheet_metal_data:
        print()
        print("No Sheet Metal parts found.")
        return

    # --------------------------------------------------
    # Group by thickness
    # --------------------------------------------------

    groups = {}

    for item in sheet_metal_data:

        # Round to avoid floating-point differences
        thickness_key = round(item["thickness"], 3)

        if thickness_key not in groups:
            groups[thickness_key] = []

        groups[thickness_key].append(item)

    # --------------------------------------------------
    # Show groups
    # --------------------------------------------------

    print()
    print("=" * 60)
    print("Sheet Metal Groups")
    print("=" * 60)

    for thickness, parts in sorted(groups.items()):

        thickness_name = format_thickness(thickness)

        print()
        print(f"[{thickness_name}]")

        for part in parts:
            print(
                f"  {part['name']} | "
                f"QTY: {part['quantity']}"
            )

    # --------------------------------------------------
    # Export each unique part
    # into its thickness temporary folder
    # --------------------------------------------------

    print()
    print("=" * 60)
    print("Exporting Flat Patterns")
    print("=" * 60)

    for thickness, parts in sorted(groups.items()):

        thickness_name = format_thickness(thickness)

        thickness_folder = temp_folder / thickness_name
        thickness_folder.mkdir(parents=True, exist_ok=True)

        print()
        print("-" * 60)
        print(f"Thickness Group: {thickness_name}")
        print("-" * 60)

        for part in parts:

            print()
            print(f"Part: {part['name']}")
            print(f"QTY: {part['quantity']}")

            # ------------------------------------------
            # Open part
            # ------------------------------------------

            model = detector.open_part(part["path"])

            if model is None:
                print("[FAILED] Could not open part.")
                continue

            # ------------------------------------------
            # Find Flat Pattern
            # ------------------------------------------

            flat_pattern = detector.find_flat_pattern_feature(
                model
            )

            if flat_pattern is None:
                print("[NO FLAT PATTERN]")
                continue

            print("[FLAT PATTERN FOUND]")

            # ------------------------------------------
            # Activate Flat Pattern
            # ------------------------------------------

            activated = detector.activate_flat_pattern(model)

            if not activated:
                print("[FLAT PATTERN ACTIVATION FAILED]")
                continue

            print("[FLAT PATTERN ACTIVE]")

            # ------------------------------------------
            # Temporary DWG path
            # ------------------------------------------

            output_file = (
                thickness_folder
                / f"{part['name']}.dwg"
            )

            # ------------------------------------------
            # Export
            # ------------------------------------------

            exported = detector.export_dxf(
                model,
                output_file
            )

            if exported:
                print(
                    f"[EXPORT SUCCESS] "
                    f"{output_file}"
                )
            else:
                print(
                    f"[EXPORT FAILED] "
                    f"{part['name']}"
                )

    # --------------------------------------------------
    # Final summary
    # --------------------------------------------------

    print()
    print("=" * 60)
    print("Export Summary")
    print("=" * 60)

    for thickness, parts in sorted(groups.items()):

        thickness_name = format_thickness(thickness)

        thickness_folder = temp_folder / thickness_name

        print()
        print(
            f"{thickness_name} "
            f"-> {len(parts)} unique part(s)"
        )

        for part in parts:

            expected_file = (
                thickness_folder
                / f"{part['name']}.dwg"
            )

            if expected_file.exists():
                print(
                    f"  [OK] {part['name']} "
                    f"| QTY: {part['quantity']}"
                )
            else:
                print(
                    f"  [MISSING] {part['name']} "
                    f"| QTY: {part['quantity']}"
                )

    print()
    print("=" * 60)
    print("Program finished.")
    print("=" * 60)


if __name__ == "__main__":
    main()
