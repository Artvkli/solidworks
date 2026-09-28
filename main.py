from pathlib import Path

from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.sheet_metal import SheetMetalDetector


def main():

    print("=" * 50)
    print("SolidWorks Automation")
    print("=" * 50)

    # ----------------------------------------
    # Project paths
    # ----------------------------------------

    project_folder = Path(__file__).parent

    input_folder = project_folder / "input"
    output_folder = project_folder / "output"

    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)

    # ----------------------------------------
    # Connect to SolidWorks
    # ----------------------------------------

    connection = SolidWorksConnection()

    if not connection.connect():
        print()
        print("Program stopped.")
        return

    sw_app = connection.get_application()

    # ----------------------------------------
    # Find assemblies
    # ----------------------------------------

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

    # ----------------------------------------
    # Open first assembly
    # ----------------------------------------

    assembly_file = assemblies[0]

    print()
    print("Opening assembly:")
    print(assembly_file.name)

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_file):
        return

    # ----------------------------------------
    # Get components
    # ----------------------------------------

    components = assembly.get_components()

    print()
    print("Total components:", len(components))

    # ----------------------------------------
    # Sheet Metal Detection
    # ----------------------------------------

    detector = SheetMetalDetector(sw_app)

    print()
    print("Sheet Metal Parts")
    print("=" * 50)

    sheet_metal_data = []

    # Get unique parts and their quantities
    for item in assembly.get_unique_part_quantities():

        component = next(
            (
                c
                for c in components
                if c.is_part
                and not c.suppressed
                and str(c.path).lower() == str(item["path"]).lower()
            ),
            None,
        )

        if component is None:
            continue

        # Read Sheet Metal thickness
        thickness = detector.get_thickness(component)

        if thickness is None:
            continue

        thickness_mm = thickness * 1000

        sheet_metal_data.append(
            {
                "name": item["name"],
                "path": item["path"],
                "quantity": item["quantity"],
                "thickness": thickness_mm,
            }
        )

        print(
            f"{item['name']} | "
            f"Thickness: {thickness_mm:.3f} mm | "
            f"Quantity: {item['quantity']}"
        )

    # ----------------------------------------
    # Summary
    # ----------------------------------------

    print()
    print("=" * 50)
    print("Sheet Metal Summary")
    print("=" * 50)

    print(
        "Unique Sheet Metal parts:",
        len(sheet_metal_data)
    )

    print()
    print("Program finished.")


if __name__ == "__main__":
    main()
