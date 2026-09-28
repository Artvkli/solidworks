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

    # Create folders if they don't exist
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

        print(
            f"{index}. {assembly_file.name}"
        )

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
    detector = SheetMetalDetector(sw_app)

    print()
    print("Sheet Metal Detection")
    print("=" * 50)

    sheet_metal_parts = []
    non_sheet_metal_parts = []

    print()
    print("Sheet Metal Detection")
    print("=" * 50)

    for component in components:

        if not component.is_part:
            continue

        if component.suppressed:
            continue

        thickness = detector.get_thickness(component)

        if thickness is not None:
            model = detector.open_part(component.path)
            flat_pattern = detector.find_flat_pattern_feature(model)
            if flat_pattern is not None:
                print(f"[FLAT PATTERN FOUND] {component.name}")

                activated = detector.activate_flat_pattern(model)

                if activated:
                    print(f"[FLAT PATTERN ACTIVE] {component.name}")
                else:
                    print(f"[FLAT PATTERN ACTIVATION FAILED] {component.name}")
            else:
                print(f"[NO FLAT PATTERN] {component.name}")
            sheet_metal_parts.append(component)

            thickness_mm = thickness * 1000

            print(f"[SHEET METAL] {component.name}")
            print(f"Thickness: {thickness_mm:.3f} mm")

        else:

            non_sheet_metal_parts.append(component)

            print(f"[NORMAL PART] {component.name}")

    print()
    print("Sheet Metal Summary:")
    print("Sheet Metal parts:", len(sheet_metal_parts))
    print("Normal parts:", len(non_sheet_metal_parts))
    

    
    
if __name__ == "__main__":
    main()