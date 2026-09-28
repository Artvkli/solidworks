from pathlib import Path

from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly, find_assemblies


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

    print()
    print("Number of components:", len(components))

    print()
    print("Components:")

    for component in components:

        try:

            print("-" * 40)
            print("Name:", component.Name2)
            print("Path:", component.GetPathName())

        except Exception as e:

            print("Could not read component:")
            print(e)

    print()
    print("Program finished.")


if __name__ == "__main__":
    main()