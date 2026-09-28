from pathlib import Path


class SolidWorksAssembly:

    def __init__(self, sw_app):
        self.sw_app = sw_app
        self.model = None

    def open(self, file_path):
        """Open a SolidWorks assembly."""

        file_path = Path(file_path)

        if not file_path.exists():
            print(f"File not found: {file_path}")
            return False

        if file_path.suffix.lower() != ".sldasm":
            print("Selected file is not a SolidWorks Assembly.")
            return False

        try:

            # SolidWorks document type
            # 2 = Assembly
            document_type = 2

            # Open options
            options = 0

            self.model = self.sw_app.OpenDoc6(
                str(file_path),
                document_type,
                options,
                "",
                0,
                0
            )

            if self.model is None:
                print("Could not open assembly.")
                return False

            print()
            print("Assembly opened successfully!")
            print("File:", file_path)
            print("Title:", self.model.GetTitle())

            return True

        except Exception as e:

            print("Error while opening assembly:")
            print(e)

            return False

    def get_components(self):
        """Get all components from the assembly."""

        if self.model is None:
            print("No assembly is open.")
            return []

        try:

            components = self.model.GetComponents(True)

            if components is None:
                return []

            return list(components)

        except Exception as e:

            print("Error while reading components:")
            print(e)

            return []


def find_assemblies(input_folder):
    """Find SolidWorks assemblies inside input folder."""

    input_folder = Path(input_folder)

    if not input_folder.exists():
        print(f"Input folder does not exist: {input_folder}")
        return []

    assemblies = [
        file
        for file in input_folder.iterdir()
        if file.is_file() and file.suffix.lower() == ".sldasm"
    ]

    return sorted(assemblies)