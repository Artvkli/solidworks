from pathlib import Path

import pythoncom
import win32com.client
from solidworks.component import Component

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

        print()
        print("Opening assembly:")
        print(file_path)

        try:
            # SolidWorks document type
            # 2 = Assembly
            document_type = 2

            # Open options
            options = 0

            # OpenDoc6 expects Errors and Warnings
            # as ByRef 32-bit integers.
            errors = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4,
                0
            )

            warnings = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4,
                0
            )

            self.model = self.sw_app.OpenDoc6(
                str(file_path),
                document_type,
                options,
                "",
                errors,
                warnings
            )

            if self.model is None:
                print("SolidWorks did not return a model.")

                print("Errors:", errors.value)
                print("Warnings:", warnings.value)

                return False

            print()
            print("Assembly opened successfully!")
            print("Title:", self.model.GetTitle)

            print("Errors:", errors.value)
            print("Warnings:", warnings.value)

            return True

        except Exception as e:

            print()
            print("Error while opening assembly:")
            print(e)

            return False

    def get_components(self):

        if self.model is None:
            print("No assembly is open.")
            return []

        try:
            sw_components = self.model.GetComponents(True)

            if sw_components is None:
                return []

            components = []

            for sw_component in sw_components:

                try:
                    name = sw_component.Name2
                    path = sw_component.GetPathName

                    if not path:
                        continue

                    path = Path(path)

                    extension = path.suffix.lower()

                    if extension == ".sldprt":
                        component_type = "PART"

                    elif extension == ".sldasm":
                        component_type = "ASSEMBLY"

                    else:
                        continue

                    suppressed = False

                    try:
                        suppressed = sw_component.IsSuppressed
                    except Exception:
                        pass

                    component = Component(
                        name=name,
                        path=path,
                        component_type=component_type,
                        suppressed=suppressed
                    )

                    components.append(component)

                except Exception as e:
                    print("Could not process component:")
                    print(e)

            return components

        except Exception as e:
            print("Error while reading components:")
            print(e)
            return []




    def get_unique_part_quantities(self):
        
        """Group active part instances by their source file."""
        components = self.get_components()

        grouped = {}

        for component in components:
            if not component.is_part:
                continue

            if component.suppressed:
                continue

            key = str(component.path).lower()

            if key not in grouped:
                grouped[key] = {
                    "name": component.path.stem,
                    "path": component.path,
                    "quantity": 0,
                }

            grouped[key]["quantity"] += 1

        return list(grouped.values())




def find_assemblies(input_folder):
    """Find all SolidWorks assemblies inside input folder."""

    input_folder = Path(input_folder)

    if not input_folder.exists():
        print(f"Input folder does not exist: {input_folder}")
        return []

    assemblies = [
        file
        for file in input_folder.iterdir()
        if file.is_file()
        and file.suffix.lower() == ".sldasm"
        and not file.name.startswith("~$")
    ]

    return sorted(assemblies)