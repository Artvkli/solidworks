from pathlib import Path

import pythoncom
import win32com.client

from solidworks.component import Component


class SolidWorksAssembly:
    def __init__(self, sw_app):
        self.sw_app = sw_app
        self.model = None

    def _get(self, obj, name, default=None):
        if obj is None:
            return default

        try:
            value = getattr(obj, name)

            if callable(value):
                try:
                    return value()
                except TypeError:
                    return value

            return value

        except Exception:
            return default

    def open(self, file_path):
        file_path = Path(file_path)

        if not file_path.exists():
            print()
            print("Assembly file not found:")
            print(file_path)
            return False

        if file_path.suffix.lower() != ".sldasm":
            print()
            print("Selected file is not a SolidWorks Assembly.")
            return False

        if file_path.name.startswith("~$"):
            print()
            print("Temporary SolidWorks file detected:")
            print(file_path.name)
            return False

        print()
        print("=" * 60)
        print("OPEN ASSEMBLY")
        print("=" * 60)
        print()
        print("Path:")
        print(file_path)

        try:
            existing_model = self.sw_app.GetOpenDocumentByName(str(file_path))

            if existing_model is not None:
                self.model = existing_model

                print()
                print("Assembly is already open.")

                title = self._get(self.model, "GetTitle", file_path.stem)

                print("Title:", title)

                return self._prepare_assembly()

        except Exception as e:
            print()
            print("Could not check existing document.")
            print("Continuing with OpenDoc6...")
            print("Details:", e)

        try:
            document_type = 2
            options = 0

            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

            warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

            print()
            print("Opening with OpenDoc6...")

            model = self.sw_app.OpenDoc6(
                str(file_path), document_type, options, "", errors, warnings
            )

            self.model = model

            if self.model is None:
                print()
                print("SolidWorks returned no model.")
                print("OpenDoc6 Errors:", errors.value)
                print("OpenDoc6 Warnings:", warnings.value)
                return False

            print()
            print("Assembly opened successfully.")
            print("OpenDoc6 Errors:", errors.value)
            print("OpenDoc6 Warnings:", warnings.value)

            title = self._get(self.model, "GetTitle", file_path.stem)

            print("Title:", title)

            return self._prepare_assembly()

        except Exception as e:
            print()
            print("Assembly open failed:")
            print(e)
            return False

    def _prepare_assembly(self):
        if self.model is None:
            return False

        print()
        print("Resolving lightweight components...")

        try:
            result = self.model.ResolveAllLightweightComponents(False)
            print("Resolve result:", result)

        except Exception as e:
            print("Could not resolve lightweight components:")
            print(e)

        return True

    def get_components(self):
        if self.model is None:
            print()
            print("No assembly is open.")
            return []

        print()
        print("=" * 60)
        print("READING ASSEMBLY COMPONENTS")
        print("=" * 60)

        try:
            sw_components = self.model.GetComponents(False)

        except Exception as e:
            print()
            print("GetComponents failed:")
            print(e)
            return []

        if sw_components is None:
            print()
            print("SolidWorks returned no components.")
            return []

        components = []

        for sw_component in sw_components:
            try:
                name = self._get(sw_component, "Name2", "Unknown")

                path = self._get(sw_component, "GetPathName", None)

                if not path:
                    continue

                path = Path(str(path))
                extension = path.suffix.lower()

                if extension == ".sldprt":
                    component_type = "PART"
                elif extension == ".sldasm":
                    component_type = "ASSEMBLY"
                else:
                    continue

                try:
                    suppression = sw_component.GetSuppression
                    if callable(suppression):
                        suppression = suppression()
                except Exception:
                    suppression = None

                try:
                    lightweight = sw_component.IsLightWeight
                    if callable(lightweight):
                        lightweight = lightweight()
                    lightweight = bool(lightweight)
                except Exception:
                    lightweight = False

                try:
                    configuration = sw_component.ReferencedConfiguration
                    if callable(configuration):
                        configuration = configuration()
                except Exception:
                    configuration = None

                component = Component(
                    name=str(name),
                    path=path,
                    component_type=component_type,
                    suppressed=False,
                )

                component.sw_component = sw_component
                component.suppression = suppression
                component.lightweight = lightweight
                component.configuration = configuration

                components.append(component)

            except Exception as e:
                print()
                print("Could not process component:")
                print(e)

        print()
        print("Components collected:", len(components))

        return components

    def get_part_components(self, components=None):
        if components is None:
            components = self.get_components()

        return [component for component in components if component.is_part]

    def get_assembly_components(self, components=None):
        if components is None:
            components = self.get_components()

        return [component for component in components if component.is_assembly]

    def get_unique_part_quantities(self, components=None):
        if components is None:
            components = self.get_components()

        grouped = {}

        for component in components:
            if not component.is_part:
                continue

            key = str(component.path).lower()

            if key not in grouped:
                grouped[key] = {
                    "name": component.path.stem,
                    "path": component.path,
                    "quantity": 0,
                    "components": [],
                }

            grouped[key]["quantity"] += 1
            grouped[key]["components"].append(component)

        return list(grouped.values())


def find_assemblies(input_folder):
    input_folder = Path(input_folder)

    if not input_folder.exists():
        print()
        print("Input folder does not exist:")
        print(input_folder)
        return []

    assemblies = []

    for file in input_folder.iterdir():
        if not file.is_file():
            continue

        if file.name.startswith("~$"):
            continue

        if file.name.startswith("."):
            continue

        if file.suffix.lower() != ".sldasm":
            continue

        assemblies.append(file)

    return sorted(assemblies, key=lambda p: p.name.lower())
