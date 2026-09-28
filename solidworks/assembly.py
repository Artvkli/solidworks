# assembly.py

from pathlib import Path
import pythoncom
import win32com.client


SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2


class Component:
    def __init__(
        self,
        name,
        path,
        component_type,
        sw_component=None,
        suppression=None,
        lightweight=False,
        configuration=None,
    ):
        self.name = name
        self.path = path
        self.component_type = component_type
        self.sw_component = sw_component
        self.suppression = suppression
        self.lightweight = lightweight
        self.configuration = configuration

    @property
    def is_part(self):
        return self.component_type == "PART"

    @property
    def is_assembly(self):
        return self.component_type == "SUB-ASSEMBLY"

    @property
    def suppressed(self):
        return self.suppression in (0, 1)


class SolidWorksAssembly:
    def __init__(self, sw_app):
        self.sw_app = sw_app
        self.model = None
        self.file_path = None

    # ---------------------------------------------------------
    # COM helper
    # ---------------------------------------------------------

    @staticmethod
    def _read(obj, name, default=None):
        """
        Safely read either:
        - COM method
        - COM property
        """

        if obj is None:
            return default

        try:
            value = getattr(obj, name)
        except Exception:
            return default

        try:
            if callable(value):
                return value()
            return value
        except Exception:
            return default

    # ---------------------------------------------------------
    # OPEN ASSEMBLY
    # ---------------------------------------------------------

    def open(self, file_path):

        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(file_path)

        if file_path.suffix.lower() != ".sldasm":
            raise ValueError("Selected file is not a SolidWorks assembly.")

        self.file_path = str(file_path)

        print()
        print("=" * 60)
        print("OPEN ASSEMBLY")
        print("=" * 60)

        # Check if already open
        try:
            self.model = self.sw_app.GetOpenDocumentByName(str(file_path))
        except Exception:
            self.model = None

        # Open if not already open
        if self.model is None:
            errors = 0
            warnings = 0

            try:
                self.model = self.sw_app.OpenDoc6(
                    str(file_path),
                    SW_DOC_ASSEMBLY,
                    1,
                    "",
                    errors,
                    warnings,
                )
            except Exception as e:
                raise RuntimeError(f"Could not open assembly:\n{e}")

        if self.model is None:
            raise RuntimeError("SolidWorks returned None while opening assembly.")

        # Resolve lightweight components
        try:
            print("Resolving lightweight components...")

            resolve_method = getattr(
                self.model,
                "ResolveAllLightweightComponents",
                None,
            )

            if callable(resolve_method):
                resolve_method(False)

        except Exception as e:
            print(f"Warning resolving lightweight components: {e}")

        print("Assembly opened successfully.")

        return self.model

    # ---------------------------------------------------------
    # GET COMPONENTS
    # ---------------------------------------------------------

    def get_components(self):

        if self.model is None:
            raise RuntimeError("Assembly is not open.")

        print()
        print("=" * 60)
        print("SCANNING COMPONENTS")
        print("=" * 60)

        components = []

        try:
            sw_components = self.model.GetComponents(False)
        except Exception as e:
            raise RuntimeError(f"GetComponents failed: {e}")

        if sw_components is None:
            print("No components found.")
            return components

        for sw_component in sw_components:
            try:
                # ---------------------------------------------
                # Name
                # ---------------------------------------------

                name = self._read(sw_component, "Name2", "")

                if not name:
                    name = self._read(sw_component, "Name", "")

                # ---------------------------------------------
                # Path
                # ---------------------------------------------

                path = self._read(sw_component, "GetPathName", "")

                if path is None:
                    path = ""

                path = str(path)

                # ---------------------------------------------
                # Suppression
                # ---------------------------------------------

                suppression = self._read(sw_component, "GetSuppression", None)

                # ---------------------------------------------
                # Lightweight
                # ---------------------------------------------

                lightweight = False

                try:
                    lightweight_value = self._read(sw_component, "IsLightweight", False)

                    lightweight = bool(lightweight_value)

                except Exception:
                    lightweight = False

                # ---------------------------------------------
                # Referenced configuration
                # ---------------------------------------------

                configuration = self._read(sw_component, "ReferencedConfiguration", "")

                # ---------------------------------------------
                # Determine component type
                # ---------------------------------------------

                extension = Path(path).suffix.lower()

                if extension == ".sldprt":
                    component_type = "PART"

                elif extension == ".sldasm":
                    component_type = "SUB-ASSEMBLY"

                else:
                    # If extension isn't available, try model doc
                    model_doc = None

                    try:
                        get_model = getattr(sw_component, "GetModelDoc2", None)

                        if callable(get_model):
                            model_doc = get_model()

                    except Exception:
                        model_doc = None

                    if model_doc is not None:
                        try:
                            doc_type = self._read(model_doc, "GetType", None)

                            if doc_type == SW_DOC_PART:
                                component_type = "PART"

                            elif doc_type == SW_DOC_ASSEMBLY:
                                component_type = "SUB-ASSEMBLY"

                            else:
                                component_type = "UNKNOWN"

                        except Exception:
                            component_type = "UNKNOWN"

                    else:
                        component_type = "UNKNOWN"

                component = Component(
                    name=name,
                    path=path,
                    component_type=component_type,
                    sw_component=sw_component,
                    suppression=suppression,
                    lightweight=lightweight,
                    configuration=configuration,
                )

                components.append(component)

            except Exception as e:
                print(f"Component error: {e}")

        print()
        print(f"Total components found: {len(components)}")

        return components

    # ---------------------------------------------------------
    # UNIQUE PARTS
    # ---------------------------------------------------------

    def get_unique_part_quantities(self, components=None):

        if components is None:
            components = self.get_components()

        unique = {}

        for component in components:
            if not component.is_part:
                continue

            if not component.path:
                continue

            # IMPORTANT:
            # Configuration is included because one SLDPRT can
            # contain several different configurations.
            key = (component.path.lower(), str(component.configuration or "").lower())

            if key not in unique:
                unique[key] = {
                    "name": component.name,
                    "path": component.path,
                    "configuration": component.configuration,
                    "quantity": 0,
                    "components": [],
                }

            unique[key]["quantity"] += 1
            unique[key]["components"].append(component)

        return list(unique.values())


def find_assemblies(input_folder):

    input_folder = Path(input_folder)

    if not input_folder.exists():
        return []

    return sorted(input_folder.glob("*.SLDASM"))
