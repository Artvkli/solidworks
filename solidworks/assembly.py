
from pathlib import Path

import pythoncom
import win32com.client


LINE = "=" * 60

SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2


class Component:

    def __init__(
        self,
        name,
        path,
        component_type,
        sw_component=None,
    ):
        self.name = name
        self.path = Path(path)
        self.component_type = component_type
        self.sw_component = sw_component

        self.suppression = None
        self.lightweight = False
        self.configuration = None

    @property
    def is_part(self):
        return self.component_type == "PART"

    @property
    def is_assembly(self):
        return self.component_type == "ASSEMBLY"

    @property
    def suppressed(self):
        return self.suppression == 0


class SolidWorksAssembly:

    def __init__(self, sw_app):

        self.sw_app = sw_app
        self.model = None

    # =====================================================
    # COM HELPERS
    # =====================================================

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

    # =====================================================
    # OPEN ASSEMBLY
    # =====================================================

    def open(self, file_path):

        file_path = Path(file_path)

        if not file_path.exists():

            print(
                "File not found:",
                file_path
            )

            return False

        if file_path.suffix.lower() != ".sldasm":

            print(
                "Not an SLDASM file:",
                file_path
            )

            return False

        print()
        print(LINE)
        print("OPEN ASSEMBLY")
        print(LINE)

        try:

            # ---------------------------------------------
            # Already open?
            # ---------------------------------------------

            try:

                model = (
                    self.sw_app
                    .GetOpenDocumentByName(
                        str(file_path)
                    )
                )

            except Exception:

                model = None

            if model is not None:

                self.model = model

                print(
                    "Assembly already open."
                )

            else:

                errors = win32com.client.VARIANT(
                    pythoncom.VT_BYREF | pythoncom.VT_I4,
                    0
                )

                warnings = win32com.client.VARIANT(
                    pythoncom.VT_BYREF | pythoncom.VT_I4,
                    0
                )

                self.model = (
                    self.sw_app.OpenDoc6(
                        str(file_path),
                        SW_DOC_ASSEMBLY,
                        0,
                        "",
                        errors,
                        warnings,
                    )
                )

                if self.model is None:

                    print(
                        "Could not open assembly."
                    )

                    print(
                        "Errors:",
                        errors.value
                    )

                    print(
                        "Warnings:",
                        warnings.value
                    )

                    return False

            # ---------------------------------------------
            # Resolve lightweight
            # ---------------------------------------------

            print(
                "Resolving lightweight components..."
            )

            try:

                self.model.ResolveAllLightweightComponents(
                    False
                )

            except Exception as e:

                print(
                    "Resolve warning:",
                    e
                )

            print(
                "Assembly opened successfully."
            )

            return True

        except Exception as e:

            print(
                "Assembly open error:"
            )

            print(e)

            return False

    # =====================================================
    # GET COMPONENTS
    # =====================================================

    def get_components(self):

        if self.model is None:
            return []

        components = []

        try:

            print()
            print(LINE)
            print("SCANNING COMPONENTS")
            print(LINE)

            # False = include all levels
            sw_components = (
                self.model.GetComponents(False)
            )

            if sw_components is None:
                return []

            for sw_component in sw_components:

                try:

                    path = (
                        sw_component.GetPathName()
                    )

                    if not path:
                        continue

                    path = Path(path)

                    extension = (
                        path.suffix.lower()
                    )

                    if extension == ".sldprt":

                        component_type = "PART"

                    elif extension == ".sldasm":

                        component_type = "ASSEMBLY"

                    else:

                        continue

                    # -------------------------------------
                    # Suppression
                    # -------------------------------------

                    try:

                        suppression = (
                            sw_component
                            .GetSuppression()
                        )

                    except Exception:

                        suppression = None

                    # -------------------------------------
                    # Lightweight
                    # -------------------------------------

                    try:

                        lightweight = bool(
                            sw_component
                            .IsLightWeight
                        )

                    except Exception:

                        lightweight = False

                    # -------------------------------------
                    # Configuration
                    # -------------------------------------

                    try:

                        configuration = (
                            sw_component
                            .ReferencedConfiguration
                        )

                    except Exception:

                        configuration = None

                    # -------------------------------------
                    # Component
                    # -------------------------------------

                    component = Component(
                        name=(
                            self._get(
                                sw_component,
                                "Name2",
                                path.stem
                            )
                        ),
                        path=path,
                        component_type=component_type,
                        sw_component=sw_component,
                    )

                    component.suppression = (
                        suppression
                    )

                    component.lightweight = (
                        lightweight
                    )

                    component.configuration = (
                        configuration
                    )

                    components.append(
                        component
                    )

                except Exception as e:

                    print(
                        "Component error:",
                        e
                    )

            print()
            print(
                "Components found:",
                len(components)
            )

            return components

        except Exception as e:

            print(
                "Component scan error:",
                e
            )

            return []

    # =====================================================
    # UNIQUE PARTS
    # =====================================================

    def get_unique_part_quantities(
        self,
        components=None
    ):

        if components is None:
            components = self.get_components()

        grouped = {}

        for component in components:

            if not component.is_part:
                continue

            key = (
                str(component.path).lower()
            )

            if key not in grouped:

                grouped[key] = {
                    "name": component.path.stem,
                    "path": component.path,
                    "quantity": 0,
                    "components": [],
                }

            grouped[key]["quantity"] += 1

            grouped[key]["components"].append(
                component
            )

        return list(
            grouped.values()
        )


# =========================================================
# FIND ASSEMBLIES
# =========================================================

def find_assemblies(input_folder):

    input_folder = Path(input_folder)

    if not input_folder.exists():
        return []

    return sorted(
        [
            file
            for file in input_folder.iterdir()
            if (
                file.is_file()
                and file.suffix.lower() == ".sldasm"
                and not file.name.startswith("~$")
            )
        ]
    )

