
from pathlib import Path

import pythoncom
import win32com.client

from solidworks.component import Component
from solidworks.utils import com_value


class SolidWorksAssembly:

    def __init__(self, sw_app):
        self.sw_app = sw_app
        self.model = None

    # --------------------------------------------------
    # Open Assembly
    # --------------------------------------------------

    def open(self, file_path):
        """Open a SolidWorks assembly."""

        file_path = Path(file_path)

        if not file_path.exists():
            print(f"File not found: {file_path}")
            return False

        if file_path.suffix.lower() != ".sldasm":
            print(
                "Selected file is not a SolidWorks Assembly."
            )
            return False

        print()
        print("Opening assembly:")
        print(file_path)

        try:

            # ------------------------------------------
            # Check if already open
            # ------------------------------------------

            existing_model = None

            try:
                existing_model = (
                    self.sw_app.GetOpenDocumentByName(
                        str(file_path)
                    )
                )
            except Exception:
                existing_model = None

            if existing_model is not None:

                print()
                print(
                    "Assembly is already open "
                    "in SolidWorks."
                )

                self.model = existing_model

            else:

                # --------------------------------------
                # Open document
                # --------------------------------------

                document_type = 2
                options = 0

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

                    print()
                    print(
                        "SolidWorks did not return "
                        "a model."
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

                print()
                print(
                    "Assembly opened successfully!"
                )

                print(
                    "Errors:",
                    errors.value
                )

                print(
                    "Warnings:",
                    warnings.value
                )

            # ------------------------------------------
            # Title
            # ------------------------------------------

            print(
                "Title:",
                com_value(
                    self.model,
                    "GetTitle"
                )
            )

            # ------------------------------------------
            # Resolve lightweight components
            # ------------------------------------------

            print()
            print(
                "Resolving lightweight components..."
            )

            try:

                result = (
                    self.model
                    .ResolveAllLightweightComponents(
                        False
                    )
                )

                print(
                    "Resolve result:",
                    result
                )

            except Exception as e:

                print(
                    "Could not resolve "
                    "lightweight components:"
                )

                print(e)

            return True

        except Exception as e:

            print()
            print(
                "Error while opening assembly:"
            )

            print(e)

            return False

    # --------------------------------------------------
    # Get All Components
    # --------------------------------------------------

    def get_components(self):
        """
        Return all components from all levels
        of the assembly.
        """

        if self.model is None:

            print(
                "No assembly is open."
            )

            return []

        try:

            print()
            print(
                "Reading assembly components..."
            )

            # False means:
            # include components from
            # subassemblies as well.
            sw_components = (
                self.model.GetComponents(False)
            )

            if sw_components is None:

                print(
                    "SolidWorks returned "
                    "no components."
                )

                return []

            components = []

            for sw_component in sw_components:

                try:

                    # ----------------------------------
                    # Name
                    # ----------------------------------

                    name = com_value(
                        sw_component,
                        "Name2"
                    )

                    # ----------------------------------
                    # Path
                    # ----------------------------------

                    path = com_value(
                        sw_component,
                        "GetPathName"
                    )

                    if not path:
                        continue

                    path = Path(path)

                    # ----------------------------------
                    # Component Type
                    # ----------------------------------

                    extension = (
                        path.suffix.lower()
                    )

                    if extension == ".sldprt":

                        component_type = "PART"

                    elif extension == ".sldasm":

                        component_type = "ASSEMBLY"

                    else:

                        continue

                    # ----------------------------------
                    # Suppression
                    # ----------------------------------

                    try:

                        suppression = (
                            sw_component
                            .GetSuppression
                        )

                    except Exception:

                        suppression = None

                    # ----------------------------------
                    # Lightweight
                    # ----------------------------------

                    try:

                        lightweight = bool(
                            sw_component
                            .IsLightWeight
                        )

                    except Exception:

                        lightweight = False

                    # ----------------------------------
                    # Referenced configuration
                    # ----------------------------------

                    try:

                        configuration = (
                            sw_component
                            .ReferencedConfiguration
                        )

                    except Exception:

                        configuration = None

                    # ----------------------------------
                    # Component object
                    # ----------------------------------

                    component = Component(
                        name=name,
                        path=path,
                        component_type=component_type,
                        suppressed=(
                            suppression == 0
                        )
                    )

                    # Attach additional runtime
                    # information.

                    component.sw_component = (
                        sw_component
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

                    print()
                    print(
                        "Could not process "
                        "component:"
                    )

                    print(e)

            print()
            print(
                "Components collected:",
                len(components)
            )

            return components

        except Exception as e:

            print()
            print(
                "Error while reading "
                "components:"
            )

            print(e)

            return []

    # --------------------------------------------------
    # Part Components
    # --------------------------------------------------

    def get_part_components(
        self,
        components=None
    ):
        """
        Return all PART components.
        """

        if components is None:

            components = (
                self.get_components()
            )

        parts = []

        for component in components:

            if not component.is_part:
                continue

            parts.append(component)

        return parts

    # --------------------------------------------------
    # Assembly Components
    # --------------------------------------------------

    def get_assembly_components(
        self,
        components=None
    ):
        """
        Return all SUBASSEMBLY components.
        """

        if components is None:

            components = (
                self.get_components()
            )

        assemblies = []

        for component in components:

            if not component.is_assembly:
                continue

            assemblies.append(component)

        return assemblies

    # --------------------------------------------------
    # Unique Parts
    # --------------------------------------------------

    def get_unique_part_quantities(
        self,
        components=None
    ):
        """
        Group PART instances by source file.

        Quantity = number of occurrences
        in the complete assembly tree.
        """

        if components is None:

            components = (
                self.get_components()
            )

        grouped = {}

        for component in components:

            if not component.is_part:
                continue

            key = str(
                component.path
            ).lower()

            if key not in grouped:

                grouped[key] = {
                    "name": (
                        component.path.stem
                    ),
                    "path": (
                        component.path
                    ),
                    "quantity": 0,
                    "components": [],
                }

            grouped[key][
                "quantity"
            ] += 1

            grouped[key][
                "components"
            ].append(component)

        return list(
            grouped.values()
        )


# ======================================================
# Find Assemblies
# ======================================================

def find_assemblies(input_folder):

    input_folder = Path(
        input_folder
    )

    if not input_folder.exists():

        print(
            f"Input folder does not exist: "
            f"{input_folder}"
        )

        return []

    assemblies = [

        file

        for file in input_folder.iterdir()

        if (
            file.is_file()
            and file.suffix.lower()
            == ".sldasm"
            and not file.name.startswith("~$")
        )
    ]

    return sorted(assemblies)

