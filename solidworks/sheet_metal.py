from pathlib import Path

import pythoncom
import win32com.client


LINE = "=" * 60

# SolidWorks constants
SW_DOC_PART = 1
SW_SOLID_BODY = 0
SW_SUPPRESSION_LIGHTWEIGHT = 1
SW_COMPONENT_RESOLVED = 2
SW_FEATURE_UNSUPPRESS = 2


class SheetMetalDetector:
    """Detect sheet metal parts in a SolidWorks assembly and read their data."""

    def __init__(self, sw_app):
        self.sw_app = sw_app
        self._opened_models = {}

    # =========================================================
    # COM Helpers
    # =========================================================

    def _get_com_value(self, obj, name, default=None):
        """
        Safely read a SolidWorks COM property/method.

        pywin32 may expose some SolidWorks API members as
        properties and others as callables depending on the
        COM interface.
        """

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

    def _call_com(self, obj, name, *args):
        """Safely call a SolidWorks COM method."""

        if obj is None:
            return None

        try:
            method = getattr(obj, name)

            if callable(method):
                return method(*args)

            return method

        except Exception:
            return None

    # =========================================================
    # Feature Tree Helpers
    # =========================================================

    def _iter_features(self, model):
        """Yield every top-level feature in the feature tree."""

        if model is None:
            return

        feature = self._get_com_value(model, "FirstFeature")

        while feature is not None:
            yield feature
            feature = self._get_com_value(feature, "GetNextFeature")

    def _find_feature_by_type(self, model, type_name):
        """Return the first feature whose type name matches (case-insensitive)."""

        type_name = type_name.lower()

        for feature in self._iter_features(model):
            feature_type = self._get_com_value(feature, "GetTypeName2")

            if feature_type and str(feature_type).lower() == type_name:
                return feature

        return None

    # =========================================================
    # Open Part
    # =========================================================

    def open_part(self, part_path):
        """
        Open a SLDPRT only when the model cannot be obtained
        directly from the assembly component.
        """

        part_path = Path(part_path)

        if not part_path.exists():
            print()
            print("Part file does not exist:")
            print(part_path)
            return None

        key = str(part_path).lower()

        # Already opened by this detector
        model = self._opened_models.get(key)
        if model is not None:
            return model

        # Already open in SolidWorks
        try:
            model = self.sw_app.GetOpenDocumentByName(str(part_path))

            if model is not None:
                self._opened_models[key] = model
                return model

        except Exception:
            pass

        # Open the part
        try:
            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

            model = self.sw_app.OpenDoc6(
                str(part_path),
                SW_DOC_PART,
                0,
                "",
                errors,
                warnings,
            )

            if model is None:
                print()
                print("Could not open part:")
                print(part_path)
                print("Errors:", errors.value)
                print("Warnings:", warnings.value)
                return None

            self._opened_models[key] = model
            return model

        except Exception as e:
            print()
            print("Exception while opening part:")
            print(part_path)
            print(e)
            return None

    # =========================================================
    # Get Model From Assembly Component
    # =========================================================

    def get_model_from_component(self, component):
        """
        Get the actual ModelDoc2 from the assembly component.

        This is preferred over opening the SLDPRT again.
        """

        if component is None:
            return None

        sw_component = getattr(component, "sw_component", None)

        if sw_component is None:
            return None

        # First choice: GetModelDoc2
        model = self._call_com(sw_component, "GetModelDoc2")

        if model is not None:
            return model

        # If the component is lightweight, try resolving it.
        #
        # Suppression states: 0 = suppressed, 1 = lightweight, 2 = resolved.
        # The value is not blindly assumed to be identical between
        # API versions/configurations.
        try:
            suppression = self._call_com(sw_component, "GetSuppression")

            if suppression == SW_SUPPRESSION_LIGHTWEIGHT:
                print()
                print("Component is lightweight. Trying to resolve:")
                print(component.name)

                result = self._call_com(
                    sw_component,
                    "SetSuppression2",
                    SW_COMPONENT_RESOLVED,
                )

                if result is not None:
                    model = self._call_com(sw_component, "GetModelDoc2")

                    if model is not None:
                        return model

        except Exception as e:
            print()
            print("Could not resolve component:")
            print(component.name)
            print(e)

        # Last resort: open the physical file
        try:
            return self.open_part(component.path)
        except Exception:
            return None

    # =========================================================
    # Bodies
    # =========================================================

    def get_bodies(self, model):
        """Get all solid bodies from a Part document."""

        if model is None:
            return []

        try:
            bodies = self._call_com(model, "GetBodies2", SW_SOLID_BODY, True)

            if bodies is None:
                return []

            return list(bodies)

        except Exception:
            return []

    def find_sheet_metal_body(self, model):
        """Find the first Sheet Metal body in the part."""

        for body in self.get_bodies(model):
            try:
                if bool(self._call_com(body, "IsSheetMetal")):
                    return body
            except Exception:
                continue

        return None

    # =========================================================
    # Sheet Metal Detection
    # =========================================================

    def find_sheet_metal_feature(self, model):
        """
        Find a SheetMetal feature in the feature tree.

        Used as a secondary detection method.
        """

        return self._find_feature_by_type(model, "SheetMetal")

    def get_sheet_metal_feature(self, model):
        """Return the Sheet Metal feature used for thickness data."""

        return self.find_sheet_metal_feature(model)

    def is_sheet_metal(self, component):
        """Determine whether an assembly component is Sheet Metal."""

        if component is None or not component.is_part:
            return False

        model = self.get_model_from_component(component)

        if model is None:
            print()
            print("Could not obtain ModelDoc2:")
            print(component.name)
            return False

        # Primary detection: Body.IsSheetMetal()
        if self.find_sheet_metal_body(model) is not None:
            return True

        # Secondary detection: Feature Tree
        return self.find_sheet_metal_feature(model) is not None

    # =========================================================
    # Thickness
    # =========================================================

    def get_thickness(self, component):
        """
        Return sheet metal thickness in meters.

        SolidWorks internally uses meters.
        """

        if component is None or not component.is_part:
            return None

        model = self.get_model_from_component(component)

        if model is None:
            return None

        feature = self.get_sheet_metal_feature(model)

        if feature is None:
            print()
            print("Sheet Metal body detected, but Sheet Metal feature was not found:")
            print(component.name)
            return None

        definition = self._call_com(feature, "GetDefinition")

        if definition is None:
            return None

        try:
            thickness = self._get_com_value(definition, "Thickness")

            if thickness is None:
                return None

            thickness = float(thickness)

            return thickness if thickness > 0 else None

        except Exception as e:
            print()
            print("Could not read thickness:")
            print(component.name)
            print(e)
            return None

    # =========================================================
    # Flat Pattern
    # =========================================================

    def find_flat_pattern_feature(self, model):
        """Find the FlatPattern feature."""

        return self._find_feature_by_type(model, "FlatPattern")

    def activate_flat_pattern(self, model):
        """Unsuppress/activate the Flat Pattern feature."""

        flat_pattern = self.find_flat_pattern_feature(model)

        if flat_pattern is None:
            print()
            print("Flat Pattern feature not found.")
            return False

        # SetSuppression2 parameters can vary depending on the API
        # interface, so try the full signature first, then the short one.
        attempts = (
            (SW_FEATURE_UNSUPPRESS, 2, None),
            (SW_FEATURE_UNSUPPRESS,),
        )

        for args in attempts:
            try:
                if self._call_com(flat_pattern, "SetSuppression2", *args):
                    return True
            except Exception as e:
                print()
                print("Could not activate Flat Pattern:")
                print(e)

        return False

    # =========================================================
    # Export DWG / DXF
    # =========================================================

    def export_dxf(self, model, output_path):
        """Export the active flat pattern to DWG/DXF."""

        if model is None:
            return False

        output_path = Path(output_path)

        try:
            model_path = self._get_com_value(model, "GetPathName")

            if not model_path:
                return False

            result = self._call_com(
                model,
                "ExportToDWG2",
                str(output_path),
                str(model_path),
                1,
                True,
                None,
                False,
                False,
                0,
                None,
            )

            return bool(result)

        except Exception as e:
            print()
            print("DWG/DXF export failed:")
            print(output_path)
            print(e)
            return False

    # =========================================================
    # Debug
    # =========================================================

    def debug_features(self, model):
        """Print the complete top-level feature tree."""

        if model is None:
            print("Model is None.")
            return

        print()
        print("Feature Tree:")
        print(LINE)

        count = 0

        for count, feature in enumerate(self._iter_features(model), start=1):
            name = self._get_com_value(feature, "Name", "?")
            feature_type = self._get_com_value(feature, "GetTypeName2", "?")
            print(f"{count}. {name} | {feature_type}")

        print(LINE)
        print("Feature count:", count)

    def debug_bodies(self, model):
        """Print all solid bodies and their Sheet Metal state."""

        if model is None:
            print("Model is None.")
            return

        bodies = self.get_bodies(model)

        print()
        print("Bodies:")
        print(LINE)
        print("Body count:", len(bodies))

        for index, body in enumerate(bodies, start=1):
            name = self._get_com_value(body, "Name", "?")
            sheet_metal = self._call_com(body, "IsSheetMetal")
            print(f"{index}. {name} | SheetMetal={bool(sheet_metal)}")

        print(LINE)
