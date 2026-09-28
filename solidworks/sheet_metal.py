import pythoncom
import win32com.client


class SheetMetalDetector:
    def __init__(self, sw_app):
        self.sw_app = sw_app

    def open_part(self, part_path):
        """Open a SolidWorks part."""

        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        model = self.sw_app.OpenDoc6(str(part_path), 1, 0, "", errors, warnings)

        return model

    def find_sheet_metal_feature(self, model):
        """Find the Sheet Metal feature in a part."""

        feature = model.FirstFeature

        while feature is not None:
            try:
                feature_type = feature.GetTypeName2

                if feature_type:
                    if feature_type.lower() == "sheetmetal":
                        return feature

            except Exception:
                pass

            try:
                feature = feature.GetNextFeature
            except Exception:
                break

        return None

    def find_flat_pattern_feature(self, model):
        """Find the Flat Pattern feature in a part."""
        feature = model.FirstFeature

        while feature is not None:
            try:
                feature_type = feature.GetTypeName2

                if feature_type:
                    if feature_type.lower() == "flatpattern":
                        return feature

            except Exception:
                pass

            try:
                feature = feature.GetNextFeature
            except Exception:
                break

        return None

    def activate_flat_pattern(self, model):
        """Activate the Flat Pattern feature."""
        flat_pattern = self.find_flat_pattern_feature(model)

        if flat_pattern is None:
            print("Flat Pattern feature not found.")
            return False

        try:
            success = flat_pattern.SetSuppression2(1, 2, None)

            if success:
                print("Flat Pattern activated.")
                return True

            print("Could not activate Flat Pattern.")
            return False

        except Exception as e:
            print("Error while activating Flat Pattern:")
            print(e)
            return False

    def is_sheet_metal(self, component):
        """Check whether a component is a Sheet Metal part."""

        if not component.is_part:
            return False

        if component.suppressed:
            return False

        try:
            model = self.open_part(component.path)

            if model is None:
                print(f"Could not open part: {component.name}")
                return False

            feature = self.find_sheet_metal_feature(model)

            return feature is not None

        except Exception as e:
            print(f"Could not inspect part: {component.name}")
            print(e)
            return False

    def get_thickness(self, component):
        """Get Sheet Metal thickness in meters."""

        if not component.is_part:
            return None

        if component.suppressed:
            return None

        try:
            model = self.open_part(component.path)

            if model is None:
                print(f"Could not open part: {component.name}")
                return None

            feature = self.find_sheet_metal_feature(model)

            if feature is None:
                return None

            feature_data = feature.GetDefinition

            if feature_data is None:
                return None

            thickness = feature_data.Thickness

            return thickness

        except Exception as e:
            print(f"Could not read thickness: {component.name}")
            print(e)
            return None
