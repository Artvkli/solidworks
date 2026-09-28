import pythoncom
import win32com.client


class SheetMetalDetector:
    def __init__(self, sw_app):
        self.sw_app = sw_app

    def open_part(self, part_path):
        """Open a SolidWorks part."""

        errors = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0
        )

        warnings = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0
        )

        model = self.sw_app.OpenDoc6(
            str(part_path),
            1,
            0,
            "",
            errors,
            warnings
        )

        return model

    def is_sheet_metal(self, component):
        """
        Check whether a SolidWorks part is a Sheet Metal part.
        """

        if not component.is_part:
            return False

        if component.suppressed:
            return False

        try:
            model = self.open_part(component.path)

            if model is None:
                print(f"Could not open part: {component.name}")
                return False

            feature = model.FirstFeature

            while feature is not None:

                try:
                    feature_type = feature.GetTypeName2

                    if feature_type:
                        feature_type = feature_type.lower()

                        if feature_type == "sheetmetal":
                            return True

                except Exception:
                    pass

                try:
                    feature = feature.GetNextFeature
                except Exception:
                    break

            return False

        except Exception as e:
            print(f"Could not inspect part: {component.name}")
            print(e)
            return False