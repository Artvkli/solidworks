
import pythoncom
import win32com.client


class SheetMetalDetector:

    def __init__(self, sw_app):
        self.sw_app = sw_app

    # ==================================================
    # Open Part
    # ==================================================

    def open_part(self, part_path):
        """
        Open a SolidWorks part.

        If the part is already open, use the
        existing document.
        """

        try:

            part_path = str(part_path)

            # ------------------------------------------
            # Check existing document
            # ------------------------------------------

            try:

                existing_model = (
                    self.sw_app.GetOpenDocumentByName(
                        part_path
                    )
                )

                if existing_model is not None:

                    return existing_model

            except Exception:
                pass

            # ------------------------------------------
            # Open part
            # ------------------------------------------

            errors = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4,
                0
            )

            warnings = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4,
                0
            )

            model = self.sw_app.OpenDoc6(
                part_path,
                1,          # swDocPART
                0,
                "",
                errors,
                warnings
            )

            if model is None:

                print()
                print(
                    "Could not open part:"
                )

                print(part_path)

                print(
                    "Errors:",
                    errors.value
                )

                print(
                    "Warnings:",
                    warnings.value
                )

                return None

            return model

        except Exception as e:

            print()
            print(
                "Error opening part:"
            )

            print(part_path)
            print(e)

            return None

    # ==================================================
    # Get ModelDoc2
    # ==================================================

    def get_model_from_component(
        self,
        component
    ):
        """
        Get the actual SolidWorks ModelDoc2
        from an assembly component.
        """

        try:

            sw_component = getattr(
                component,
                "sw_component",
                None
            )

            if sw_component is None:

                return None

            model = (
                sw_component.GetModelDoc2
            )

            if model is None:

                return None

            return model

        except Exception as e:

            print()
            print(
                "Could not get ModelDoc2:"
            )

            print(
                component.name
            )

            print(e)

            return None

    # ==================================================
    # Find Sheet Metal Feature
    # ==================================================

    def find_sheet_metal_feature(
        self,
        model
    ):
        """
        Search the FeatureManager tree for
        a Sheet Metal feature.
        """

        if model is None:
            return None

        try:

            feature = model.FirstFeature

        except Exception:

            return None

        while feature is not None:

            try:

                feature_type = (
                    feature.GetTypeName2
                )

                if feature_type:

                    if (
                        feature_type.lower()
                        == "sheetmetal"
                    ):

                        return feature

            except Exception:
                pass

            try:

                feature = (
                    feature.GetNextFeature
                )

            except Exception:

                break

        return None

    # ==================================================
    # Find Sheet Metal Feature - Debug
    # ==================================================

    def debug_features(
        self,
        model
    ):
        """
        Print the complete FeatureManager tree.
        """

        if model is None:

            print(
                "Model is None."
            )

            return

        print()
        print(
            "Feature Tree:"
        )

        try:

            feature = model.FirstFeature

        except Exception as e:

            print(e)
            return

        count = 0

        while feature is not None:

            count += 1

            try:
                name = feature.Name
            except Exception:
                name = "?"

            try:
                feature_type = (
                    feature.GetTypeName2
                )
            except Exception:
                feature_type = "?"

            print(
                f"  {count}. "
                f"{name} | "
                f"{feature_type}"
            )

            try:

                feature = (
                    feature.GetNextFeature
                )

            except Exception:

                break

        print(
            "Feature count:",
            count
        )

    # ==================================================
    # Is Sheet Metal
    # ==================================================

    def is_sheet_metal(
        self,
        component
    ):
        """
        Determine whether an assembly component
        is a Sheet Metal part.
        """

        if not component.is_part:

            return False

        # ----------------------------------------------
        # First try ModelDoc2 from assembly component
        # ----------------------------------------------

        model = (
            self.get_model_from_component(
                component
            )
        )

        # ----------------------------------------------
        # If unavailable, open the part
        # ----------------------------------------------

        if model is None:

            model = self.open_part(
                component.path
            )

        if model is None:

            return False

        # ----------------------------------------------
        # Search Sheet Metal feature
        # ----------------------------------------------

        feature = (
            self.find_sheet_metal_feature(
                model
            )
        )

        if feature is not None:

            return True

        return False

    # ==================================================
    # Get Thickness
    # ==================================================

    def get_thickness(
        self,
        component
    ):
        """
        Return Sheet Metal thickness in meters.

        Returns None if the part is not Sheet Metal.
        """

        if not component.is_part:

            return None

        # ----------------------------------------------
        # Get ModelDoc2
        # ----------------------------------------------

        model = (
            self.get_model_from_component(
                component
            )
        )

        if model is None:

            model = self.open_part(
                component.path
            )

        if model is None:

            return None

        # ----------------------------------------------
        # Find Sheet Metal feature
        # ----------------------------------------------

        feature = (
            self.find_sheet_metal_feature(
                model
            )
        )

        if feature is None:

            return None

        # ----------------------------------------------
        # Get feature definition
        # ----------------------------------------------

        try:

            feature_data = (
                feature.GetDefinition
            )

            if feature_data is None:

                return None

        except Exception:

            return None

        # ----------------------------------------------
        # Thickness
        # ----------------------------------------------

        try:

            thickness = (
                feature_data.Thickness
            )

            if thickness is None:

                return None

            thickness = float(
                thickness
            )

            if thickness <= 0:

                return None

            return thickness

        except Exception as e:

            print()
            print(
                "Could not read thickness:"
            )

            print(
                component.name
            )

            print(e)

            return None

    # ==================================================
    # Find Flat Pattern
    # ==================================================

    def find_flat_pattern_feature(
        self,
        model
    ):
        """
        Find Flat Pattern feature.
        """

        if model is None:

            return None

        try:

            feature = (
                model.FirstFeature
            )

        except Exception:

            return None

        while feature is not None:

            try:

                feature_type = (
                    feature.GetTypeName2
                )

                if feature_type:

                    if (
                        feature_type.lower()
                        == "flatpattern"
                    ):

                        return feature

            except Exception:
                pass

            try:

                feature = (
                    feature.GetNextFeature
                )

            except Exception:

                break

        return None

    # ==================================================
    # Activate Flat Pattern
    # ==================================================

    def activate_flat_pattern(
        self,
        model
    ):
        """
        Suppress/un-suppress Flat Pattern.
        """

        flat_pattern = (
            self.find_flat_pattern_feature(
                model
            )
        )

        if flat_pattern is None:

            print(
                "Flat Pattern feature "
                "not found."
            )

            return False

        try:

            result = (
                flat_pattern.SetSuppression2(
                    1,
                    2,
                    None
                )
            )

            if result:

                return True

            return False

        except Exception as e:

            print()
            print(
                "Could not activate "
                "Flat Pattern:"
            )

            print(e)

            return False

    # ==================================================
    # Export DWG
    # ==================================================

    def export_dxf(
        self,
        model,
        output_path
    ):
        """
        Export active Flat Pattern
        as DWG/DXF.
        """

        if model is None:

            return False

        try:

            model_path = (
                model.GetPathName
            )

            result = (
                model.ExportToDWG2(
                    str(output_path),
                    model_path,
                    1,
                    True,
                    None,
                    False,
                    False,
                    0,
                    None
                )
            )

            return bool(result)

        except Exception as e:

            print()
            print(
                "DWG/DXF export failed:"
            )

            print(e)

            return False

