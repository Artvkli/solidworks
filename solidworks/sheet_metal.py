
from pathlib import Path

import pythoncom
import win32com.client


SW_DOC_PART = 1
SW_SOLID_BODY = 0

SW_EXPORT_SHEET_METAL = 1

# Flat Pattern Geometry
# + Bend Lines
SW_SHEET_METAL_OPTIONS = 5


class SheetMetalDetector:

    def __init__(self, sw_app):

        self.sw_app = sw_app

        self._opened_models = {}

    # =====================================================
    # COM
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

    def _call(
        self,
        obj,
        name,
        *args
    ):

        if obj is None:
            return None

        try:

            method = getattr(
                obj,
                name
            )

            if callable(method):

                return method(*args)

            return method

        except Exception:

            return None

    # =====================================================
    # OPEN PART
    # =====================================================

    def open_part(self, part_path):

        part_path = Path(part_path)

        if not part_path.exists():

            print(
                "Part does not exist:",
                part_path
            )

            return None

        key = str(
            part_path
        ).lower()

        # Already opened by us
        if key in self._opened_models:

            return self._opened_models[key]

        # Already open in SolidWorks
        try:

            model = (
                self.sw_app
                .GetOpenDocumentByName(
                    str(part_path)
                )
            )

            if model is not None:

                self._opened_models[key] = model

                return model

        except Exception:
            pass

        # Open
        try:

            errors = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4,
                0
            )

            warnings = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4,
                0
            )

            model = (
                self.sw_app.OpenDoc6(
                    str(part_path),
                    SW_DOC_PART,
                    0,
                    "",
                    errors,
                    warnings,
                )
            )

            if model is None:

                print()
                print(
                    "Could not open:",
                    part_path
                )

                print(
                    "Errors:",
                    errors.value
                )

                print(
                    "Warnings:",
                    warnings.value
                )

                return None

            self._opened_models[key] = model

            return model

        except Exception as e:

            print(
                "Open part error:",
                e
            )

            return None

    # =====================================================
    # MODEL FROM COMPONENT
    # =====================================================

    def get_model_from_component(
        self,
        component
    ):

        if component is None:
            return None

        sw_component = (
            component.sw_component
        )

        if sw_component is None:
            return None

        # ---------------------------------------------
        # First: GetModelDoc2
        # ---------------------------------------------

        model = self._call(
            sw_component,
            "GetModelDoc2"
        )

        if model is not None:

            return model

        # ---------------------------------------------
        # Try resolve
        # ---------------------------------------------

        try:

            sw_component.SetSuppression2(
                2
            )

            model = self._call(
                sw_component,
                "GetModelDoc2"
            )

            if model is not None:

                return model

        except Exception:
            pass

        # ---------------------------------------------
        # Last resort: open file
        # ---------------------------------------------

        return self.open_part(
            component.path
        )

    # =====================================================
    # BODIES
    # =====================================================

    def get_bodies(self, model):

        if model is None:
            return []

        try:

            bodies = self._call(
                model,
                "GetBodies2",
                SW_SOLID_BODY,
                False
            )

            if bodies is None:
                return []

            return list(bodies)

        except Exception:

            return []

    # =====================================================
    # SHEET METAL BODY
    # =====================================================

    def get_sheet_metal_bodies(
        self,
        model
    ):

        result = []

        for body in self.get_bodies(model):

            try:

                if bool(
                    self._call(
                        body,
                        "IsSheetMetal"
                    )
                ):

                    result.append(
                        body
                    )

            except Exception:

                pass

        return result

    # =====================================================
    # FEATURE TREE
    # =====================================================

    def iter_features(self, model):

        if model is None:
            return

        feature = self._get(
            model,
            "FirstFeature"
        )

        while feature is not None:

            yield feature

            feature = self._get(
                feature,
                "GetNextFeature"
            )

    def find_feature(
        self,
        model,
        type_name
    ):

        target = (
            type_name.lower()
        )

        for feature in self.iter_features(
            model
        ):

            feature_type = self._get(
                feature,
                "GetTypeName2"
            )

            if (
                feature_type
                and
                str(feature_type).lower()
                == target
            ):

                return feature

        return None

    # =====================================================
    # SHEET METAL FEATURE
    # =====================================================

    def get_sheet_metal_feature(
        self,
        model
    ):

        return self.find_feature(
            model,
            "SheetMetal"
        )

    # =====================================================
    # DETECT
    # =====================================================

    def is_sheet_metal(
        self,
        component
    ):

        if (
            component is None
            or not component.is_part
        ):

            return False

        model = (
            self.get_model_from_component(
                component
            )
        )

        if model is None:
            return False

        # Primary detection
        bodies = (
            self.get_sheet_metal_bodies(
                model
            )
        )

        if bodies:

            return True

        # Secondary detection
        feature = (
            self.get_sheet_metal_feature(
                model
            )
        )

        return feature is not None

    # =====================================================
    # THICKNESS
    # =====================================================

    def get_thickness(
        self,
        component
    ):

        if (
            component is None
            or not component.is_part
        ):

            return None

        model = (
            self.get_model_from_component(
                component
            )
        )

        if model is None:
            return None

        feature = (
            self.get_sheet_metal_feature(
                model
            )
        )

        if feature is None:
            return None

        definition = self._call(
            feature,
            "GetDefinition"
        )

        if definition is None:
            return None

        try:

            thickness = self._get(
                definition,
                "Thickness"
            )

            if thickness is None:
                return None

            thickness = float(
                thickness
            )

            if thickness <= 0:
                return None

            return thickness

        except Exception:

            return None

    # =====================================================
    # FLAT PATTERN
    # =====================================================

    def get_flat_pattern(
        self,
        model
    ):

        return self.find_feature(
            model,
            "FlatPattern"
        )

    # =====================================================
    # ACTIVATE FLAT PATTERN
    # =====================================================

    def activate_flat_pattern(
        self,
        model
    ):

        flat_pattern = (
            self.get_flat_pattern(
                model
            )
        )

        if flat_pattern is None:

            print(
                "Flat Pattern not found."
            )

            return False

        try:

            # swFeatureSuppressionAction_e
            # swUnSuppressFeature = 2

            result = (
                flat_pattern.SetSuppression2(
                    2
                )
            )

            return bool(result)

        except Exception as e:

            print(
                "Flat Pattern activation error:",
                e
            )

            return False

    # =====================================================
    # EXPORT DWG
    # =====================================================

    def export_dwg(
        self,
        component,
        output_path
    ):

        model = (
            self.get_model_from_component(
                component
            )
        )

        if model is None:

            print(
                "No ModelDoc2:"
            )

            print(
                component.name
            )

            return False

        # ---------------------------------------------
        # Sheet Metal check
        # ---------------------------------------------

        bodies = (
            self.get_sheet_metal_bodies(
                model
            )
        )

        if not bodies:

            print(
                "Not Sheet Metal:",
                component.name
            )

            return False

        # ---------------------------------------------
        # Output
        # ---------------------------------------------

        output_path = Path(
            output_path
        )

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        # ---------------------------------------------
        # Source model path
        # ---------------------------------------------

        model_path = self._get(
            model,
            "GetPathName"
        )

        if not model_path:

            print(
                "Model has no file path."
            )

            return False

        # ---------------------------------------------
        # Export
        # ---------------------------------------------

        try:

            print()
            print(
                "Exporting DWG:"
            )

            print(
                component.name
            )

            print(
                output_path
            )

            result = (
                model.ExportToDWG2(
                    str(output_path),
                    str(model_path),
                    SW_EXPORT_SHEET_METAL,
                    True,
                    None,
                    False,
                    False,
                    SW_SHEET_METAL_OPTIONS,
                    None,
                )
            )

            if result:

                print(
                    "DWG exported successfully."
                )

                return True

            print(
                "ExportToDWG2 returned False."
            )

            return False

        except Exception as e:

            print()
            print(
                "DWG export error:"
            )

            print(e)

            return False

    # =====================================================
    # SCAN + EXPORT
    # =====================================================

    def scan_and_export(
        self,
        components,
        unique_parts,
        output_folder
    ):

        output_folder = Path(
            output_folder
        )

        output_folder.mkdir(
            parents=True,
            exist_ok=True
        )

        exported = []
        sheet_metal = []
        failed = []
        non_sheet_metal = []

        print()
        print("=" * 60)
        print("SHEET METAL SCAN")
        print("=" * 60)

        total = len(
            unique_parts
        )

        for index, item in enumerate(
            unique_parts,
            start=1
        ):

            path = Path(
                item["path"]
            )

            quantity = item[
                "quantity"
            ]

            instances = item[
                "components"
            ]

            print()
            print("-" * 60)
            print(
                f"[{index}/{total}]"
            )

            print(
                "Part:",
                path.name
            )

            print(
                "Quantity:",
                quantity
            )

            # -----------------------------------------
            # Use first instance
            # -----------------------------------------

            component = instances[0]

            # -----------------------------------------
            # Detect
            # -----------------------------------------

            try:

                if not self.is_sheet_metal(
                    component
                ):

                    print(
                        "Result: NOT SHEET METAL"
                    )

                    non_sheet_metal.append(
                        path
                    )

                    continue

            except Exception as e:

                print(
                    "Detection failed:",
                    e
                )

                failed.append(
                    {
                        "path": path,
                        "reason": str(e),
                    }
                )

                continue

            print(
                "Result: SHEET METAL"
            )

            sheet_metal.append(
                path
            )

            # -----------------------------------------
            # Thickness
            # -----------------------------------------

            thickness = (
                self.get_thickness(
                    component
                )
            )

            if thickness:

                print(
                    "Thickness:",
                    f"{thickness * 1000:.3f} mm"
                )

            # -----------------------------------------
            # Output
            # -----------------------------------------

            output_path = (
                output_folder
                / f"{path.stem}.dwg"
            )

            # Prevent overwrite
            counter = 1

            while output_path.exists():

                output_path = (
                    output_folder
                    / f"{path.stem}_{counter}.dwg"
                )

                counter += 1

            # -----------------------------------------
            # Export
            # -----------------------------------------

            try:

                success = self.export_dwg(
                    component,
                    output_path
                )

                if success:

                    exported.append(
                        {
                            "part": path,
                            "output": output_path,
                            "quantity": quantity,
                            "thickness": thickness,
                        }
                    )

                else:

                    failed.append(
                        {
                            "path": path,
                            "reason": "DWG export failed",
                        }
                    )

            except Exception as e:

                failed.append(
                    {
                        "path": path,
                        "reason": str(e),
                    }
                )

        # =================================================
        # REPORT
        # =================================================

        print()
        print("=" * 60)
        print("FINAL REPORT")
        print("=" * 60)

        print()
        print(
            "Unique parts:",
            len(unique_parts)
        )

        print(
            "Sheet Metal:",
            len(sheet_metal)
        )

        print(
            "Exported:",
            len(exported)
        )

        print(
            "Not Sheet Metal:",
            len(non_sheet_metal)
        )

        print(
            "Failed:",
            len(failed)
        )

        print()
        print(
            "Output:",
            output_folder
        )

        if exported:

            print()
            print("DWG FILES:")

            for item in exported:

                print(
                    " ",
                    item["output"]
                )

        return {
            "exported": exported,
            "sheet_metal": sheet_metal,
            "non_sheet_metal": non_sheet_metal,
            "failed": failed,
        }

