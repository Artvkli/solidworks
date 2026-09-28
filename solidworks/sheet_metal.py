from pathlib import Path

import pythoncom
import win32com.client

from solidworks.utils import com_value


class SheetMetalDetector:
    # Feature type names that mark a part as sheet metal
    SHEET_METAL_TYPES = {"sheetmetal", "smbaseflange", "flatpattern"}

    def __init__(self, sw_app):
        self.sw_app = sw_app
        self._opened_by_us = set()

    # ------------------------------------------------------------------
    # Open / close
    # ------------------------------------------------------------------

    def open_part(self, part_path):
        """Open a SolidWorks part (or reuse it if already open)."""

        part_path = str(part_path)

        # Already open (e.g. loaded with the assembly)?
        try:
            model = self.sw_app.GetOpenDocumentByName(part_path)
        except Exception:
            model = None

        if model is not None:
            return model

        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        model = self.sw_app.OpenDoc6(part_path, 1, 0, "", errors, warnings)

        if model is None:
            print(f"OpenDoc6 failed. Errors: {errors.value}")
            return None

        self._opened_by_us.add(part_path.lower())
        return model

    def close_part(self, model):
        """Close a part only if this class opened it."""

        if model is None:
            return

        path = com_value(model, "GetPathName") or ""

        if path.lower() not in self._opened_by_us:
            return

        try:
            self.sw_app.CloseDoc(path)
            self._opened_by_us.discard(path.lower())
        except Exception as e:
            print("Could not close part:", e)

    def activate_document(self, model):
        """Make the part the active document (needed for export)."""

        try:
            title = Path(com_value(model, "GetPathName")).name

            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

            self.sw_app.ActivateDoc3(title, False, 0, errors)
            return True

        except Exception as e:
            print("Could not activate document:", e)
            return False

    # ------------------------------------------------------------------
    # Feature search
    # ------------------------------------------------------------------

    def _iter_features(self, model):
        feature = com_value(model, "FirstFeature")

        while feature is not None:
            yield feature
            feature = com_value(feature, "GetNextFeature")

    def _find_feature(self, model, type_name):
        type_name = type_name.lower()

        for feature in self._iter_features(model):
            t = (com_value(feature, "GetTypeName2") or "").lower()
            if t == type_name:
                return feature

        return None

    def find_sheet_metal_feature(self, model):
        return self._find_feature(model, "SheetMetal")

    def find_flat_pattern_feature(self, model):
        return self._find_feature(model, "FlatPattern")

    def is_sheet_metal_model(self, model):
        for feature in self._iter_features(model):
            t = (com_value(feature, "GetTypeName2") or "").lower()
            if t in self.SHEET_METAL_TYPES:
                return True
        return False

    def print_features(self, model):
        """Debug helper: print every feature name and type."""

        for feature in self._iter_features(model):
            print(
                com_value(feature, "Name"),
                "|",
                com_value(feature, "GetTypeName2"),
            )

    # ------------------------------------------------------------------
    # Flat pattern
    # ------------------------------------------------------------------

    def activate_flat_pattern(self, model):
        """Unsuppress the Flat Pattern feature (current configuration)."""

        flat_pattern = self.find_flat_pattern_feature(model)

        if flat_pattern is None:
            print("Flat Pattern feature not found.")
            return False

        try:
            # 1 = swUnSuppressFeature, 1 = swThisConfiguration
            success = flat_pattern.SetSuppression2(1, 1, None)

            if success:
                model.EditRebuild3()
                print("Flat Pattern activated.")
                return True

            print("Could not activate Flat Pattern.")
            return False

        except Exception as e:
            print("Error while activating Flat Pattern:")
            print(e)
            return False

    def deactivate_flat_pattern(self, model):
        """Suppress the Flat Pattern again (leave the part as it was)."""

        flat_pattern = self.find_flat_pattern_feature(model)

        if flat_pattern is None:
            return

        try:
            # 0 = swSuppressFeature, 1 = swThisConfiguration
            flat_pattern.SetSuppression2(0, 1, None)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def export_dxf(self, model, output_path, sheet_metal_options=1):
        """
        Export the active Flat Pattern to DWG/DXF.

        sheet_metal_options (bit mask):
            1 = geometry, 2 = hidden edges, 4 = bend lines,
            8 = sketches, 16 = forming tools
        """

        try:
            model_path = com_value(model, "GetPathName")

            print()
            print("DWG Export")
            print("Model:", model_path)
            print("Output:", output_path)

            if not self.activate_document(model):
                return False

            # Identity alignment: origin + X axis + Y axis
            alignment = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8,
                [0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 0.0],
            )

            result = model.ExportToDWG2(
                str(output_path),
                model_path,
                1,  # swExportToDWG_ExportSheetMetal
                True,  # single file
                alignment,
                False,
                False,
                sheet_metal_options,
                None,
            )

            print("ExportToDWG2 result:", result)

            if result:
                print(f"DWG exported: {output_path}")
                return True

            print("DWG export returned False.")
            return False

        except Exception as e:
            print("Error while exporting DWG:")
            print(type(e).__name__)
            print(e)
            return False

    # ------------------------------------------------------------------
    # Component helpers
    # ------------------------------------------------------------------

    def is_sheet_metal(self, component):
        if not component.is_part or component.suppressed:
            return False

        model = None

        try:
            model = self.open_part(component.path)

            if model is None:
                print(f"Could not open part: {component.name}")
                return False

            return self.is_sheet_metal_model(model)

        except Exception as e:
            print(f"Could not inspect part: {component.name}")
            print(e)
            return False

        finally:
            self.close_part(model)

    def get_thickness(self, component):
        """Get Sheet Metal thickness in meters (None if not sheet metal)."""

        if not component.is_part or component.suppressed:
            return None

        model = None

        try:
            model = self.open_part(component.path)

            if model is None:
                print(f"Could not open part: {component.name}")
                return None

            # Try the Sheet-Metal feature first, then the Base-Flange
            for type_name in ("SheetMetal", "SMBaseFlange"):
                feature = self._find_feature(model, type_name)

                if feature is None:
                    continue

                data = com_value(feature, "GetDefinition")

                if data is None:
                    continue

                thickness = com_value(data, "Thickness")

                if thickness:
                    return float(thickness)

            return None

        except Exception as e:
            print(f"Could not read thickness: {component.name}")
            print(e)
            return None

        finally:
            self.close_part(model)
