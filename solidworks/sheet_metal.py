# sheet_metal.py

from pathlib import Path
import win32com.client


# =========================================================
# SOLIDWORKS CONSTANTS
# =========================================================

SW_DOC_PART = 1
SW_SOLID_BODY = 0

# ExportToDWG2 Action
SW_EXPORT_SHEET_METAL = 1

# Sheet Metal options:
# Bit 1 = Flat pattern geometry
# Bit 3 = Bend lines
# 1 + 4 = 5
SW_SHEET_METAL_OPTIONS = 5

# Suppression state
SW_SUPPRESS = 0
SW_RESOLVE = 2


class SheetMetalDetector:
    def __init__(self, sw_app, assembly_path=None):

        self.sw_app = sw_app

        self.assembly_path = Path(assembly_path) if assembly_path else None

        self._opened_models = {}

    # =====================================================
    # SAFE COM READ
    # =====================================================

    @staticmethod
    def _read(obj, name, default=None):

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

    # =====================================================
    # SAFE COM CALL
    # =====================================================

    @staticmethod
    def _call(obj, name, *args, default=None):

        if obj is None:
            return default

        try:
            method = getattr(obj, name)

        except Exception:
            return default

        try:
            if callable(method):
                return method(*args)

            return default

        except Exception:
            return default

    # =====================================================
    # GET MODEL FROM COMPONENT
    # =====================================================

    def get_model_from_component(self, component):

        sw_component = component.sw_component

        # -------------------------------------------------
        # 1. GetModelDoc2
        # -------------------------------------------------

        if sw_component is not None:
            try:
                model = self._call(sw_component, "GetModelDoc2")

                if model is not None:
                    return model

            except Exception:
                pass

        # -------------------------------------------------
        # 2. Original referenced path
        # -------------------------------------------------

        original_path = Path(component.path)

        if original_path.exists():
            model = self.open_part(original_path)

            if model is not None:
                return model

        # -------------------------------------------------
        # 3. Search beside assembly
        # -------------------------------------------------

        if self.assembly_path:
            assembly_folder = self.assembly_path.parent

            candidate = assembly_folder / original_path.name

            if candidate.exists():
                print(f"Found local part:\n  {candidate}")

                model = self.open_part(candidate)

                if model is not None:
                    return model

        # -------------------------------------------------
        # 4. Recursive search
        # -------------------------------------------------

        if self.assembly_path:
            assembly_folder = self.assembly_path.parent

            filename = original_path.name

            try:
                matches = list(assembly_folder.rglob(filename))

            except Exception:
                matches = []

            if matches:
                candidate = matches[0]

                print(f"Found part:\n  {candidate}")

                model = self.open_part(candidate)

                if model is not None:
                    return model

        # -------------------------------------------------
        # Nothing found
        # -------------------------------------------------

        print(f"Could not find model for:\n  {component.name}")

        return None

    # =====================================================
    # OPEN PART
    # =====================================================

    def open_part(self, part_path):

        part_path = Path(part_path)

        if not part_path.exists():
            print(f"Part does not exist:\n  {part_path}")

            return None

        key = str(part_path).lower()

        # Already opened by scanner
        if key in self._opened_models:
            return self._opened_models[key]

        # -------------------------------------------------
        # Check SolidWorks opened documents
        # -------------------------------------------------

        try:
            model = self.sw_app.GetOpenDocumentByName(str(part_path))

            if model is not None:
                self._opened_models[key] = model

                return model

        except Exception:
            pass

        # -------------------------------------------------
        # Open part
        # -------------------------------------------------

        try:
            errors = 0
            warnings = 0

            model = self.sw_app.OpenDoc6(
                str(part_path), SW_DOC_PART, 1, "", errors, warnings
            )

            if model is not None:
                self._opened_models[key] = model

                return model

        except Exception as e:
            print(f"Open part error:\n  {part_path}\n  {e}")

        return None

    # =====================================================
    # ACTIVATE MODEL
    # =====================================================

    def activate_model(self, model):

        try:
            title = self._read(model, "GetTitle", "")

            if not title:
                return False

            errors = 0

            result = self.sw_app.ActivateDoc3(title, True, 2)

            return result is not None

        except Exception as e:
            print(f"Warning: ActivateDoc3 failed: {e}")

            return False

    # =====================================================
    # REBUILD MODEL
    # =====================================================

    def rebuild_model(self, model):

        try:
            result = self._call(model, "ForceRebuild3", False)

            return result

        except Exception:
            return None

    # =====================================================
    # GET BODIES
    # =====================================================

    def get_bodies(self, model):

        try:
            bodies = model.GetBodies2(SW_SOLID_BODY, False)

            if bodies is None:
                return []

            return list(bodies)

        except Exception as e:
            print(f"GetBodies2 error: {e}")

            return []

    # =====================================================
    # GET SHEET METAL BODIES
    # =====================================================

    def get_sheet_metal_bodies(self, model):

        result = []

        bodies = self.get_bodies(model)

        for body in bodies:
            try:
                if body.IsSheetMetal():
                    result.append(body)

            except Exception:
                pass

        return result

    # =====================================================
    # FEATURE ITERATOR
    # =====================================================

    def iter_features(self, model):

        try:
            feature = model.FirstFeature

            while feature is not None:
                yield feature

                try:
                    feature = feature.GetNextFeature()

                except Exception:
                    break

        except Exception:
            return

    # =====================================================
    # FIND FEATURE
    # =====================================================

    def find_feature(self, model, type_name):

        wanted = type_name.lower()

        for feature in self.iter_features(model):
            try:
                feature_type = feature.GetTypeName2()

                if feature_type and feature_type.lower() == wanted:
                    return feature

            except Exception:
                pass

        return None

    # =====================================================
    # GET SHEET METAL FEATURE
    # =====================================================

    def get_sheet_metal_feature(self, model):

        return self.find_feature(model, "SheetMetal")

    # =====================================================
    # GET FLAT PATTERN
    # =====================================================

    def get_flat_pattern(self, model):

        return self.find_feature(model, "FlatPattern")

    # =====================================================
    # ACTIVATE / UNSUPPRESS FLAT PATTERN
    # =====================================================

    def activate_flat_pattern(self, model):

        flat_pattern = self.get_flat_pattern(model)

        if flat_pattern is None:
            return False

        # -------------------------------------------------
        # Try unsuppress
        # -------------------------------------------------

        try:
            result = flat_pattern.SetSuppression2(SW_RESOLVE)

            return bool(result)

        except Exception:
            return True

    # =====================================================
    # CHECK SHEET METAL
    # =====================================================

    def is_sheet_metal(self, component):

        if not component.is_part:
            return False

        model = self.get_model_from_component(component)

        if model is None:
            return False

        # -------------------------------------------------
        # Method 1:
        # Body.IsSheetMetal()
        # -------------------------------------------------

        sheet_bodies = self.get_sheet_metal_bodies(model)

        if sheet_bodies:
            return True

        # -------------------------------------------------
        # Method 2:
        # SheetMetal feature
        # -------------------------------------------------

        feature = self.get_sheet_metal_feature(model)

        if feature is not None:
            return True

        return False

    # =====================================================
    # GET THICKNESS
    # =====================================================

    def get_thickness(self, component):

        model = self.get_model_from_component(component)

        if model is None:
            return None

        feature = self.get_sheet_metal_feature(model)

        if feature is None:
            return None

        try:
            definition = feature.GetDefinition()

            if definition is None:
                return None

            thickness = getattr(definition, "Thickness", None)

            if callable(thickness):
                try:
                    thickness = thickness()
                except Exception:
                    pass

            return thickness

        except Exception:
            return None

    # =====================================================
    # EXPORT DWG
    # =====================================================

    def export_dwg(self, component, output_path):

        # -------------------------------------------------
        # Get model
        # -------------------------------------------------

        model = self.get_model_from_component(component)

        if model is None:
            print("Could not obtain ModelDoc2.")

            return False

        # -------------------------------------------------
        # Activate document
        # -------------------------------------------------

        self.activate_model(model)

        # -------------------------------------------------
        # Rebuild
        # -------------------------------------------------

        self.rebuild_model(model)

        # -------------------------------------------------
        # Verify Sheet Metal body
        # -------------------------------------------------

        sheet_bodies = self.get_sheet_metal_bodies(model)

        if not sheet_bodies:
            print("No Sheet Metal body found.")

            return False

        print(f"Sheet Metal bodies: {len(sheet_bodies)}")

        # -------------------------------------------------
        # Get model path
        # -------------------------------------------------

        model_path = self._read(model, "GetPathName", "")

        if not model_path:
            model_path = component.path

        model_path = str(model_path)

        # -------------------------------------------------
        # Make sure output folder exists
        # -------------------------------------------------

        output_path = Path(output_path)

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # -------------------------------------------------
        # Flat Pattern
        # -------------------------------------------------

        flat_pattern = self.get_flat_pattern(model)

        if flat_pattern is not None:
            print("FlatPattern feature found.")

            try:
                self.activate_flat_pattern(model)

            except Exception as e:
                print(f"Warning activating FlatPattern: {e}")

        else:
            print("Warning: FlatPattern feature not found.")

        # -------------------------------------------------
        # Alignment
        #
        # 12 values:
        #
        # Origin
        # X axis
        # Y axis
        # Normal
        # -------------------------------------------------

        alignment = (0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0)

        # -------------------------------------------------
        # Sheet Metal Options
        #
        # Flat Pattern = 1
        # Bend Lines   = 4
        #
        # Total = 5
        # -------------------------------------------------

        sheet_metal_options = SW_SHEET_METAL_OPTIONS

        print()
        print("DWG EXPORT")

        print("-" * 40)

        print(f"Model:")

        print(f"  {model_path}")

        print(f"Output:")

        print(f"  {output_path}")

        print(f"Options:")

        print(f"  {sheet_metal_options}")

        # -------------------------------------------------
        # Export
        # -------------------------------------------------

        try:
            result = model.ExportToDWG2(
                str(output_path),
                model_path,
                SW_EXPORT_SHEET_METAL,
                True,
                alignment,
                False,
                False,
                sheet_metal_options,
                None,
            )

            print(f"ExportToDWG2 result: {result}")

            # -------------------------------------------------
            # Check output
            # -------------------------------------------------

            if result and output_path.exists():
                try:
                    file_size = output_path.stat().st_size

                    print("DWG created successfully.")

                    print(f"File size: {file_size:,} bytes")

                except Exception:
                    print("DWG created successfully.")

                return True

            # -------------------------------------------------
            # Some SolidWorks versions can create the file
            # while returning False.
            # -------------------------------------------------

            if output_path.exists():
                try:
                    file_size = output_path.stat().st_size

                    if file_size > 0:
                        print("DWG file exists despite False result.")

                        print(f"File size: {file_size:,} bytes")

                        return True

                except Exception:
                    pass

            print("SolidWorks returned False and no valid DWG was created.")

            return False

        except Exception as e:
            print()
            print("ExportToDWG2 exception:")

            print(f"  {e}")

            return False

    # =====================================================
    # UNIQUE OUTPUT PATH
    # =====================================================

    @staticmethod
    def unique_output_path(output_folder, part_name):

        output_folder = Path(output_folder)

        output_folder.mkdir(parents=True, exist_ok=True)

        base = output_folder / f"{part_name}.dwg"

        if not base.exists():
            return base

        counter = 1

        while True:
            candidate = output_folder / f"{part_name}_{counter}.dwg"

            if not candidate.exists():
                return candidate

            counter += 1

    # =====================================================
    # SCAN + EXPORT
    # =====================================================

    def scan_and_export(self, components, unique_parts, output_folder):

        output_folder = Path(output_folder)

        output_folder.mkdir(parents=True, exist_ok=True)

        exported = 0
        sheet_metal = 0
        non_sheet_metal = 0
        failed = 0

        total = len(unique_parts)

        print()
        print("=" * 60)
        print("SHEET METAL SCANNER")
        print("=" * 60)

        for index, part in enumerate(unique_parts, start=1):
            print()
            print("-" * 60)

            print(f"[{index}/{total}]")

            print(f"Part: {part['name']}")

            print(f"Quantity: {part['quantity']}")

            # -------------------------------------------------
            # Representative component
            # -------------------------------------------------

            component = part["components"][0]

            try:
                print("Referenced path:")

                print(f"  {component.path}")

                # -------------------------------------------------
                # Get model
                # -------------------------------------------------

                model = self.get_model_from_component(component)

                if model is None:
                    print("Result: FAILED - MODEL NOT FOUND")

                    failed += 1

                    continue

                # -------------------------------------------------
                # Check Sheet Metal
                # -------------------------------------------------

                is_sheet = self.is_sheet_metal(component)

                if not is_sheet:
                    print("Result: NOT SHEET METAL")

                    non_sheet_metal += 1

                    continue

                # -------------------------------------------------
                # Sheet Metal
                # -------------------------------------------------

                sheet_metal += 1

                print("Result: SHEET METAL")

                # -------------------------------------------------
                # Thickness
                # -------------------------------------------------

                thickness = self.get_thickness(component)

                if thickness is not None:
                    try:
                        thickness_mm = float(thickness) * 1000.0

                        print(f"Thickness: {thickness_mm:.3f} mm")

                    except Exception:
                        print(f"Thickness: {thickness}")

                # -------------------------------------------------
                # Output filename
                # -------------------------------------------------

                part_name = Path(component.path).stem

                output_path = self.unique_output_path(output_folder, part_name)

                print("Exporting:")

                print(f"  {output_path}")

                # -------------------------------------------------
                # Export
                # -------------------------------------------------

                success = self.export_dwg(component, output_path)

                if success:
                    exported += 1

                    print("DWG exported successfully.")

                else:
                    failed += 1

                    print("DWG export FAILED.")

            except Exception as e:
                failed += 1

                print(f"Component processing error: {e}")

        # =====================================================
        # FINAL REPORT
        # =====================================================

        print()
        print("=" * 60)
        print("FINAL REPORT")
        print("=" * 60)

        print(f"Unique parts: {total}")

        print(f"Sheet Metal: {sheet_metal}")

        print(f"Exported: {exported}")

        print(f"Not Sheet Metal: {non_sheet_metal}")

        print(f"Failed: {failed}")

        print()
        print(f"Output: {output_folder}")

        return {
            "unique_parts": total,
            "sheet_metal": sheet_metal,
            "exported": exported,
            "not_sheet_metal": non_sheet_metal,
            "failed": failed,
        }
