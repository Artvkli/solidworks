# sheet_metal.py

from pathlib import Path


# =========================================================
# SOLIDWORKS CONSTANTS
# =========================================================

SW_DOC_PART = 1
SW_SOLID_BODY = 0

# ExportToDWG2
SW_EXPORT_SHEET_METAL = 1

# Sheet metal export:
# 1 = flat pattern geometry
# 4 = bend lines
# 5 = 1 + 4
SW_SHEET_METAL_OPTIONS = 5

# Suppression
SW_RESOLVE = 2


class SheetMetalDetector:
    def __init__(self, sw_app, assembly_path=None):

        self.sw_app = sw_app

        self.assembly_path = Path(assembly_path) if assembly_path else None

        self._opened_models = {}

    # =====================================================
    # SAFE READ
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
    # SAFE CALL
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

        except Exception:
            pass

        return default

    # =====================================================
    # GET MODEL
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
        # 2. Original path
        # -------------------------------------------------

        original_path = Path(component.path)

        if original_path.exists():
            model = self.open_part(original_path)

            if model is not None:
                return model

        # -------------------------------------------------
        # 3. Same folder as assembly
        # -------------------------------------------------

        if self.assembly_path:
            candidate = self.assembly_path.parent / original_path.name

            if candidate.exists():
                print(f"Found local part:\n  {candidate}")

                model = self.open_part(candidate)

                if model is not None:
                    return model

        # -------------------------------------------------
        # 4. Recursive search
        # -------------------------------------------------

        if self.assembly_path:
            try:
                matches = list(self.assembly_path.parent.rglob(original_path.name))

            except Exception:
                matches = []

            if matches:
                model = self.open_part(matches[0])

                if model is not None:
                    return model

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

        if key in self._opened_models:
            return self._opened_models[key]

        # -------------------------------------------------
        # Already open
        # -------------------------------------------------

        try:
            model = self.sw_app.GetOpenDocumentByName(str(part_path))

            if model is not None:
                self._opened_models[key] = model

                return model

        except Exception:
            pass

        # -------------------------------------------------
        # Open
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
    # REBUILD
    # =====================================================

    def rebuild_model(self, model):

        try:
            return self._call(model, "ForceRebuild3", False)

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
    # SHEET METAL BODIES
    # =====================================================

    def get_sheet_metal_bodies(self, model):

        result = []

        for body in self.get_bodies(model):
            try:
                if body.IsSheetMetal():
                    result.append(body)

            except Exception:
                pass

        return result

    # =====================================================
    # RECURSIVE FEATURE WALK
    # =====================================================

    def walk_features(self, feature, level=0):
        """
        Recursively walk FeatureManager.

        FlatPattern can be nested under other features,
        so FirstFeature/GetNextFeature alone is not enough.
        """

        current = feature

        while current is not None:
            yield current

            # ---------------------------------------------
            # Sub features
            # ---------------------------------------------

            try:
                sub_feature = current.GetFirstSubFeature()

            except Exception:
                sub_feature = None

            if sub_feature is not None:
                yield from self.walk_features(sub_feature, level + 1)

            # ---------------------------------------------
            # Next feature
            # ---------------------------------------------

            try:
                current = current.GetNextFeature()

            except Exception:
                current = None

    # =====================================================
    # ALL FEATURES
    # =====================================================

    def iter_features(self, model):

        try:
            first = model.FirstFeature

        except Exception:
            first = None

        if first is not None:
            yield from self.walk_features(first)

    # =====================================================
    # FIND FEATURE
    # =====================================================

    def find_feature(self, model, type_name):

        wanted = type_name.lower()

        for feature in self.iter_features(model):
            try:
                feature_type = feature.GetTypeName2()

                if not feature_type:
                    continue

                if feature_type.lower() == wanted:
                    return feature

            except Exception:
                pass

        return None

    # =====================================================
    # PRINT FEATURE TREE
    # =====================================================

    def debug_feature_tree(self, model):

        print()
        print("FEATURE TREE")
        print("-" * 50)

        for feature in self.iter_features(model):
            try:
                name = self._read(feature, "Name", "")

                type_name = feature.GetTypeName2()

                print(f"{name} [{type_name}]")

            except Exception:
                pass

    # =====================================================
    # SHEET METAL FEATURE
    # =====================================================

    def get_sheet_metal_feature(self, model):

        return self.find_feature(model, "SheetMetal")

    # =====================================================
    # FLAT PATTERN
    # =====================================================

    def get_flat_pattern(self, model):

        # Normal feature
        feature = self.find_feature(model, "FlatPattern")

        if feature is not None:
            return feature

        # Alternative type name
        feature = self.find_feature(model, "FlatPatternFolder")

        if feature is not None:
            return feature

        return None

    # =====================================================
    # RESOLVE FLAT PATTERN
    # =====================================================

    def resolve_flat_pattern(self, model):

        flat_pattern = self.get_flat_pattern(model)

        if flat_pattern is None:
            return None

        try:
            result = flat_pattern.SetSuppression2(SW_RESOLVE)

            print(f"FlatPattern resolve: {result}")

        except Exception as e:
            print(f"FlatPattern resolve warning: {e}")

        return flat_pattern

    # =====================================================
    # IS SHEET METAL
    # =====================================================

    def is_sheet_metal(self, component):

        if not component.is_part:
            return False

        model = self.get_model_from_component(component)

        if model is None:
            return False

        bodies = self.get_sheet_metal_bodies(model)

        if bodies:
            return True

        feature = self.get_sheet_metal_feature(model)

        return feature is not None

    # =====================================================
    # THICKNESS
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
    # EXPORT
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
        # Rebuild
        # -------------------------------------------------

        self.rebuild_model(model)

        # -------------------------------------------------
        # Check bodies
        # -------------------------------------------------

        sheet_bodies = self.get_sheet_metal_bodies(model)

        print(f"Sheet Metal bodies: {len(sheet_bodies)}")

        if not sheet_bodies:
            return False

        # -------------------------------------------------
        # Get real model path
        # -------------------------------------------------

        model_path = self._read(model, "GetPathName", "")

        if not model_path:
            model_path = component.path

        model_path = str(model_path)

        # -------------------------------------------------
        # Flat Pattern
        # -------------------------------------------------

        flat_pattern = self.resolve_flat_pattern(model)

        if flat_pattern is not None:
            print("FlatPattern found.")

        else:
            print("FlatPattern feature not found.")

        # -------------------------------------------------
        # Output
        # -------------------------------------------------

        output_path = Path(output_path)

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # -------------------------------------------------
        # Alignment
        # -------------------------------------------------

        alignment = (0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0)

        # -------------------------------------------------
        # Export
        # -------------------------------------------------

        print()
        print("DWG EXPORT")

        print("-" * 40)

        print(f"Model:\n  {model_path}")

        print(f"Output:\n  {output_path}")

        print(f"SheetMetalOptions: {SW_SHEET_METAL_OPTIONS}")

        try:
            result = model.ExportToDWG2(
                str(output_path),
                model_path,
                SW_EXPORT_SHEET_METAL,
                True,
                alignment,
                False,
                False,
                SW_SHEET_METAL_OPTIONS,
                None,
            )

            print(f"ExportToDWG2 result: {result}")

            # -------------------------------------------------
            # Verify file
            # -------------------------------------------------

            if output_path.exists():
                try:
                    size = output_path.stat().st_size

                    if size > 0:
                        print("DWG created successfully.")

                        print(f"File size: {size:,} bytes")

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
    # UNIQUE OUTPUT
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

            component = part["components"][0]

            try:
                print("Referenced path:")

                print(f"  {component.path}")

                # -----------------------------------------
                # Model
                # -----------------------------------------

                model = self.get_model_from_component(component)

                if model is None:
                    print("Result: FAILED - MODEL NOT FOUND")

                    failed += 1
                    continue

                # -----------------------------------------
                # Sheet Metal
                # -----------------------------------------

                if not self.is_sheet_metal(component):
                    print("Result: NOT SHEET METAL")

                    non_sheet_metal += 1
                    continue

                sheet_metal += 1

                print("Result: SHEET METAL")

                # -----------------------------------------
                # Thickness
                # -----------------------------------------

                thickness = self.get_thickness(component)

                if thickness is not None:
                    try:
                        print(f"Thickness: {float(thickness) * 1000:.3f} mm")

                    except Exception:
                        pass

                # -----------------------------------------
                # Output name
                # -----------------------------------------

                part_name = Path(component.path).stem

                output_path = self.unique_output_path(output_folder, part_name)

                print("Exporting:")

                print(f"  {output_path}")

                # -----------------------------------------
                # Export
                # -----------------------------------------

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

        # =================================================
        # FINAL REPORT
        # =================================================

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
