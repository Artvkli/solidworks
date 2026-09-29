from pathlib import Path

# =========================================================
# SOLIDWORKS CONSTANTS
# =========================================================

SW_DOC_PART = 1
SW_SOLID_BODY = 0

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
                str(part_path),
                SW_DOC_PART,
                1,
                "",
                errors,
                warnings,
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
    # SHEET METAL FEATURE
    # =====================================================

    def get_sheet_metal_feature(self, model):
        return self.find_feature(model, "SheetMetal")

    # =====================================================
    # IS SHEET METAL
    # =====================================================

    def is_sheet_metal(self, component):
        if not component.is_part:
            return False

        model = self.get_model_from_component(component)

        if model is None:
            return False

        # -------------------------------------------------
        # First method: Sheet Metal Bodies
        # -------------------------------------------------

        bodies = self.get_sheet_metal_bodies(model)

        if bodies:
            return True

        # -------------------------------------------------
        # Second method: SheetMetal Feature
        # -------------------------------------------------

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

            if thickness is None:
                return None

            try:
                return float(thickness)
            except Exception:
                return None

        except Exception:
            return None

    # =====================================================
    # NORMALIZE THICKNESS
    # =====================================================

    @staticmethod
    def thickness_mm(thickness):
        if thickness is None:
            return None

        try:
            return round(float(thickness) * 1000, 3)
        except Exception:
            return None

    # =====================================================
    # SCAN SHEET METAL
    # =====================================================

    def scan_sheet_metal(self, unique_parts):
        total = len(unique_parts)

        sheet_metal_parts = []
        not_sheet_metal = 0
        failed = 0

        print()
        print("=" * 60)
        print("SHEET METAL SCANNER")
        print("=" * 60)

        print()
        print(f"Unique Parts to scan: {total}")

        # =================================================
        # SCAN
        # =================================================

        for index, part in enumerate(unique_parts, start=1):
            print()
            print("-" * 60)
            print(f"[{index}/{total}]")
            print(f"Part: {part['name']}")
            print(f"Quantity: {part['quantity']}")

            component = part["components"][0]

            try:
                print("Path:")
                print(f"  {component.path}")

                # -----------------------------------------
                # MODEL
                # -----------------------------------------

                model = self.get_model_from_component(component)

                if model is None:
                    print("Result: FAILED - MODEL NOT FOUND")
                    failed += 1
                    continue

                # -----------------------------------------
                # SHEET METAL
                # -----------------------------------------

                if not self.is_sheet_metal(component):
                    print("Result: NOT SHEET METAL")
                    not_sheet_metal += 1
                    continue

                # -----------------------------------------
                # THICKNESS
                # -----------------------------------------

                thickness = self.get_thickness(component)
                thickness_mm = self.thickness_mm(thickness)

                if thickness_mm is None:
                    print("Result: SHEET METAL - THICKNESS NOT FOUND")
                    failed += 1
                    continue

                # -----------------------------------------
                # SAVE
                # -----------------------------------------

                item = {
                    "name": part["name"],
                    "path": part["path"],
                    "quantity": part["quantity"],
                    "thickness": thickness_mm,
                    "component": component,
                    "model": model,
                }

                sheet_metal_parts.append(item)

                print("Result: SHEET METAL")
                print(f"Thickness: {thickness_mm:.3f} mm")

            except Exception as e:
                failed += 1
                print(f"Component processing error: {e}")

        # =================================================
        # GROUP BY THICKNESS
        # =================================================

        thickness_groups = {}

        for item in sheet_metal_parts:
            thickness = item["thickness"]

            if thickness not in thickness_groups:
                thickness_groups[thickness] = {
                    "thickness": thickness,
                    "unique_parts": 0,
                    "quantity": 0,
                    "parts": [],
                }

            thickness_groups[thickness]["unique_parts"] += 1
            thickness_groups[thickness]["quantity"] += item["quantity"]
            thickness_groups[thickness]["parts"].append(item)

        # =================================================
        # SORT
        # =================================================

        thickness_groups = dict(sorted(thickness_groups.items(), key=lambda x: x[0]))

        # =================================================
        # FINAL REPORT
        # =================================================

        print()
        print("=" * 60)
        print("SHEET METAL THICKNESS SUMMARY")
        print("=" * 60)

        print()
        print(f"Unique Parts:          {total}")
        print(f"Sheet Metal Parts:     {len(sheet_metal_parts)}")
        print(f"Not Sheet Metal:       {not_sheet_metal}")
        print(f"Failed:                {failed}")

        print()
        print("-" * 60)
        print("THICKNESS GROUPS")
        print("-" * 60)

        total_sheet_quantity = 0

        for thickness, group in thickness_groups.items():
            quantity = group["quantity"]
            unique_count = group["unique_parts"]

            total_sheet_quantity += quantity

            print(
                f"{thickness:>8.3f} mm"
                f"  |  "
                f"Unique Parts: {unique_count:<3}"
                f"  |  "
                f"Quantity: {quantity}"
            )

        print("-" * 60)
        print(f"Total Sheet Metal Quantity: {total_sheet_quantity}")

        # =================================================
        # RETURN
        # =================================================

        return {
            "unique_parts": total,
            "sheet_metal": len(sheet_metal_parts),
            "not_sheet_metal": not_sheet_metal,
            "failed": failed,
            "sheet_metal_parts": sheet_metal_parts,
            "thickness_groups": thickness_groups,
        }
