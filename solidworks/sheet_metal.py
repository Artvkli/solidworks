import csv
from pathlib import Path

try:
    import pythoncom
    from win32com.client import VARIANT
except ImportError:  # allows importing the file outside Windows
    pythoncom = None
    VARIANT = None

# =========================================================
# SOLIDWORKS CONSTANTS
# =========================================================

SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2
SW_SOLID_BODY = 0

SW_OPEN_SILENT = 1
SW_OPEN_READONLY = 2

# swTnSheetMetal
SW_TN_SHEET_METAL = "SheetMetal"


class SheetMetalDetector:
    def __init__(self, sw_app, assembly_path=None):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path) if assembly_path else None
        self._opened_models = {}

    # =====================================================
    # SAFE HELPERS
    # =====================================================

    @staticmethod
    def _read(obj, name, default=None):
        """Read a scalar property (str / float / int). Calls it if it is a method."""
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

    @staticmethod
    def _call(obj, name, *args, default=None):
        """Call a COM method safely."""
        if obj is None:
            return default
        try:
            method = getattr(obj, name)
            if callable(method):
                result = method(*args)
                return default if result is None else result
        except Exception:
            pass
        return default

    # =====================================================
    # GET MODEL
    # =====================================================

    def get_model_from_component(self, component):
        # 1. Directly from the resolved component
        sw_component = getattr(component, "sw_component", None)
        if sw_component is not None:
            model = self._call(sw_component, "GetModelDoc2")
            if model is not None:
                return model

        path_value = getattr(component, "path", None)
        if not path_value:
            return None
        original_path = Path(path_value)

        # 2. Original path
        if original_path.exists():
            model = self.open_part(original_path)
            if model is not None:
                return model

        if self.assembly_path:
            # 3. Same folder as assembly
            candidate = self.assembly_path.parent / original_path.name
            if candidate.exists():
                model = self.open_part(candidate)
                if model is not None:
                    return model

            # 4. Recursive search
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
            print(f"Part does not exist: {part_path}")
            return None

        key = str(part_path).lower()
        if key in self._opened_models:
            return self._opened_models[key]

        # Already open in SolidWorks?
        try:
            model = self.sw_app.GetOpenDocumentByName(str(part_path))
            if model is not None:
                self._opened_models[key] = model
                return model
        except Exception:
            pass

        # Open (errors / warnings must be passed ByRef through COM)
        try:
            if VARIANT is not None:
                errors = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                warnings = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            else:
                errors = 0
                warnings = 0

            model = self.sw_app.OpenDoc6(
                str(part_path),
                SW_DOC_PART,
                SW_OPEN_SILENT | SW_OPEN_READONLY,
                "",
                errors,
                warnings,
            )
            if model is not None:
                self._opened_models[key] = model
                return model

        except Exception as e:
            print(f"Open part error: {part_path}\n  {e}")

        return None

    # =====================================================
    # BODIES
    # =====================================================

    def get_bodies(self, model):
        try:
            bodies = model.GetBodies2(SW_SOLID_BODY, False)
            return list(bodies) if bodies else []
        except Exception:
            return []

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
    # FEATURES
    # =====================================================

    def iter_features(self, model):
        """Walk the top level of the FeatureManager tree."""
        first = self._call(model, "FirstFeature")

        if first is None:
            # some wrappers expose it as a property
            try:
                first = model.FirstFeature
            except Exception:
                first = None

        current = first
        while current is not None:
            yield current
            current = self._call(current, "GetNextFeature")

    def find_feature(self, model, type_name):
        wanted = type_name.lower()
        for feature in self.iter_features(model):
            feature_type = self._call(feature, "GetTypeName2")
            if feature_type and str(feature_type).lower() == wanted:
                return feature
        return None

    def get_sheet_metal_feature(self, model):
        return self.find_feature(model, SW_TN_SHEET_METAL)

    # =====================================================
    # THICKNESS (meters)
    # =====================================================

    def _thickness_from_feature(self, model, feature):
        if feature is None:
            return None

        # Method 1: SheetMetalFeatureData.Thickness
        definition = self._call(feature, "GetDefinition")
        value = self._read(definition, "Thickness")
        try:
            value = float(value)
            if value > 0:
                return value
        except Exception:
            pass

        # Method 2: dimension "Thickness@<feature name>"
        name = self._read(feature, "Name")
        if name:
            try:
                dimension = model.Parameter(f"Thickness@{name}")
                value = float(self._read(dimension, "SystemValue"))
                if value > 0:
                    return value
            except Exception:
                pass

        return None

    # =====================================================
    # ANALYZE ONE MODEL -> (is_sheet_metal, thickness_m)
    # =====================================================

    def analyze_model(self, model):
        doc_type = self._call(model, "GetType")
        if doc_type is not None and doc_type != SW_DOC_PART:
            return False, None

        feature = self.get_sheet_metal_feature(model)

        is_sm = feature is not None
        if not is_sm:
            is_sm = bool(self.get_sheet_metal_bodies(model))

        thickness = self._thickness_from_feature(model, feature) if is_sm else None
        return is_sm, thickness

    # Backward-compatible wrappers
    def is_sheet_metal(self, component):
        if not getattr(component, "is_part", True):
            return False
        model = self.get_model_from_component(component)
        return model is not None and self.analyze_model(model)[0]

    def get_thickness(self, component):
        model = self.get_model_from_component(component)
        if model is None:
            return None
        return self.analyze_model(model)[1]

    @staticmethod
    def thickness_mm(thickness):
        if thickness is None:
            return None
        try:
            return round(float(thickness) * 1000, 3)
        except Exception:
            return None

    # =====================================================
    # SCAN
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
        print(f"Unique parts to scan: {total}")
        print()

        for index, part in enumerate(unique_parts, start=1):
            label = f"[{index}/{total}] {part['name']} (x{part['quantity']})"

            try:
                component = part["components"][0]
                model = self.get_model_from_component(component)

                if model is None:
                    print(f"{label} -> FAILED: model not found")
                    failed += 1
                    continue

                is_sm, thickness = self.analyze_model(model)

                if not is_sm:
                    print(f"{label} -> not sheet metal")
                    not_sheet_metal += 1
                    continue

                thickness_mm = self.thickness_mm(thickness)

                if thickness_mm is None:
                    print(f"{label} -> FAILED: sheet metal, thickness not found")
                    failed += 1
                    continue

                sheet_metal_parts.append(
                    {
                        "name": part["name"],
                        "path": part.get("path"),
                        "quantity": part["quantity"],
                        "thickness": thickness_mm,
                        "component": component,
                        "model": model,
                    }
                )
                print(f"{label} -> SHEET METAL, {thickness_mm:.3f} mm")

            except Exception as e:
                failed += 1
                print(f"{label} -> ERROR: {e}")

        # ---------------- group by thickness ----------------

        thickness_groups = {}
        for item in sheet_metal_parts:
            group = thickness_groups.setdefault(
                item["thickness"],
                {
                    "thickness": item["thickness"],
                    "unique_parts": 0,
                    "quantity": 0,
                    "parts": [],
                },
            )
            group["unique_parts"] += 1
            group["quantity"] += item["quantity"]
            group["parts"].append(item)

        thickness_groups = dict(sorted(thickness_groups.items()))

        # ---------------- report ----------------

        print()
        print("=" * 60)
        print("SHEET METAL THICKNESS SUMMARY")
        print("=" * 60)
        print(f"Unique parts:      {total}")
        print(f"Sheet metal parts: {len(sheet_metal_parts)}")
        print(f"Not sheet metal:   {not_sheet_metal}")
        print(f"Failed:            {failed}")
        print("-" * 60)

        total_quantity = 0
        for thickness, group in thickness_groups.items():
            total_quantity += group["quantity"]
            print(
                f"{thickness:>8.3f} mm  |  "
                f"Unique parts: {group['unique_parts']:<3}  |  "
                f"Quantity: {group['quantity']}"
            )

        print("-" * 60)
        print(f"Total sheet metal quantity: {total_quantity}")

        return {
            "unique_parts": total,
            "sheet_metal": len(sheet_metal_parts),
            "not_sheet_metal": not_sheet_metal,
            "failed": failed,
            "sheet_metal_parts": sheet_metal_parts,
            "thickness_groups": thickness_groups,
        }

    # =====================================================
    # SCAN + EXPORT (name expected by the main script)
    # =====================================================

    def scan_and_export(self, unique_parts, output_path=None, *args, **kwargs):
        result = self.scan_sheet_metal(unique_parts)

        if output_path:
            self.export_csv(result, output_path)

        return result

    def export_csv(self, result, output_path):
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        rows = sorted(
            result["sheet_metal_parts"],
            key=lambda item: (item["thickness"], item["name"]),
        )

        with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["Thickness (mm)", "Part", "Quantity"])
            for item in rows:
                writer.writerow([item["thickness"], item["name"], item["quantity"]])

            writer.writerow([])
            writer.writerow(["Thickness (mm)", "Unique Parts", "Total Quantity"])
            for thickness, group in result["thickness_groups"].items():
                writer.writerow([thickness, group["unique_parts"], group["quantity"]])

        print(f"Exported: {output_path}")
