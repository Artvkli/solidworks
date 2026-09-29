import csv
from pathlib import Path

try:
    import pythoncom
    from win32com.client import VARIANT
except ImportError:  # allows importing outside Windows
    pythoncom = None
    VARIANT = None

# =========================================================
# SOLIDWORKS CONSTANTS
# =========================================================

SW_DOC_PART = 1  # swDocPART
SW_SOLID_BODY = 0  # swSolidBody

SW_OPEN_SILENT = 1  # swOpenDocOptions_Silent
SW_OPEN_READONLY = 2  # swOpenDocOptions_ReadOnly

SW_TN_SHEET_METAL = "SheetMetal"  # swTnSheetMetal

MAX_FEATURES = 100000  # safety guard against endless loops
MAX_DEBUG_PARTS = 3  # how many failed parts print diagnostic lines


class SheetMetalDetector:
    def __init__(self, sw_app, assembly_path=None):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path) if assembly_path else None
        self._opened_models = {}
        self._opened_by_us = []  # titles of documents this class opened
        self.last_debug = []

    # =====================================================
    # SAFE COM HELPERS
    # =====================================================

    @staticmethod
    def _read(obj, name, default=None):
        """Read a scalar property (str / float / int); calls it if it is a method."""
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
        """Call a COM method safely. Returns default on error or None result."""
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
        # 1. Straight from the resolved SolidWorks component
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
            # 3. Same folder as the assembly
            candidate = self.assembly_path.parent / original_path.name
            if candidate.exists():
                model = self.open_part(candidate)
                if model is not None:
                    return model

            # 4. Recursive search below the assembly folder
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
    # OPEN / CLOSE PART
    # =====================================================

    def open_part(self, part_path):
        part_path = Path(part_path)

        if not part_path.exists():
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

        # Open silently and read-only (errors / warnings must be ByRef)
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
                title = self._read(model, "GetTitle")
                if title:
                    self._opened_by_us.append(title)
                return model

        except Exception as e:
            print(f"  Open part error: {part_path}\n  {e}")

        return None

    def close_opened_models(self):
        """Close only the documents that this class opened itself."""
        for title in self._opened_by_us:
            try:
                self.sw_app.CloseDoc(title)
            except Exception:
                pass
        self._opened_by_us = []
        self._opened_models = {}

    # =====================================================
    # SMALL HELPERS
    # =====================================================

    @staticmethod
    def _obj(obj, name):
        """Read a COM object property (never 'calls' a COM object by mistake)."""
        if obj is None:
            return None
        try:
            value = getattr(obj, name)
        except Exception:
            return None
        if value is None:
            return None
        if hasattr(value, "_oleobj_"):  # already a COM object
            return value
        if callable(value):
            try:
                return value()
            except Exception:
                return None
        return value

    @staticmethod
    def _positive(value):
        try:
            value = float(value)
        except Exception:
            return None
        return value if value > 0 else None

    # =====================================================
    # BODIES (secondary detection method)
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
        """Walk the top level of the FeatureManager tree (FirstFeature chain)."""
        current = self._obj(model, "FirstFeature")
        count = 0
        while current is not None and count < MAX_FEATURES:
            yield current
            current = self._call(current, "GetNextFeature")
            count += 1

    def get_features(self, model):
        """All top-level features. Uses FeatureManager.GetFeatures, then FirstFeature."""
        features = []

        feature_manager = self._obj(model, "FeatureManager")
        raw = self._call(feature_manager, "GetFeatures", True)
        if raw:
            try:
                features = [f for f in raw if f is not None]
            except Exception:
                features = []

        if not features:
            features = list(self.iter_features(model))

        return features

    def _type_name(self, feature):
        name = self._call(feature, "GetTypeName2")
        return str(name).lower() if name else ""

    def get_sheet_metal_feature(self, model, features=None):
        wanted = SW_TN_SHEET_METAL.lower()

        if features is None:
            features = self.get_features(model)

        # Pass 1: top level
        for feature in features:
            if self._type_name(feature) == wanted:
                return feature

        # Pass 2: default feature name
        feature = self._call(model, "FeatureByName", "Sheet-Metal1")
        if feature is not None and self._type_name(feature) == wanted:
            return feature

        # Pass 3: sub features
        for feature in features:
            sub = self._call(feature, "GetFirstSubFeature")
            count = 0
            while sub is not None and count < MAX_FEATURES:
                if self._type_name(sub) == wanted:
                    return sub
                sub = self._call(sub, "GetNextSubFeature")
                count += 1

        return None

    # =====================================================
    # THICKNESS (meters)
    # =====================================================

    def _active_config_name(self, model):
        config_manager = self._obj(model, "ConfigurationManager")
        config = self._obj(config_manager, "ActiveConfiguration")
        return self._read(config, "Name")

    def _dimension_value(self, model, dimension):
        """Dimension value in meters (SI) or None."""
        if dimension is None:
            return None

        value = self._positive(self._read(dimension, "SystemValue"))
        if value:
            return value

        config_name = self._active_config_name(model)
        if config_name:
            value = self._positive(
                self._call(dimension, "GetSystemValue2", config_name)
            )
            if value:
                return value

        return None

    def _thickness_from_definition(self, model, feature):
        definition = self._call(feature, "GetDefinition")
        if definition is None:
            return None

        value = self._positive(self._read(definition, "Thickness"))
        if value:
            return value

        # Some feature data objects need selection access first
        self._call(definition, "AccessSelections", model, None)
        try:
            return self._positive(self._read(definition, "Thickness"))
        finally:
            self._call(definition, "ReleaseSelectionAccess")

    def _thickness_from_parameter(self, model, feature_name):
        try:
            dimension = model.Parameter(f"Thickness@{feature_name}")
        except Exception:
            return None
        return self._dimension_value(model, dimension)

    def _thickness_from_display_dimensions(self, model, feature):
        display = self._call(feature, "GetFirstDisplayDimension")
        count = 0

        while display is not None and count < 500:
            dimension = self._call(display, "GetDimension2", 0)

            full_name = str(self._read(dimension, "FullName") or "").lower()
            short_name = str(self._read(dimension, "Name") or "").lower()

            if full_name.startswith("thickness@") or short_name == "thickness":
                value = self._dimension_value(model, dimension)
                if value:
                    return value

            display = self._call(feature, "GetNextDisplayDimension", display)
            count += 1

        return None

    def _find_thickness(self, model, feature, features, debug):
        # 1. SheetMetalFeatureData.Thickness
        if feature is not None:
            value = self._thickness_from_definition(model, feature)
            debug.append(f"definition.Thickness -> {value}")
            if value:
                return value

        # 2. Dimension "Thickness@<feature name>"
        names = []
        if feature is not None:
            feature_name = self._read(feature, "Name")
            if feature_name:
                names.append(str(feature_name))
        if "Sheet-Metal1" not in names:
            names.append("Sheet-Metal1")

        for name in names:
            value = self._thickness_from_parameter(model, name)
            debug.append(f"Parameter('Thickness@{name}') -> {value}")
            if value:
                return value

        # 3. Display dimensions of sheet metal related features
        related = ("sheetmetal", "smbaseflange", "baseflange")
        for candidate in features:
            if self._type_name(candidate) in related:
                value = self._thickness_from_display_dimensions(model, candidate)
                if value:
                    debug.append(f"display dimension of sheet metal feature -> {value}")
                    return value

        # 4. Last resort: any feature that owns a 'Thickness' dimension
        for candidate in features:
            value = self._thickness_from_display_dimensions(model, candidate)
            if value:
                debug.append(f"display dimension of any feature -> {value}")
                return value

        debug.append("display dimensions -> nothing found")
        return None

    # =====================================================
    # ANALYZE ONE MODEL -> (is_sheet_metal, thickness_m)
    # =====================================================

    def analyze_model(self, model):
        debug = []
        self.last_debug = debug

        doc_type = self._call(model, "GetType")
        if doc_type is not None and doc_type != SW_DOC_PART:
            debug.append(f"document type is {doc_type}, not a part")
            return False, None

        features = self.get_features(model)
        feature = self.get_sheet_metal_feature(model, features)

        sample = []
        for f in features[:12]:
            sample.append(f"{self._read(f, 'Name')}:{self._call(f, 'GetTypeName2')}")
        debug.append(f"top-level features: {len(features)}  {sample}")
        debug.append(
            "SheetMetal feature: "
            + (str(self._read(feature, "Name")) if feature is not None else "NOT FOUND")
        )

        is_sm = feature is not None
        if not is_sm:
            hits = len(self.get_sheet_metal_bodies(model))
            debug.append(f"sheet metal bodies: {hits}")
            is_sm = hits > 0

        thickness = None
        if is_sm:
            thickness = self._find_thickness(model, feature, features, debug)

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
        failed_parts = []
        debug_printed = 0

        print()
        print("=" * 60)
        print("SHEET METAL SCANNER")
        print("=" * 60)
        print(f"Unique parts to scan: {total}")
        print()

        try:
            for index, part in enumerate(unique_parts, start=1):
                label = f"[{index}/{total}] {part['name']} (x{part['quantity']})"

                try:
                    component = part["components"][0]
                    model = self.get_model_from_component(component)

                    if model is None:
                        print(f"{label} -> FAILED: model not found")
                        failed_parts.append(part["name"])
                        continue

                    is_sm, thickness = self.analyze_model(model)

                    if not is_sm:
                        print(f"{label} -> not sheet metal")
                        not_sheet_metal += 1
                        continue

                    thickness_mm = self.thickness_mm(thickness)

                    if thickness_mm is None:
                        print(f"{label} -> FAILED: thickness not found")
                        failed_parts.append(part["name"])
                        if debug_printed < MAX_DEBUG_PARTS:
                            debug_printed += 1
                            for line in self.last_debug:
                                print(f"      [debug] {line}")
                        continue

                    sheet_metal_parts.append(
                        {
                            "name": part["name"],
                            "path": part.get("path"),
                            "quantity": part["quantity"],
                            "thickness": thickness_mm,
                        }
                    )
                    print(f"{label} -> SHEET METAL, {thickness_mm:.3f} mm")

                except Exception as e:
                    failed_parts.append(part.get("name", "?"))
                    print(f"{label} -> ERROR: {e}")
        finally:
            self.close_opened_models()

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
        total_quantity = sum(g["quantity"] for g in thickness_groups.values())

        # ---------------- report ----------------

        print()
        print("=" * 60)
        print("SHEET METAL QUANTITY BY THICKNESS")
        print("=" * 60)
        print(f"Unique parts scanned:  {total}")
        print(f"Sheet metal (unique):  {len(sheet_metal_parts)}")
        print(f"Not sheet metal:       {not_sheet_metal}")
        print(f"Failed:                {len(failed_parts)}")
        print("-" * 60)

        for thickness, group in thickness_groups.items():
            print(
                f"{thickness:>8.3f} mm  |  "
                f"Unique parts: {group['unique_parts']:<3}  |  "
                f"Quantity: {group['quantity']}"
            )

        print("-" * 60)
        print(f"TOTAL SHEET METAL QUANTITY: {total_quantity}")

        if failed_parts:
            print()
            print("Failed parts (not counted):")
            for name in failed_parts:
                print(f"  - {name}")

        return {
            "unique_parts": total,
            "sheet_metal": len(sheet_metal_parts),
            "not_sheet_metal": not_sheet_metal,
            "failed": len(failed_parts),
            "failed_parts": failed_parts,
            "total_quantity": total_quantity,
            "sheet_metal_parts": sheet_metal_parts,
            "thickness_groups": thickness_groups,
            "csv_path": None,
        }

    # =====================================================
    # SCAN + EXPORT
    # =====================================================

    def scan_and_export(self, unique_parts, output_folder=None):
        result = self.scan_sheet_metal(unique_parts)

        if output_folder:
            stem = self.assembly_path.stem if self.assembly_path else "assembly"
            csv_path = Path(output_folder) / f"{stem}_sheet_metal.csv"
            if self.export_csv(result, csv_path):
                result["csv_path"] = csv_path

        return result

    def export_csv(self, result, csv_path):
        csv_path = Path(csv_path)

        try:
            csv_path.parent.mkdir(parents=True, exist_ok=True)

            with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)

                writer.writerow(["Thickness (mm)", "Unique Parts", "Total Quantity"])
                for thickness, group in result["thickness_groups"].items():
                    writer.writerow(
                        [thickness, group["unique_parts"], group["quantity"]]
                    )
                writer.writerow(
                    ["TOTAL", result["sheet_metal"], result["total_quantity"]]
                )

                writer.writerow([])
                writer.writerow(["Thickness (mm)", "Part", "Quantity"])
                rows = sorted(
                    result["sheet_metal_parts"],
                    key=lambda item: (item["thickness"], item["name"]),
                )
                for item in rows:
                    writer.writerow([item["thickness"], item["name"], item["quantity"]])

            print()
            print(f"CSV saved: {csv_path}")
            return True

        except Exception as e:
            print()
            print(f"Could not save CSV (is it open in Excel?): {e}")
            return False
