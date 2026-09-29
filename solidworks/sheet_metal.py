import csv
import gc
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

SW_DOC_PART = 1
SW_SOLID_BODY = 0
SW_OPEN_SILENT = 1
SW_OPEN_READONLY = 2
SW_TN_SHEET_METAL = "sheetmetal"
MAX_FEATURES = 100000
STANDARD_THICKNESSES_MM = (0.5, 0.6, 0.7, 0.8, 1.0, 1.2, 1.25, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 6.0, 8.0, 10.0)
SNAP_TOLERANCE = 0.08


class SheetMetalDetector:
    """Detect sheet-metal parts and thickness without retaining Part COM objects."""

    def __init__(self, sw_app, assembly_path=None, read_thickness=True):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path).resolve() if assembly_path else None
        self.read_thickness = read_thickness
        self.last_method = ""
        self.last_debug = []
        self.last_thickness_source = ""
        self.last_thickness_raw_mm = None

    @staticmethod
    def _get(obj, name, default=None):
        if obj is None:
            return default
        try:
            value = getattr(obj, name)
            return value() if callable(value) else value
        except Exception:
            return default

    @staticmethod
    def _call(obj, name, *args, default=None):
        if obj is None:
            return default
        try:
            attr = getattr(obj, name)
            if args:
                return attr(*args) if callable(attr) else default
            if hasattr(attr, "_oleobj_"):
                return attr
            return attr() if callable(attr) else attr
        except Exception:
            return default

    @staticmethod
    def _obj(obj, name):
        return SheetMetalDetector._call(obj, name)

    @staticmethod
    def _positive(value):
        try:
            value = float(value)
            return value if value > 0 else None
        except Exception:
            return None

    @staticmethod
    def _byref_i4():
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    def open_part(self, path, read_only=True):
        path = Path(path).resolve()
        if not path.exists():
            return None, False, "file not found"

        try:
            existing = self.sw_app.GetOpenDocumentByName(str(path))
            if existing is not None:
                return existing, False, "already open"
        except Exception:
            pass

        options = SW_OPEN_SILENT | (SW_OPEN_READONLY if read_only else 0)
        errors = self._byref_i4()
        warnings = self._byref_i4()
        try:
            model = self.sw_app.OpenDoc6(str(path), SW_DOC_PART, options, "", errors, warnings)
            if model is None:
                return None, False, f"OpenDoc6 returned None (errors={getattr(errors, 'value', errors)}, warnings={getattr(warnings, 'value', warnings)})"
            return model, True, f"opened (errors={getattr(errors, 'value', errors)}, warnings={getattr(warnings, 'value', warnings)})"
        except Exception as exc:
            return None, False, f"OpenDoc6 exception: {exc}"

    def close_if_opened(self, model, opened_by_us):
        if not opened_by_us or model is None:
            return
        title = self._get(model, "GetTitle")
        try:
            if title:
                self.sw_app.CloseDoc(str(title))
        except Exception:
            pass
        del model
        gc.collect()
        try:
            pythoncom.CoFreeUnusedLibraries()
        except Exception:
            pass

    def activate_configuration(self, model, configuration):
        if not configuration:
            return True
        try:
            result = model.ShowConfiguration2(str(configuration))
            return bool(result) or str(self._get(model, "GetActiveConfiguration", "")) == str(configuration)
        except Exception:
            return False

    def rebuild(self, model):
        try:
            result = model.ForceRebuild3(True)
            return bool(result), "ForceRebuild3"
        except Exception as exc:
            try:
                result = model.EditRebuild3()
                return bool(result), "EditRebuild3"
            except Exception as exc2:
                return False, f"rebuild failed: {exc}; fallback: {exc2}"

    def get_bodies(self, model):
        try:
            raw = model.GetBodies2(SW_SOLID_BODY, False)
            return list(raw or [])
        except Exception:
            return []

    def get_sheet_metal_bodies(self, model):
        found = []
        for body in self.get_bodies(model):
            try:
                if body.IsSheetMetal():
                    found.append(body)
            except Exception:
                continue
        return found

    def iter_features(self, model):
        current = self._obj(model, "FirstFeature")
        count = 0
        while current is not None and count < MAX_FEATURES:
            yield current
            current = self._call(current, "GetNextFeature")
            count += 1

    def get_features(self, model):
        manager = self._obj(model, "FeatureManager")
        raw = self._call(manager, "GetFeatures", True)
        try:
            features = [f for f in (raw or []) if f is not None]
        except Exception:
            features = []
        return features or list(self.iter_features(model))

    def _type_name(self, feature):
        value = self._call(feature, "GetTypeName2")
        return str(value).lower() if value else ""

    def get_sheet_metal_feature(self, model, features=None):
        features = self.get_features(model) if features is None else features
        for feature in features:
            if self._type_name(feature) == SW_TN_SHEET_METAL:
                return feature
        for name in ("Sheet-Metal1", "Sheet-Metal"):
            feature = self._call(model, "FeatureByName", name)
            if feature is not None and self._type_name(feature) == SW_TN_SHEET_METAL:
                return feature
        for feature in features:
            sub = self._call(feature, "GetFirstSubFeature")
            count = 0
            while sub is not None and count < MAX_FEATURES:
                if self._type_name(sub) == SW_TN_SHEET_METAL:
                    return sub
                sub = self._call(sub, "GetNextSubFeature")
                count += 1
        return None

    def _active_config_name(self, model):
        manager = self._obj(model, "ConfigurationManager")
        config = self._obj(manager, "ActiveConfiguration")
        return self._get(config, "Name")

    def _dimension_value(self, model, dimension):
        value = self._positive(self._get(dimension, "SystemValue"))
        if value:
            return value
        config = self._active_config_name(model)
        if config:
            return self._positive(self._call(dimension, "GetSystemValue2", config))
        return None

    def _thickness_from_definition(self, model, feature):
        definition = self._call(feature, "GetDefinition")
        if definition is None:
            return None
        try:
            self._call(definition, "AccessSelections", model, None)
            return self._positive(self._get(definition, "Thickness"))
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
        for _ in range(500):
            if display is None:
                break
            dimension = self._call(display, "GetDimension2", 0)
            full_name = str(self._get(dimension, "FullName", "")).lower()
            short_name = str(self._get(dimension, "Name", "")).lower()
            if full_name.startswith("thickness@") or short_name == "thickness":
                value = self._dimension_value(model, dimension)
                if value:
                    return value
            display = self._call(feature, "GetNextDisplayDimension", display)
        return None

    def snap_thickness_mm(self, raw_mm):
        nearest = min(STANDARD_THICKNESSES_MM, key=lambda x: abs(x - raw_mm))
        if abs(nearest - raw_mm) / nearest <= SNAP_TOLERANCE:
            return nearest
        return round(raw_mm, 2)

    def _estimate_thickness(self, model, debug):
        extension = self._obj(model, "Extension")
        mass = self._call(extension, "CreateMassProperty2") or self._call(extension, "CreateMassProperty")
        volume = self._positive(self._get(mass, "Volume"))
        area = self._positive(self._get(mass, "SurfaceArea"))
        debug.append(f"mass properties: volume={volume} m3, surface={area} m2")
        if not volume or not area:
            return None
        raw_mm = 2.0 * volume / area * 1000.0
        self.last_thickness_raw_mm = round(raw_mm, 3)
        return self.snap_thickness_mm(raw_mm) / 1000.0

    def find_thickness(self, model, feature, features, debug):
        if feature is not None:
            value = self._thickness_from_definition(model, feature)
            debug.append(f"definition.Thickness={value}")
            if value:
                self.last_thickness_source = "SolidWorks"
                self.last_thickness_raw_mm = round(value * 1000, 4)
                return value

        names = []
        if feature is not None:
            name = self._get(feature, "Name")
            if name:
                names.append(str(name))
        names += ["Sheet-Metal1", "Sheet-Metal"]
        for name in dict.fromkeys(names):
            value = self._thickness_from_parameter(model, name)
            debug.append(f"Thickness@{name}={value}")
            if value:
                self.last_thickness_source = "SolidWorks"
                self.last_thickness_raw_mm = round(value * 1000, 4)
                return value

        for candidate in features:
            if self._type_name(candidate) in {"sheetmetal", "smbaseflange", "baseflange"}:
                value = self._thickness_from_display_dimensions(model, candidate)
                if value:
                    self.last_thickness_source = "SolidWorks"
                    self.last_thickness_raw_mm = round(value * 1000, 4)
                    return value

        estimated = self._estimate_thickness(model, debug)
        if estimated:
            self.last_thickness_source = "estimated"
            return estimated
        return None

    def analyze_model(self, model):
        debug = []
        self.last_debug = debug
        self.last_method = ""
        self.last_thickness_source = ""
        self.last_thickness_raw_mm = None

        doc_type = self._call(model, "GetType")
        if doc_type is not None and doc_type != SW_DOC_PART:
            return False, None

        features = self.get_features(model)
        feature = self.get_sheet_metal_feature(model, features)
        debug.append(f"features={len(features)}")
        debug.append("SheetMetal feature=" + (str(self._get(feature, "Name")) if feature else "NOT FOUND"))

        if feature is not None:
            self.last_method = "SheetMetal feature"
            is_sm = True
        else:
            body_count = len(self.get_sheet_metal_bodies(model))
            debug.append(f"sheet metal bodies={body_count}")
            is_sm = body_count > 0
            if is_sm:
                self.last_method = "Sheet metal body"

        thickness = self.find_thickness(model, feature, features, debug) if is_sm and self.read_thickness else None
        return is_sm, thickness

    @staticmethod
    def thickness_mm(value):
        return round(float(value) * 1000, 3) if value is not None else None

    def scan_sheet_metal(self, unique_parts):
        total = len(unique_parts)
        sheet_parts, failed, not_sheet_metal = [], [], 0
        print("\n" + "=" * 60 + "\nSTEP 1/2: SCAN SHEET METAL\n" + "=" * 60)

        for index, part in enumerate(unique_parts, 1):
            name = part["name"]
            config = part.get("configuration") or "Default"
            label = f"[{index}/{total}] {name} [{config}] (x{part['quantity']})"
            model = None
            opened = False
            try:
                model, opened, status = self.open_part(part["path"], read_only=True)
                if model is None:
                    failed.append((name, status))
                    print(f"{label} -> FAILED: {status}")
                    continue
                if part.get("configuration"):
                    self.activate_configuration(model, part["configuration"])
                rebuilt, rebuild_msg = self.rebuild(model)
                if not rebuilt:
                    print(f"{label} -> warning: {rebuild_msg}")
                is_sm, thickness = self.analyze_model(model)
                if not is_sm:
                    not_sheet_metal += 1
                    print(f"{label} -> not sheet metal")
                    continue
                thickness_mm = self.thickness_mm(thickness) if self.read_thickness else None
                if self.read_thickness and thickness_mm is None:
                    failed.append((name, "sheet metal detected, thickness not found"))
                    print(f"{label} -> FAILED: thickness not found")
                    for line in self.last_debug[:8]:
                        print(f"    [debug] {line}")
                    continue
                record = {
                    "name": name,
                    "path": str(part["path"]),
                    "quantity": int(part["quantity"]),
                    "configuration": part.get("configuration"),
                    "thickness": thickness_mm,
                    "method": self.last_method,
                    "thickness_source": self.last_thickness_source,
                    "thickness_raw_mm": self.last_thickness_raw_mm,
                }
                sheet_parts.append(record)
                suffix = f", {thickness_mm:g} mm [{self.last_thickness_source}]" if thickness_mm is not None else ""
                print(f"{label} -> SHEET METAL ({self.last_method}){suffix}")
            except Exception as exc:
                failed.append((name, f"scan error: {exc}"))
                print(f"{label} -> FAILED: {exc}")
            finally:
                self.close_if_opened(model, opened)

        groups = {}
        for item in sheet_parts:
            key = item["thickness"] if self.read_thickness else None
            g = groups.setdefault(key, {"thickness": key, "unique_parts": 0, "quantity": 0, "parts": []})
            g["unique_parts"] += 1
            g["quantity"] += item["quantity"]
            g["parts"].append(item)

        print("\n" + "=" * 60 + "\nSCAN SUMMARY\n" + "=" * 60)
        print(f"Unique parts scanned: {total}")
        print(f"Sheet metal:          {len(sheet_parts)}")
        print(f"Not sheet metal:      {not_sheet_metal}")
        print(f"Scan failures:        {len(failed)}")
        print(f"Total quantity:       {sum(x['quantity'] for x in sheet_parts)}")
        for thickness in sorted(k for k in groups if k is not None):
            g = groups[thickness]
            print(f"  {thickness:g} mm | parts={g['unique_parts']} | qty={g['quantity']}")

        return {
            "unique_parts": total,
            "sheet_metal": len(sheet_parts),
            "not_sheet_metal": not_sheet_metal,
            "failed": len(failed),
            "failed_parts": failed,
            "total_quantity": sum(x["quantity"] for x in sheet_parts),
            "sheet_metal_parts": sheet_parts,
            "thickness_groups": groups,
            "csv_path": None,
        }

    def scan_and_export(self, unique_parts, output_folder=None):
        result = self.scan_sheet_metal(unique_parts)
        if output_folder:
            path = Path(output_folder) / f"{self.assembly_path.stem if self.assembly_path else 'assembly'}_sheet_metal.csv"
            if self.export_csv(result, path):
                result["csv_path"] = path
        return result

    def export_csv(self, result, csv_path):
        try:
            csv_path = Path(csv_path)
            csv_path.parent.mkdir(parents=True, exist_ok=True)
            with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["Thickness (mm)", "Part", "Configuration", "Quantity", "Thickness source", "Raw value (mm)"])
                for item in sorted(result["sheet_metal_parts"], key=lambda x: (x["thickness"] or 0, x["name"])):
                    writer.writerow([item["thickness"], item["name"], item.get("configuration") or "", item["quantity"], item.get("thickness_source", ""), item.get("thickness_raw_mm", "")])
            print(f"CSV saved: {csv_path}")
            return True
        except Exception as exc:
            print(f"Could not save CSV: {exc}")
            return False
