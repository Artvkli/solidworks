import re
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

# swExportFlatPatternViewOptions_e
SW_FLAT_PATTERN_NONE = 0

# swExportToDWG_e
SW_EXPORT_TO_DWG_SHEET_METAL = 1

# ExportToDWG2 sheet metal options (bit flags): geometry + bend lines
SHEET_METAL_OPTIONS = 1 | 4

# swSuppressionState_e / swInConfigurationOpts_e
SW_UNSUPPRESS = 0
SW_SUPPRESS = 1
SW_THIS_CONFIGURATION = 1

# swRebuildOnActivation_e.swDontRebuildActiveDoc
SW_DONT_REBUILD = 1

INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class DwgExporter:
    """Exports the flat pattern of sheet metal parts to DWG."""

    def __init__(self, sw_app, assembly_path=None):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path) if assembly_path else None

    # =====================================================
    # SAFE HELPERS
    # =====================================================

    @staticmethod
    def _read(obj, name, default=None):
        if obj is None:
            return default
        try:
            value = getattr(obj, name)
            if callable(value):
                return value()
            return value
        except Exception:
            return default

    @staticmethod
    def _call(obj, name, *args, default=None):
        """
        Call a COM method safely.
        pywin32 sometimes evaluates no-argument members right away and returns the
        VALUE (a string, a number or a COM object) instead of a callable. Those
        values are returned as they are.
        """
        if obj is None:
            return default
        try:
            attr = getattr(obj, name)
        except Exception:
            return default
        try:
            if args:
                result = attr(*args) if callable(attr) else default
            elif hasattr(attr, "_oleobj_"):  # COM object already evaluated
                result = attr
            elif callable(attr):
                result = attr()
            else:  # plain value already evaluated
                result = attr
        except Exception:
            return default
        return default if result is None else result

    @staticmethod
    def _safe_filename(name):
        cleaned = INVALID_FILENAME.sub("_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    @staticmethod
    def _byref_long():
        if VARIANT is not None:
            return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        return 0

    # =====================================================
    # ACTIVATE DOCUMENT
    # =====================================================

    def activate(self, model):
        """Make a document the active one. Returns True if SolidWorks accepted it."""
        candidates = [
            self._read(model, "GetTitle"),
            self._read(model, "GetPathName"),
        ]

        for name in candidates:
            if not name:
                continue
            try:
                active = self.sw_app.ActivateDoc3(
                    str(name), False, SW_DONT_REBUILD, self._byref_long()
                )
                if active is not None:
                    return True
            except Exception:
                continue

        return False

    def activate_assembly(self, assembly_model=None):
        """Bring the assembly back to the front after exporting."""
        if assembly_model is not None:
            return self.activate(assembly_model)

        if self.assembly_path:
            try:
                self.sw_app.ActivateDoc3(
                    str(self.assembly_path), False, SW_DONT_REBUILD, self._byref_long()
                )
                return True
            except Exception:
                return False
        return False

    # =====================================================
    # FLAT PATTERN FEATURE (suppression handling)
    # =====================================================

    def _flat_pattern_feature(self, model):
        feature = self._call(model, "FeatureByName", "Flat-Pattern1")
        if feature is not None:
            return feature

        # fallback: search the feature tree by type
        first = self._call(model, "FirstFeature")
        count = 0
        while first is not None and count < 100000:
            type_name = self._call(first, "GetTypeName2")
            if type_name and str(type_name).lower() == "flatpattern":
                return first
            first = self._call(first, "GetNextFeature")
            count += 1

        return None

    def _set_suppression(self, feature, state):
        try:
            feature.SetSuppression2(state, SW_THIS_CONFIGURATION, None)
            return True
        except Exception:
            return False

    # =====================================================
    # EXPORT ONE PART
    # =====================================================

    def _export_flat_pattern_view(self, model, dwg_path):
        try:
            model.ExportFlatPatternView(str(dwg_path), SW_FLAT_PATTERN_NONE)
        except Exception:
            pass
        return self._file_written(dwg_path)

    def _export_to_dwg2(self, model, dwg_path):
        if VARIANT is None:
            return False

        try:
            alignment = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8,
                (0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0),
            )
            views = VARIANT(pythoncom.VT_DISPATCH, None)

            model.ExportToDWG2(
                str(dwg_path),
                str(self._read(model, "GetPathName") or ""),
                SW_EXPORT_TO_DWG_SHEET_METAL,
                True,
                alignment,
                False,
                False,
                SHEET_METAL_OPTIONS,
                views,
            )
        except Exception:
            pass

        return self._file_written(dwg_path)

    @staticmethod
    def _file_written(dwg_path):
        try:
            return Path(dwg_path).exists() and Path(dwg_path).stat().st_size > 0
        except Exception:
            return False

    def export_part(self, model, dwg_path):
        """Returns (ok, message)."""
        dwg_path = Path(dwg_path)

        # never trust an old file: remove it so success means 'freshly written'
        try:
            if dwg_path.exists():
                dwg_path.unlink()
        except Exception:
            return (
                False,
                "could not overwrite existing file (is it open in another program?)",
            )

        self.activate(model)

        # Flat pattern must be un-suppressed for the export
        flat_pattern = self._flat_pattern_feature(model)
        was_suppressed = False
        if flat_pattern is not None:
            was_suppressed = bool(self._read(flat_pattern, "IsSuppressed", False))
            if was_suppressed:
                self._set_suppression(flat_pattern, SW_UNSUPPRESS)

        try:
            if self._export_flat_pattern_view(model, dwg_path):
                return True, "ExportFlatPatternView"

            if self._export_to_dwg2(model, dwg_path):
                return True, "ExportToDWG2"

            reason = "flat pattern export failed"
            if flat_pattern is None:
                reason += " (Flat-Pattern feature not found)"
            return False, reason

        finally:
            # put the flat pattern back the way it was
            if flat_pattern is not None and was_suppressed:
                self._set_suppression(flat_pattern, SW_SUPPRESS)

    # =====================================================
    # EXPORT ALL
    # =====================================================

    def export_all(self, sheet_metal_parts, output_folder):
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        total = len(sheet_metal_parts)
        exported = []
        failed = []
        used_names = set()

        print()
        print("=" * 60)
        print("DWG EXPORT (FLAT PATTERNS)")
        print("=" * 60)
        print(f"Parts to export: {total}")
        print(f"Output folder:   {output_folder}")
        print()

        for index, item in enumerate(
            sorted(sheet_metal_parts, key=lambda i: i["name"]), start=1
        ):
            name = item["name"]
            label = f"[{index}/{total}] {name}"

            model = item.get("model")
            if model is None:
                print(f"{label} -> FAILED: no model")
                failed.append((name, "no model"))
                continue

            base = self._safe_filename(name)
            unique = base
            counter = 2
            while unique.lower() in used_names:
                unique = f"{base}_{counter}"
                counter += 1
            used_names.add(unique.lower())

            dwg_path = output_folder / f"{unique}.dwg"

            try:
                ok, message = self.export_part(model, dwg_path)
            except Exception as e:
                ok, message = False, f"error: {e}"

            if ok:
                print(f"{label} -> OK ({message})")
                exported.append((name, dwg_path.name, item["quantity"]))
            else:
                print(f"{label} -> FAILED: {message}")
                failed.append((name, message))

        print()
        print("=" * 60)
        print("DWG EXPORT SUMMARY")
        print("=" * 60)
        print(f"Exported: {len(exported)} / {total}")
        print(f"Failed:   {len(failed)}")

        if failed:
            print()
            print("Failed parts:")
            for name, message in failed:
                print(f"  - {name}: {message}")

        report_path = self._write_report(output_folder, exported, failed)

        return {
            "exported": len(exported),
            "failed": len(failed),
            "exported_parts": exported,
            "failed_parts": failed,
            "report_path": report_path,
        }

    def _write_report(self, output_folder, exported, failed):
        import csv

        report_path = output_folder / "dwg_export_report.csv"
        try:
            with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["Part", "Quantity", "DWG file", "Status"])
                for name, filename, quantity in exported:
                    writer.writerow([name, quantity, filename, "OK"])
                for name, message in failed:
                    writer.writerow([name, "", "", f"FAILED: {message}"])
            return report_path
        except Exception as e:
            print(f"Could not save DWG report: {e}")
            return None
