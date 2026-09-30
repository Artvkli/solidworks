import gc
import json
import re
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

SW_DOC_PART = 1
SW_OPEN_SILENT = 1
SW_EXPORT_TO_DWG_SHEET_METAL = 1
SHEET_METAL_OPTIONS = 1 | 4  # geometry + bend lines
SW_DONT_REBUILD = 1
SW_UNSUPPRESS = 0
SW_SUPPRESS = 1
SW_THIS_CONFIGURATION = 1
INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class DwgExporter:
    """Stable per-part SolidWorks sheet-metal exporter.

    Important design choice: this class never receives/retains Part COM objects from
    the scanner. Each part is opened only for its export and then closed if this
    process opened it. That keeps the SolidWorks COM graph small and recoverable.
    """

    def __init__(self, sw_app, assembly_path=None, reconnect=None):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path).resolve() if assembly_path else None
        self.reconnect = reconnect
        self.last_error = None

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
    def _safe_filename(name):
        cleaned = INVALID_FILENAME.sub("_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    @staticmethod
    def _byref_i4():
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    @staticmethod
    def _written(path):
        try:
            p = Path(path)
            return p.exists() and p.stat().st_size > 0
        except Exception:
            return False

    @staticmethod
    def _is_connection_error(exc):
        text = str(exc).lower()
        return any(
            token in text
            for token in (
                "rpc server is unavailable",
                "remote procedure call failed",
                "server execution failed",
                "object invoked has disconnected",
                "disconnected from its clients",
                "automation error",
                "catastrophic failure",
            )
        )

    def _reconnect(self):
        if not self.reconnect:
            return False
        try:
            new_app = self.reconnect()
            if new_app is not None:
                self.sw_app = new_app
                return True
        except Exception as exc:
            print(f"  Reconnect failed: {exc}")
        return False

    def _open_part(self, path):
        path = Path(path).resolve()
        try:
            existing = self.sw_app.GetOpenDocumentByName(str(path))
            if existing is not None:
                return existing, False, "already open"
        except Exception:
            pass

        errors = self._byref_i4()
        warnings = self._byref_i4()
        try:
            model = self.sw_app.OpenDoc6(
                str(path), SW_DOC_PART, SW_OPEN_SILENT, "", errors, warnings
            )
            if model is None:
                return (
                    None,
                    False,
                    f"OpenDoc6 returned None (errors={getattr(errors, 'value', errors)}, warnings={getattr(warnings, 'value', warnings)})",
                )
            return (
                model,
                True,
                f"opened (errors={getattr(errors, 'value', errors)}, warnings={getattr(warnings, 'value', warnings)})",
            )
        except Exception as exc:
            if self._is_connection_error(exc) and self._reconnect():
                errors = self._byref_i4()
                warnings = self._byref_i4()
                try:
                    model = self.sw_app.OpenDoc6(
                        str(path), SW_DOC_PART, SW_OPEN_SILENT, "", errors, warnings
                    )
                    if model is not None:
                        return model, True, "opened after SolidWorks reconnect"
                except Exception as exc2:
                    return None, False, f"OpenDoc6 after reconnect failed: {exc2}"
            return None, False, f"OpenDoc6 exception: {exc}"

    def _close_part(self, model, opened_by_us):
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

    def _activate_configuration(self, model, configuration):
        if not configuration:
            return True, None
        manager = self._call(model, "ConfigurationManager")
        active = self._call(manager, "ActiveConfiguration")
        previous = self._get(active, "Name")
        try:
            result = model.ShowConfiguration2(str(configuration))
            return bool(result), previous
        except Exception:
            return False, previous

    def _restore_configuration(self, model, previous):
        if previous:
            try:
                model.ShowConfiguration2(str(previous))
            except Exception:
                pass

    def _rebuild(self, model):
        try:
            return bool(model.ForceRebuild3(True)), "ForceRebuild3"
        except Exception as exc:
            try:
                return bool(model.EditRebuild3()), "EditRebuild3"
            except Exception as exc2:
                return False, f"rebuild failed: {exc}; fallback failed: {exc2}"

    def _find_flat_pattern(self, model):
        for name in ("Flat-Pattern1", "Flat-Pattern", "FlatPattern"):
            feature = self._call(model, "FeatureByName", name)
            if feature is not None:
                return feature, name
        current = self._call(model, "FirstFeature")
        for _ in range(100000):
            if current is None:
                break
            type_name = str(self._call(current, "GetTypeName2") or "").lower()
            if type_name == "flatpattern":
                return current, str(self._get(current, "Name") or "<unnamed>")
            current = self._call(current, "GetNextFeature")
        return None, None

    def _export_to_dwg2(self, model, output_path):
        alignment = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0),
        )
        model_name = str(self._get(model, "GetPathName") or "")
        attempts = []
        for views in (None, VARIANT(pythoncom.VT_DISPATCH, None)):
            try:
                result = model.ExportToDWG2(
                    str(output_path),
                    model_name,
                    SW_EXPORT_TO_DWG_SHEET_METAL,
                    True,
                    alignment,
                    False,
                    False,
                    SHEET_METAL_OPTIONS,
                    views,
                )
                if bool(result) and self._written(output_path):
                    return True, f"ExportToDWG2 returned {result}"
                attempts.append(
                    f"returned {result!r}, file={self._written(output_path)}"
                )
            except Exception as exc:
                attempts.append(f"{type(exc).__name__}: {exc}")
        return False, "ExportToDWG2 failed: " + " | ".join(attempts)

    def _export_flat_pattern_fallback(self, model, output_path, flat_pattern):
        if flat_pattern is None:
            return False, "fallback unavailable: Flat-Pattern feature not found"
        was_suppressed = False
        try:
            was_suppressed = bool(self._get(flat_pattern, "IsSuppressed", False))
            if was_suppressed:
                try:
                    flat_pattern.SetSuppression2(
                        SW_UNSUPPRESS, SW_THIS_CONFIGURATION, None
                    )
                except Exception as exc:
                    return False, f"could not unsuppress Flat-Pattern: {exc}"
            try:
                result = model.ExportFlatPatternView(str(output_path), 0)
                if self._written(output_path):
                    return True, f"ExportFlatPatternView returned {result}"
                return (
                    False,
                    f"ExportFlatPatternView returned {result!r} but created no file",
                )
            except Exception as exc:
                return False, f"ExportFlatPatternView exception: {exc}"
        finally:
            if was_suppressed:
                try:
                    flat_pattern.SetSuppression2(
                        SW_SUPPRESS, SW_THIS_CONFIGURATION, None
                    )
                except Exception:
                    pass

    # =====================================================
    # EXPORT ONE PART (with one retry after a lost connection)
    # =====================================================

    def export_part(self, item, output_path):
        ok, message = self._export_part_once(item, output_path, attempt=1)

        if not ok and self._is_connection_error(RuntimeError(message)):
            print(
                "  SolidWorks connection problem. Reconnecting and retrying this part once..."
            )
            time.sleep(3)
            if self._reconnect():
                ok, message = self._export_part_once(item, output_path, attempt=2)

        return ok, message

    def _export_part_once(self, item, output_path, attempt=1):
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            try:
                output_path.unlink()
            except Exception as exc:
                return False, f"cannot remove old output: {exc}"

        path = Path(item["path"]).resolve()
        configuration = item.get("configuration")
        model = None
        opened_by_us = False
        previous_config = None
        diagnostics = {
            "part": item.get("name"),
            "path": str(path),
            "configuration": configuration,
            "attempt": attempt,
        }
        started = time.time()

        try:
            model, opened_by_us, open_status = self._open_part(path)
            diagnostics["open"] = open_status
            if model is None:
                return False, open_status

            config_ok, previous_config = self._activate_configuration(
                model, configuration
            )
            diagnostics["configuration_ok"] = config_ok
            diagnostics["active_configuration"] = self._get(
                self._call(
                    self._call(model, "ConfigurationManager"), "ActiveConfiguration"
                ),
                "Name",
            )
            if configuration and not config_ok:
                diagnostics["warning"] = (
                    f"could not activate configuration {configuration}"
                )

            rebuilt, rebuild_message = self._rebuild(model)
            diagnostics["rebuild"] = rebuild_message
            diagnostics["rebuild_ok"] = rebuilt

            flat_pattern, flat_name = self._find_flat_pattern(model)
            diagnostics["flat_pattern"] = flat_name

            ok, message = self._export_to_dwg2(model, output_path)
            diagnostics["primary_export"] = message
            if not ok:
                fallback_ok, fallback_message = self._export_flat_pattern_fallback(
                    model, output_path, flat_pattern
                )
                diagnostics["fallback_export"] = fallback_message
                if fallback_ok:
                    diagnostics["elapsed_sec"] = round(time.time() - started, 3)
                    self.last_error = None
                    return True, fallback_message

                diagnostics["error"] = message + "; " + fallback_message
                self.last_error = diagnostics
                return False, self._format_failure(diagnostics)

            diagnostics["elapsed_sec"] = round(time.time() - started, 3)
            self.last_error = None
            return True, message

        except Exception as exc:
            diagnostics["exception"] = f"{type(exc).__name__}: {exc}"
            diagnostics["elapsed_sec"] = round(time.time() - started, 3)
            self.last_error = diagnostics
            return False, self._format_failure(diagnostics)
        finally:
            if model is not None:
                self._restore_configuration(model, previous_config)
            self._close_part(model, opened_by_us)

    @staticmethod
    def _format_failure(d):
        return json.dumps(d, ensure_ascii=False, separators=(", ", ": "))

    def activate_assembly(self):
        if not self.assembly_path:
            return False
        try:
            self.sw_app.ActivateDoc3(
                str(self.assembly_path), False, SW_DONT_REBUILD, self._byref_i4()
            )
            return True
        except Exception:
            return False
    