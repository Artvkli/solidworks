import os
import re
import shutil
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

from .dwg_exporter import DwgExporter

SW_DOC_DRAWING = 3
SW_OPEN_SILENT = 1


class SwDrawingPdfExporter(DwgExporter):
    """
    For every sheet metal part:
      1) copies master.slddrw and replaces the referenced model,
      2) removes the master's model views,
      3) inserts a flat-pattern view of the target part,
      4) saves the .slddrw.
    """

    def __init__(self, sw_app, assembly_path=None, reconnect=None,
                 master_drawing=None, **legacy_options):
        super().__init__(sw_app, assembly_path=assembly_path, reconnect=reconnect)
        self.master_drawing = Path(master_drawing).resolve() if master_drawing else None
        self._typelib_cache = None

    # =====================================================
    # TYPED API (early binding)
    # =====================================================

    def _typelib_module(self):
        if self._typelib_cache is not None:
            return self._typelib_cache or None
        self._typelib_cache = False
        try:
            from win32com.client import gencache

            exe = self._get(self.sw_app, "GetExecutablePath")
            if not exe:
                return None
            install = Path(str(exe))
            if install.suffix:
                install = install.parent
            for tlb_path in install.glob("sldworks.tlb"):
                tlb = pythoncom.LoadTypeLib(str(tlb_path))
                guid, lcid, _syskind, major, minor, _flags = tlb.GetLibAttr()
                module = gencache.EnsureModule(str(guid), lcid, major, minor)
                if module is not None:
                    self._typelib_cache = module
                    return module
        except Exception:
            pass
        return None

    def _cast(self, obj, interface):
        if obj is None:
            return None
        module = self._typelib_module()
        if module is None:
            return obj
        try:
            iid = module.NamesToIIDMap[interface]
            cls = getattr(module, interface)
            return cls(obj._oleobj_.QueryInterface(iid, pythoncom.IID_IDispatch))
        except Exception:
            return obj

    # =====================================================
    # SMALL HELPERS
    # =====================================================

    def _clear(self, drawing):
        try:
            drawing.ClearSelection2(True)
        except Exception:
            pass

    def _close_doc(self, doc):
        if doc is None:
            return
        for getter in ("GetTitle", "GetPathName"):
            value = self._get(doc, getter)
            if value:
                try:
                    self.sw_app.CloseDoc(str(value))
                except Exception:
                    pass

    def _outline(self, view):
        values = self._call(view, "GetOutline")
        try:
            x0, y0, x1, y1 = (float(v) for v in values[:4])
            return x0, y0, x1, y1
        except Exception:
            return None

    def _model_views(self, drawing):
        """All real views of the drawing (the first 'view' is the sheet itself and is skipped)."""
        views = []
        view = self._call(drawing, "GetFirstView")
        view = self._call(view, "GetNextView") if view is not None else None
        guard = 0
        while view is not None and guard < 100:
            guard += 1
            views.append(view)
            view = self._call(view, "GetNextView")
        return views

    # =====================================================
    # OPEN / SAVE
    # =====================================================

    def _open_drawing(self, path, steps):
        try:
            return self.sw_app.OpenDoc6(
                str(path), SW_DOC_DRAWING, SW_OPEN_SILENT, "",
                self._byref_i4(), self._byref_i4(),
            )
        except Exception as exc:
            steps.append(f"opening the .slddrw failed: {str(exc)[:100]}")
            return None

    def _save_drawing(self, drawing, slddrw_path, steps):
        before = slddrw_path.stat().st_mtime if slddrw_path.exists() else 0
        saved = False
        try:
            result = drawing.Save3(1, self._byref_i4(), self._byref_i4())
            saved = bool(result) or (slddrw_path.exists() and slddrw_path.stat().st_mtime > before)
        except Exception as exc:
            steps.append(f"Save3 failed ({str(exc)[:60]})")
        if not saved:
            try:
                drawing.SaveAs3(str(slddrw_path), 0, 1)
                saved = self._written(slddrw_path)
            except Exception as exc:
                steps.append(f"SaveAs3 failed ({str(exc)[:60]})")
        steps.append("saved .slddrw" if saved else "save .slddrw: NOT confirmed")
        return saved

    # =====================================================
    # MODEL REPLACEMENT
    # =====================================================

    def _replace_model(self, drawing_path, old_ref, new_ref, steps):
        """The drawing must be CLOSED."""
        try:
            result = self.sw_app.ReplaceReferencedDocument(str(drawing_path), str(old_ref), str(new_ref))
            steps.append(f"model replaced: result={result}")
            return result in (None, True, 1)
        except Exception as exc:
            steps.append(f"ReplaceReferencedDocument failed ({type(exc).__name__}: {str(exc)[:80]})")
            return False

    def _replace_model_in_open_drawing(self, drawing, views, new_ref, steps):
        """Fallback: replace the model while the drawing is open."""
        names = [str(self._call(v, "GetName2") or "") for v in views]
        attempts = (
            lambda: drawing.ReplaceViewModel(str(new_ref), tuple(names), tuple([""] * len(names))),
            lambda: drawing.ReplaceViewModel(
                str(new_ref),
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, tuple(names)),
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, tuple([""] * len(names)))),
        )
        for attempt in attempts:
            try:
                result = attempt()
                steps.append(f"model replaced (ReplaceViewModel): result={result}")
                if result in (None, True, 1):
                    return True
            except Exception as exc:
                steps.append(f"ReplaceViewModel failed ({type(exc).__name__}: {str(exc)[:70]})")
        return False

    # =====================================================
    # MASTER CLEANUP + FLAT PATTERN
    # =====================================================

    def _delete_model_views(self, drawing, steps):
        """Delete every model view of the copied master."""
        drawing_doc = self._cast(drawing, "IDrawingDoc")
        removed, failed = 0, []
        for _ in range(len(self._model_views(drawing)) + 2):
            views = self._model_views(drawing)
            if not views:
                break
            name = str(self._call(views[-1], "GetName2") or "")
            if not name:
                failed.append("unnamed view")
                break
            try:
                self._clear(drawing)
                if drawing_doc.DeleteView(name) is False:
                    failed.append(name)
                    break
                removed += 1
                self._call(drawing, "ForceRebuild3", False)
            except Exception as exc:
                failed.append(f"{name} ({type(exc).__name__}: {str(exc)[:80]})")
                break
        self._clear(drawing)
        steps.append(f"master views removed: {removed}" + (f"; failed: {', '.join(failed)}" if failed else ""))
        return not failed

    def _create_flat_pattern_view(self, drawing, model_path, configuration, slot, steps):
        """Insert a new flat-pattern view from the target part."""
        x, y = (float(slot["box"][0]), float(slot["box"][1])) if slot else (0.15, 0.15)
        config_name = str(configuration or "")
        drawing_doc = self._cast(drawing, "IDrawingDoc")
        model_name = str(Path(model_path).resolve())
        view, errors = None, []

        for cfg in dict.fromkeys((config_name, "")):
            try:
                view = drawing_doc.CreateFlatPatternViewFromModelView3(
                    model_name, cfg, x, y, 0.0, False, False
                )
                if view is not None:
                    steps.append(f"flat-pattern view created (configuration={cfg!r})")
                    break
                errors.append(f"configuration={cfg!r}: returned None")
            except Exception as exc:
                errors.append(f"configuration={cfg!r}: {type(exc).__name__}: {str(exc)[:90]}")

        if view is None:
            try:
                view = drawing_doc.CreateFlatPatternViewFromModelView2(
                    model_name, config_name, x, y, 0.0, False
                )
                if view is not None:
                    steps.append("flat-pattern view created (CreateFlatPatternViewFromModelView2)")
            except Exception as exc:
                errors.append(f"ViewFromModelView2: {type(exc).__name__}: {str(exc)[:90]}")

        if view is None:
            steps.append("flat-pattern insertion failed: " + " | ".join(errors))
            steps.append("Check that the part is sheet metal, Flat-Pattern is unsuppressed, "
                         "and the configuration can be unfolded.")
            return None

        self._call(drawing, "ForceRebuild3", False)
        for candidate in self._model_views(drawing):  # reacquire to avoid stale COM wrappers
            if bool(self._call(candidate, "IsFlatPatternView", default=False)):
                return candidate
        return view

    # =====================================================
    # ONE PART
    # =====================================================

    def _build_drawing(self, item, slddrw_path):
        steps = []
        path = Path(item["path"]).resolve()
        configuration = item.get("configuration")
        model = None
        opened_by_us = False
        previous_config = None
        drawing = None

        try:
            master = self.master_drawing
            if master is None or not master.is_file():
                return False, f"master drawing not found: {master}", steps
            steps.append(f"master: {master.name}")

            model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                time.sleep(2)
                model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                return False, f"cannot open part: {open_status}", steps
            _, previous_config = self._activate_configuration(model, configuration)

            # ---- 1) copy the master ----
            slddrw_path.parent.mkdir(parents=True, exist_ok=True)
            if slddrw_path.exists():
                slddrw_path.unlink()
            shutil.copy2(str(master), str(slddrw_path))
            try:
                os.chmod(str(slddrw_path), 0o666)
            except Exception:
                pass

            # ---- 2) read the master's model reference and the position of its flat view ----
            doc = self._open_drawing(slddrw_path, steps)
            if doc is None:
                return False, "cannot open the copy of the master", steps
            old_ref, flat_slot = None, None
            for view in self._model_views(doc):
                ref = self._call(view, "GetReferencedModelName")
                if ref and not old_ref:
                    old_ref = str(ref)
                outline = self._outline(view)
                if outline is not None and flat_slot is None \
                        and bool(self._call(view, "IsFlatPatternView", default=False)):
                    flat_slot = {"box": ((outline[0] + outline[2]) / 2, (outline[1] + outline[3]) / 2,
                                         outline[2] - outline[0], outline[3] - outline[1])}
            self._close_doc(doc)
            if not old_ref:
                return False, "no model views found in the master drawing", steps

            # ---- 3) replace the model (drawing closed) ----
            replaced_in_closed = True
            try:
                same_part = Path(old_ref).resolve() == path
            except Exception:
                same_part = False
            if not same_part:
                try:
                    self.sw_app.CloseDoc(old_ref)
                except Exception:
                    pass
                replaced_in_closed = self._replace_model(slddrw_path, old_ref, path, steps)

            # ---- 4) reopen, replace in the open drawing if needed ----
            drawing = self._open_drawing(slddrw_path, steps)
            if drawing is None:
                return False, "cannot reopen the drawing", steps
            if not replaced_in_closed:
                if not self._replace_model_in_open_drawing(drawing, self._model_views(drawing), path, steps):
                    return False, "could not replace the model in the drawing", steps
            self._call(drawing, "ForceRebuild3", False)

            # ---- 5) delete master views, insert the new flat pattern ----
            self._delete_model_views(drawing, steps)
            flat_view = self._create_flat_pattern_view(drawing, path, configuration, flat_slot, steps)
            if flat_view is None:
                return False, "could not create the flat-pattern view", steps
            self._call(drawing, "ForceRebuild3", False)
            self._call(drawing, "ViewZoomtofit2")

            # ---- 6) save the .slddrw ----
            if not self._save_drawing(drawing, slddrw_path, steps):
                return False, ".slddrw was not saved", steps
            return True, "SolidWorks drawing (flat pattern from master)", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
            self._close_doc(drawing)
            if model is not None:
                self._restore_configuration(model, previous_config)
            self._close_part(model, opened_by_us)

    def export_part(self, item, slddrw_path, pdf_path=None):
        # pdf_path is kept only for compatibility with the caller; it is ignored.
        ok, message, steps = self._build_drawing(item, Path(slddrw_path))

        if not ok and self._is_connection_error(RuntimeError(message + " " + " ".join(steps))):
            print("  SolidWorks connection problem. Reconnecting and retrying once...")
            time.sleep(3)
            if self._reconnect():
                ok, message, steps = self._build_drawing(item, Path(slddrw_path))

        return ok, message, steps

    @staticmethod
    def _safe_name(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"