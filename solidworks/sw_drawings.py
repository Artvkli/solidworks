"""Creates one SLDDRW per sheet-metal part containing its flat pattern view."""
import gc
import re
import shutil
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

SW_DOC_PART = 1
SW_DOC_DRAWING = 3
SW_OPEN_SILENT = 1
SW_SAVE_SILENT = 1
SW_SAVE_AS_CURRENT_VERSION = 0
SW_UNSUPPRESS = 0
SW_SUPPRESS = 1
SW_THIS_CONFIGURATION = 1
SW_PAPER_A3 = 8                       # swDwgPaperA3size
SW_DEFAULT_TEMPLATE_DRAWING = 10      # swUserPreferenceStringValue_e.swDefaultTemplateDrawing
FALLBACK_SHEET = (0.420, 0.297)       # A3 landscape, meters
FIT_RATIO = 0.70                      # flat view uses at most 70% of the sheet
STANDARD_SCALES = (10, 5, 4, 2, 1, 0.5, 0.4, 0.25, 0.2, 0.1, 0.05, 0.04, 0.02, 0.01, 0.005)
INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')




class SwDrawingExporter:
    """For every part: open part -> new drawing -> flat pattern view -> save .SLDDRW."""

    def __init__(self, sw_app, assembly_path=None, reconnect=None, template_path=None,
                 sheet_format_path=None, text_replacements=None, master_drawing=None,
                 view_width_ratio=0.5, view_left_mm=25, view_center_y_ratio=0.42):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path).resolve() if assembly_path else None
        self.reconnect = reconnect
        self.template_path = Path(template_path) if template_path else None
        self.master_drawing = Path(master_drawing) if master_drawing else None
        self.sheet_format_path = Path(sheet_format_path) if sheet_format_path else None
        self.text_replacements = dict(text_replacements or {})
        self.view_width_ratio = view_width_ratio
        self.view_left_mm = view_left_mm
        self.view_center_y_ratio = view_center_y_ratio

    # ---------------- helpers ----------------
    @staticmethod
    def _safe_name(name):
        cleaned = INVALID_FILENAME.sub("_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    @staticmethod
    def _byref_i4():
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

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
    def _written(path):
        try:
            p = Path(path)
            return p.exists() and p.stat().st_size > 0
        except Exception:
            return False

    @staticmethod
    def _is_connection_error(text):
        text = str(text).lower()
        return any(t in text for t in (
            "rpc server is unavailable", "remote procedure call failed",
            "server execution failed", "object invoked has disconnected",
            "disconnected from its clients", "catastrophic failure",
        ))

    def _reconnect(self):
        if not self.reconnect:
            return False
        try:
            app = self.reconnect()
            if app is not None:
                self.sw_app = app
                return True
        except Exception as exc:
            print(f"  Reconnect failed: {exc}")
        return False

    def _find_template(self):
        if self.template_path and self.template_path.is_file():
            return str(self.template_path.resolve())
        try:
            tpl = self.sw_app.GetUserPreferenceStringValue(SW_DEFAULT_TEMPLATE_DRAWING)
            if tpl and Path(tpl).is_file():
                return str(tpl)
        except Exception:
            pass
        return None

    def _close_doc(self, doc):
        if doc is None:
            return
        try:
            title = self._get(doc, "GetTitle")
            if title:
                self.sw_app.CloseDoc(str(title))
        except Exception:
            pass

    def _open_part(self, path):
        path = str(Path(path).resolve())
        try:
            existing = self.sw_app.GetOpenDocumentByName(path)
            if existing is not None:
                return existing, False, "already open"
        except Exception:
            pass
        errors, warnings = self._byref_i4(), self._byref_i4()
        model = self.sw_app.OpenDoc6(path, SW_DOC_PART, SW_OPEN_SILENT, "", errors, warnings)
        if model is None:
            return None, False, f"OpenDoc6 returned None (errors={getattr(errors, 'value', errors)})"
        return model, True, "opened"

    def _find_flat_pattern(self, model):
        for name in ("Flat-Pattern1", "Flat-Pattern", "FlatPattern"):
            try:
                feature = model.FeatureByName(name)
            except Exception:
                feature = None
            if feature is not None:
                return feature
        current = self._get(model, "FirstFeature")
        for _ in range(100000):
            if current is None:
                break
            if str(self._get(current, "GetTypeName2") or "").lower() == "flatpattern":
                return current
            current = self._get(current, "GetNextFeature")
        return None

    def _sheet_size(self, drawing):
        try:
            sheet = self._get(drawing, "GetCurrentSheet")
            props = self._get(sheet, "GetProperties2")
            if props and len(props) >= 7 and props[5] > 0 and props[6] > 0:
                return float(props[5]), float(props[6])
        except Exception:
            pass
        return FALLBACK_SHEET

    def _create_flat_view(self, drawing, part_path, config, x, y, steps):
        attempts = (
            ("CreateFlatPatternViewFromModelView3",
             lambda cfg: drawing.CreateFlatPatternViewFromModelView3(part_path, cfg, x, y, 0.0, False, False)),
            ("CreateDrawViewFromModelView3(*Flat pattern)",
             lambda cfg: drawing.CreateDrawViewFromModelView3(part_path, "*Flat pattern", x, y, 0.0)),
        )
        for round_no in range(1, 4):
            for name, call in attempts:
                for cfg in dict.fromkeys((config or "", "")):
                    try:
                        view = call(cfg)
                        if view is not None:
                            if round_no > 1:
                                steps.append(f"view created on attempt {round_no}")
                            return view, name
                        steps.append(f"attempt {round_no}: {name} (cfg='{cfg}') returned None")
                    except Exception as exc:
                        steps.append(f"attempt {round_no}: {name} (cfg='{cfg}') {type(exc).__name__}: {exc}")
            time.sleep(1.5)
            try:
                title = self._get(drawing, "GetTitle")
                self.sw_app.ActivateDoc3(str(title), False, 0, self._byref_i4())
            except Exception:
                pass
        return None, None

    def _fit_view(self, drawing, view):
        """Scale so the view is ~half the sheet width, then place it at the left, slightly low."""
        width, height = self._sheet_size(drawing)
        try:
            outline = self._get(view, "GetOutline")
            vw, vh = outline[2] - outline[0], outline[3] - outline[1]
            current = float(self._get(view, "ScaleDecimal"))
            if vw <= 0 or vh <= 0 or current <= 0:
                return "fit skipped (empty outline)"

            wanted = current * min(width * self.view_width_ratio / vw, height * 0.6 / vh)
            chosen = next((s for s in STANDARD_SCALES if s <= wanted * 1.001), STANDARD_SCALES[-1])
            view.ScaleDecimal = chosen

            outline = self._get(view, "GetOutline")
            new_w = outline[2] - outline[0]
            target_x = self.view_left_mm / 1000.0 + new_w / 2.0
            target_y = height * self.view_center_y_ratio

            for _ in range(2):
                outline = self._get(view, "GetOutline")
                cx = (outline[0] + outline[2]) / 2.0
                cy = (outline[1] + outline[3]) / 2.0
                pos = self._get(view, "Position")
                view.Position = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                        (pos[0] + target_x - cx, pos[1] + target_y - cy))

            outline = self._get(view, "GetOutline")
            scale_txt = f"1:{1 / chosen:g}" if chosen < 1 else f"{chosen:g}:1"
            return (f"scale {scale_txt}, width {(outline[2] - outline[0]) / width * 100:.0f}% of sheet, "
                    f"left edge {outline[0] * 1000:.1f} mm, center y {((outline[1] + outline[3]) / 2) * 1000:.1f} mm")
        except Exception as exc:
            return f"fit skipped ({exc})"


    def _replace_sheet_texts(self, drawing, values):
        """Inside Edit Sheet Format: replaces placeholders such as {PART_NAME} in the notes."""
        if not self.text_replacements:
            return 0
        count = 0
        view = self._get(drawing, "GetFirstView")      # in Edit Sheet Format the first view is the sheet format
        note = self._get(view, "GetFirstNote")
        for _ in range(10000):
            if note is None:
                break
            try:
                text = self._get(note, "GetText")
                if text:
                    new = str(text)
                    for placeholder, field in self.text_replacements.items():
                        new = new.replace(placeholder, str(values.get(field, "")))
                    if new != text:
                        note.SetText(new)
                        count += 1
            except Exception:
                pass
            note = self._get(note, "GetNext")
        return count

    def _apply_sheet_format(self, drawing, item, steps):
        """Loads the sheet format, writes drawing properties and runs Edit Sheet Format."""
        values = {
            "PartName": item.get("name") or Path(item["path"]).stem,
            "Quantity": item.get("quantity", 1),
            "Thickness": item.get("thickness"),
            "Material": item.get("material"),
            "Configuration": item.get("configuration"),
        }

        # 1) sheet format file (.slddrt); without it the format inside the .drwdot is used
        if self.sheet_format_path:
            if not self.sheet_format_path.is_file():
                steps.append(f"sheet format file not found: {self.sheet_format_path}")
            else:
                try:
                    sheet = self._get(drawing, "GetCurrentSheet")
                    sheet.SetTemplateName(str(self.sheet_format_path.resolve()))
                    sheet.ReloadTemplate(True)
                    steps.append(f"sheet format loaded: {self.sheet_format_path.name}")
                except Exception as exc:
                    steps.append(f"sheet format not loaded: {exc}")

        # 2) drawing custom properties (title block can link to them: $PRP:"PartName")
        try:
            manager = drawing.Extension.CustomPropertyManager("")
            for key, value in values.items():
                if value not in (None, ""):
                    manager.Add3(key, 30, str(value), 2)     # 30 = text, 2 = overwrite
            steps.append("drawing properties written")
        except Exception as exc:
            steps.append(f"drawing properties skipped: {exc}")

        # 3) Edit Sheet Format -> replace placeholders -> ALWAYS back to sheet
        try:
            drawing.EditTemplate()          # the return value is unreliable (None), so it is ignored
            replaced = self._replace_sheet_texts(drawing, values)
            steps.append(f"Edit Sheet Format: {replaced} text(s) replaced")
        except Exception as exc:
            steps.append(f"Edit Sheet Format failed: {exc}")
        finally:
            try:
                drawing.EditSheet()         # harmless if not in template mode
                steps.append("back to sheet (EditSheet)")
            except Exception as exc:
                steps.append(f"EditSheet failed: {exc}")

    def _open_drawing(self, path):
        errors, warnings = self._byref_i4(), self._byref_i4()
        return self.sw_app.OpenDoc6(str(path), SW_DOC_DRAWING, SW_OPEN_SILENT, "", errors, warnings)

    def _remove_existing_views(self, drawing, steps):
        """Deletes everything that came from master.slddrw: bend tables, flat pattern and other views."""
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)

        views = []
        view = self._get(self._get(drawing, "GetFirstView"), "GetNextView")   # skip the sheet-format view
        while view is not None:
            views.append(view)
            view = self._get(view, "GetNextView")
        names = [str(self._get(v, "GetName2")) for v in views]
        steps.append(f"master views found: {names}")

        # 1) tables first (bend table, hole table...) - they belong to the views
        tables_removed = 0
        for v in views:
            tables = self._get(v, "GetTableAnnotations")
            for table in (tables or ()):
                try:
                    annotation = self._get(table, "GetAnnotation")
                    drawing.ClearSelection2(True)
                    if annotation is not None and annotation.Select3(False, nothing):
                        drawing.Extension.DeleteSelection2(6)      # 2 = absorbed, 4 = children
                        tables_removed += 1
                except Exception as exc:
                    steps.append(f"table delete failed: {exc}")
        steps.append(f"tables removed: {tables_removed}")

        # 2) then the views themselves
        views_removed = 0
        for name in names:
            try:
                drawing.ClearSelection2(True)
                if drawing.Extension.SelectByID2(name, "DRAWINGVIEW", 0, 0, 0, False, 0, nothing, 0):
                    drawing.Extension.DeleteSelection2(6)
                    views_removed += 1
                else:
                    steps.append(f"could not select view: {name}")
            except Exception as exc:
                steps.append(f"view delete failed ({name}): {exc}")
        try:
            drawing.ClearSelection2(True)
        except Exception:
            pass
        steps.append(f"old views removed: {views_removed}/{len(names)}")

    # ---------------- public ----------------
    def export_part(self, item, slddrw_path):
        """Returns (ok, message, steps)."""
        ok, message, steps = self._export_once(item, slddrw_path)
        if not ok and self._is_connection_error(message):
            steps.append("connection lost -> reconnecting and retrying once")
            time.sleep(3)
            if self._reconnect():
                ok, message, steps2 = self._export_once(item, slddrw_path)
                steps += steps2
        if not ok:
            try:
                Path(slddrw_path).unlink(missing_ok=True)   # do not leave a master copy without flat view
            except Exception:
                pass
        return ok, message, steps

    def _export_once(self, item, slddrw_path):
        steps = []
        slddrw_path = Path(slddrw_path)
        slddrw_path.parent.mkdir(parents=True, exist_ok=True)
        if slddrw_path.exists():
            try:
                slddrw_path.unlink()
            except Exception as exc:
                return False, f"cannot remove old file: {exc}", steps

        part_path = str(Path(item["path"]).resolve())
        configuration = item.get("configuration")
        model = drawing = flat = None
        opened_by_us = False
        was_suppressed = False
        previous_config = None

        try:
            if not self.master_drawing or not self.master_drawing.is_file():
                return False, f"master drawing not found: {self.master_drawing}", steps

            model, opened_by_us, status = self._open_part(part_path)
            steps.append(f"part {status}")
            if model is None:
                return False, status, steps

            if configuration:
                try:
                    previous_config = self._get(
                        self._get(model.ConfigurationManager, "ActiveConfiguration"), "Name")
                    model.ShowConfiguration2(str(configuration))
                    steps.append(f"configuration: {configuration}")
                except Exception as exc:
                    steps.append(f"configuration not activated: {exc}")

            try:
                model.ForceRebuild3(True)
            except Exception:
                try:
                    model.EditRebuild3()
                except Exception:
                    pass

            flat = self._find_flat_pattern(model)
            if flat is None:
                return False, "Flat-Pattern feature not found", steps
            was_suppressed = bool(self._get(flat, "IsSuppressed", False))
            if was_suppressed:
                flat.SetSuppression2(SW_UNSUPPRESS, SW_THIS_CONFIGURATION, None)
                steps.append("flat pattern unsuppressed")

            shutil.copyfile(self.master_drawing, slddrw_path)
            steps.append(f"master copied: {self.master_drawing.name}")
            drawing = self._open_drawing(slddrw_path)
            if drawing is None:
                return False, "could not open the copied master drawing", steps
            self._apply_sheet_format(drawing, item, steps)      # Edit Sheet Format -> back to sheet
            self._remove_existing_views(drawing, steps)
            width, height = self._sheet_size(drawing)

            active_cfg = configuration or self._get(
                self._get(model.ConfigurationManager, "ActiveConfiguration"), "Name") or ""
            try:
                self.sw_app.ActivateDoc3(str(self._get(drawing, "GetTitle")), False, 0, self._byref_i4())
            except Exception:
                pass
            model_path = str(self._get(model, "GetPathName") or part_path)
            view, method = self._create_flat_view(drawing, model_path, active_cfg, width / 2, height / 2, steps)
            if view is None:
                return False, "could not create the flat pattern view", steps
            steps.append(f"flat view created ({method})")
            steps.append(self._fit_view(drawing, view))
            try:
                steps.append(f"view: {self._get(view, 'GetName2')}, outline: {tuple(round(v, 4) for v in self._get(view, 'GetOutline'))}")
            except Exception:
                pass

            try:
                drawing.ViewZoomtofit2()
            except Exception:
                pass

            try:
                drawing.ForceRebuild3(False)      # refresh linked title block fields
            except Exception:
                pass

            err = drawing.SaveAs3(str(slddrw_path), SW_SAVE_AS_CURRENT_VERSION, SW_SAVE_SILENT)
            if not self._written(slddrw_path):
                return False, f"SaveAs3 failed (code {err})", steps
            steps.append("SLDDRW saved")
            return True, "ok", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
            # close the drawing first, then restore and close the part
            self._close_doc(drawing)
            if flat is not None and was_suppressed:
                try:
                    flat.SetSuppression2(SW_SUPPRESS, SW_THIS_CONFIGURATION, None)
                except Exception:
                    pass
            if model is not None and previous_config:
                try:
                    model.ShowConfiguration2(str(previous_config))
                except Exception:
                    pass
            if opened_by_us:
                self._close_doc(model)
            del drawing, model, flat
            gc.collect()
            try:
                pythoncom.CoFreeUnusedLibraries()
            except Exception:
                pass