"""One SLDDRW per sheet-metal part:
master.slddrw + Top / Front / Flat pattern views + overall dimensions + QTY note + tagged bend table."""
import gc
import math
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
BEND_TABLE_ANCHOR_TOP_LEFT = 1         # swBOMConfigurationAnchor_TopLeft
FALLBACK_SHEET = (0.210, 0.297)        # A4 portrait, meters (used only if the sheet size cannot be read)
SCALES = (10, 5, 4, 2, 1) + tuple(1.0 / n for n in (
    2, 3, 4, 5, 8, 10, 15, 20, 25, 30, 40, 50, 75, 100, 150, 200, 250, 300, 400, 500))
INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

DEFAULT_LAYOUT = {
    "margin_mm": 10,               # sheet border inset
    "side_pad_mm": 0,              # free space between the border and the views (left / right / top)
    "top_pad_mm": 5,
    "table_gap_mm": 0,             # gap between the title block and the bend table (0 = attached)
    "table_clearance_mm": 9,       # free space between the bend table and the flat pattern area
    "folded_band_ratio": 0.38,     # top views take the upper 38% of the free area, flat pattern the rest
    "band_gap_mm": 2,
    "view_gap_mm": 12,             # gap between Top and Front view
    "dim_pad_mm": 12,              # room kept free around each view for its dimension lines
    "common_scale": False,         # True = all views share one scale (like the reference), False = per group
    "bend_dim_pad_mm": 18,         # free space below the flat pattern for the bend dimensions
    "auto_rotate_flat": True,      # rotate a tall flat pattern by 90 deg so its long side is horizontal
    "scale_mode": "fine",          # "fine" = 1:n for any whole n (largest view that fits), "standard" = fixed list
}


class SwDrawingExporter:
    """Per part: copy master.slddrw -> Edit Sheet Format -> remove old views -> Top / Front / Flat pattern
    -> dimensions -> QTY note -> bend table with letter tags -> save .SLDDRW."""

    def __init__(self, sw_app, assembly_path=None, reconnect=None, template_path=None,
                 sheet_format_path=None, text_replacements=None, master_drawing=None,
                 layout=None, add_folded_views=True,
                 insert_bend_table=True, bend_table_template=None,
                 bend_table_right_margin_mm=10, bend_table_bottom_mm=66,
                 bend_table_remove_words=(), bend_table_start_value="A",
                 add_dimensions=True, dimension_offset_mm=15,
                 add_qty_note=True, qty_note_x_mm=None, qty_note_y_mm=None,
                 add_bend_dimensions=True, bend_dim_offset_mm=8, bend_dim_step_mm=7,
                 bend_dim_text_width_mm=14, bend_dim_order="largest_first"):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path).resolve() if assembly_path else None
        self.reconnect = reconnect
        self.template_path = Path(template_path) if template_path else None
        self.master_drawing = Path(master_drawing) if master_drawing else None
        self.sheet_format_path = Path(sheet_format_path) if sheet_format_path else None
        self.text_replacements = dict(text_replacements or {})
        self.layout = dict(DEFAULT_LAYOUT, **(layout or {}))
        self.add_folded_views = add_folded_views
        self.insert_bend_table = insert_bend_table
        self.bend_table_template = Path(bend_table_template) if bend_table_template else None
        self.bend_table_right_margin_mm = bend_table_right_margin_mm
        self.bend_table_bottom_mm = bend_table_bottom_mm
        self.bend_table_remove_words = tuple(w.lower() for w in bend_table_remove_words)
        self.bend_table_start_value = str(bend_table_start_value)
        self.add_dimensions = add_dimensions
        self.dimension_offset_mm = dimension_offset_mm
        self.add_qty_note = add_qty_note
        self.qty_note_x_mm = qty_note_x_mm
        self.qty_note_y_mm = qty_note_y_mm
        self.add_bend_dimensions = add_bend_dimensions
        self.bend_dim_offset_mm = bend_dim_offset_mm
        self.bend_dim_step_mm = bend_dim_step_mm
        self.bend_dim_text_width_mm = bend_dim_text_width_mm
        self.bend_dim_order = bend_dim_order

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
        """Reads a property or calls a no-argument method (pywin32 exposes both styles)."""
        if obj is None:
            return default
        try:
            value = getattr(obj, name)
            if hasattr(value, "_oleobj_"):      # a COM object (view, sheet, annotation...) must NOT be called
                return value
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

    def _open_drawing(self, path):
        errors, warnings = self._byref_i4(), self._byref_i4()
        return self.sw_app.OpenDoc6(str(path), SW_DOC_DRAWING, SW_OPEN_SILENT, "", errors, warnings)

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

    def _describe_flat(self, flat):
        return (f"flat feature: {self._get(flat, 'Name')} ({self._get(flat, 'GetTypeName2')}), "
                f"suppressed={self._get(flat, 'IsSuppressed')}")

    def _flat_error_code(self, flat):
        try:
            warning = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, False)
            return f"flat feature error code: {flat.GetErrorCode2(warning)} (warning={warning.value})"
        except Exception as exc:
            return f"flat feature error code unavailable ({exc})"

    def _sheet_size(self, drawing):
        sheet = self._get(drawing, "GetCurrentSheet")
        try:                                              # ISheet.GetSize(ref w, ref h)
            w = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
            h = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
            sheet.GetSize(w, h)
            if w.value and h.value and w.value > 0 and h.value > 0:
                return float(w.value), float(h.value)
        except Exception:
            pass
        try:
            props = self._get(sheet, "GetProperties2")
            if props and len(props) >= 7 and props[5] > 0 and props[6] > 0:
                return float(props[5]), float(props[6])
        except Exception:
            pass
        return FALLBACK_SHEET

    # ---------------- sheet format ----------------
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
        values = {
            "PartName": item.get("name") or Path(item["path"]).stem,
            "Quantity": item.get("quantity", 1),
            "Thickness": item.get("thickness"),
            "Material": item.get("material"),
            "Configuration": item.get("configuration"),
        }

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

        try:
            manager = drawing.Extension.CustomPropertyManager("")
            for key, value in values.items():
                if value not in (None, ""):
                    manager.Add3(key, 30, str(value), 2)     # 30 = text, 2 = overwrite
            steps.append("drawing properties written")
        except Exception as exc:
            steps.append(f"drawing properties skipped: {exc}")

        try:
            drawing.EditTemplate()          # return value is unreliable (None) -> ignored
            replaced = self._replace_sheet_texts(drawing, values)
            steps.append(f"Edit Sheet Format: {replaced} text(s) replaced")
        except Exception as exc:
            steps.append(f"Edit Sheet Format failed: {exc}")
        finally:
            try:
                drawing.EditSheet()         # always leave Edit Sheet Format
                steps.append("back to sheet (EditSheet)")
            except Exception as exc:
                steps.append(f"EditSheet failed: {exc}")

    # ---------------- master cleanup ----------------
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

        tables_removed = 0
        for v in views:
            tables = self._get(v, "GetTableAnnotations")
            for table in (tables or ()):
                try:
                    annotation = self._get(table, "GetAnnotation")
                    drawing.ClearSelection2(True)
                    if annotation is not None and annotation.Select3(False, nothing):
                        drawing.Extension.DeleteSelection2(6)
                        tables_removed += 1
                except Exception as exc:
                    steps.append(f"table delete failed: {exc}")
        steps.append(f"tables removed: {tables_removed}")

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

    # ---------------- views ----------------
    def _create_flat_view(self, drawing, part_path, config, x, y, steps, rounds=3):
        attempts = (
            ("CreateFlatPatternViewFromModelView3",
             lambda cfg: drawing.CreateFlatPatternViewFromModelView3(part_path, cfg, x, y, 0.0, False, False)),
            ("CreateDrawViewFromModelView3(*Flat pattern)",
             lambda cfg: drawing.CreateDrawViewFromModelView3(part_path, "*Flat pattern", x, y, 0.0)),
        )
        for round_no in range(1, rounds + 1):
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

    def _create_model_view(self, drawing, part_path, view_name, x, y, steps):
        try:
            view = drawing.CreateDrawViewFromModelView3(part_path, view_name, x, y, 0.0)
            if view is not None:
                return view
            steps.append(f"{view_name} view returned None")
        except Exception as exc:
            steps.append(f"{view_name} view failed: {exc}")
        return None

    def _move_view(self, view, target_cx, target_cy):
        for _ in range(2):                      # second pass removes small drift
            outline = self._get(view, "GetOutline")
            cx = (outline[0] + outline[2]) / 2.0
            cy = (outline[1] + outline[3]) / 2.0
            pos = self._get(view, "Position")
            view.Position = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                    (pos[0] + target_cx - cx, pos[1] + target_cy - cy))

    def _fit_scale(self, limit):
        """Largest scale that still fits (limit = max scale factor allowed by the free area)."""
        if self.layout.get("scale_mode", "fine") == "fine":
            if limit >= 1:
                return next(v for v in (10, 5, 4, 2, 1) if v <= limit * 1.001)
            return 1.0 / max(1, math.ceil(1.0 / limit - 1e-9))        # 1:n, n = smallest whole number that fits
        return next((v for v in SCALES if v <= limit * 1.001), SCALES[-1])

    @staticmethod
    def _scale_text(scale):
        return f"1:{1 / scale:.0f}" if scale < 1 else f"{scale:g}:1"

    def _size_at_unit_scale(self, view):
        """Sets a known scale first and measures at it, so a wrong ScaleDecimal reading cannot distort the fit."""
        probe = 0.05
        view.ScaleDecimal = probe
        o = self._get(view, "GetOutline")
        return (o[2] - o[0]) / probe, (o[3] - o[1]) / probe

    def _settle_scale(self, view, scale, box_w, box_h):
        """Applies the scale and checks the REAL outline; steps down until the view fits the box."""
        w = h = 0.0
        for _ in range(8):
            view.ScaleDecimal = scale
            o = self._get(view, "GetOutline")
            w, h = o[2] - o[0], o[3] - o[1]
            if w <= box_w * 1.002 and h <= box_h * 1.002:
                break
            scale = self._fit_scale(scale * 0.93)
        return scale, (w, h)

    def _layout_views(self, drawing, folded, flat_view, table_top, steps):
        """Same page logic as the PDF drawing: the bend table sits above the title block, the free area above
        it is split into a top band (folded views) and the flat pattern area; scales are fitted to the areas."""
        width, height = self._sheet_size(drawing)
        lay = self.layout
        mm = 0.001
        m, pad, dim = lay["margin_mm"] * mm, lay["side_pad_mm"] * mm, lay["dim_pad_mm"] * mm
        vgap = lay["view_gap_mm"] * mm
        left, right = m + pad, width - m - pad
        low = table_top + lay["table_clearance_mm"] * mm
        high = height - m - lay["top_pad_mm"] * mm
        try:
            if folded:
                band = (high - low) * lay["folded_band_ratio"]
                folded_box = (left, high - band, right, high)
                flat_box = (left, low, right, high - band - lay["band_gap_mm"] * mm)
            else:
                folded_box, flat_box = None, (left, low, right, high)

            # keep room only on the sides where dimension lines are drawn:
            #   Top/Front: above + right      Flat pattern: above + left
            f_in = (folded_box[0], folded_box[1], folded_box[2] - dim, folded_box[3] - dim) if folded_box else None
            bend_pad = lay.get("bend_dim_pad_mm", 18) * mm if self.add_bend_dimensions else 0.0
            b_in = (flat_box[0] + dim, flat_box[1] + bend_pad, flat_box[2], flat_box[3] - dim)
            steps.append(f"sheet {width * 1000:.0f} x {height * 1000:.0f} mm | table top {table_top * 1000:.0f} mm | "
                         f"flat area {flat_box[1] * 1000:.0f}-{flat_box[3] * 1000:.0f} mm"
                         + (f" | folded area {folded_box[1] * 1000:.0f}-{folded_box[3] * 1000:.0f} mm" if folded_box else ""))

            if lay.get("auto_rotate_flat", True):
                o = self._get(flat_view, "GetOutline")
                if (o[3] - o[1]) > (o[2] - o[0]) * 1.05:
                    try:
                        flat_view.Angle = math.pi / 2.0
                        steps.append("flat pattern rotated 90 deg (long side horizontal)")
                    except Exception as exc:
                        steps.append(f"flat pattern rotation failed: {exc}")

            limits = {}
            if folded:
                sizes = [self._size_at_unit_scale(v) for _, v in folded]
                limits["folded"] = min((f_in[2] - f_in[0]) / max(w for w, _ in sizes),
                                       (f_in[3] - f_in[1] - vgap * (len(folded) - 1)) / sum(h for _, h in sizes))
            fw, fh = self._size_at_unit_scale(flat_view)
            limits["flat"] = min((b_in[2] - b_in[0]) / fw, (b_in[3] - b_in[1]) / fh)

            if lay["common_scale"]:
                common = self._fit_scale(min(limits.values()))
                scales = {k: common for k in limits}
            else:
                scales = {k: self._fit_scale(v) for k, v in limits.items()}

            b_w, b_h = b_in[2] - b_in[0], b_in[3] - b_in[1]
            scales["flat"], (real_w, real_h) = self._settle_scale(flat_view, scales["flat"], b_w, b_h)
            if lay["common_scale"] and "folded" in scales:
                scales["folded"] = min(scales["folded"], scales["flat"])
                scales["flat"] = scales["folded"]
                flat_view.ScaleDecimal = scales["flat"]
            for _, v in folded:
                v.ScaleDecimal = scales["folded"]
            o = self._get(flat_view, "GetOutline")
            steps.append(f"flat pattern at {self._scale_text(scales['flat'])}: "
                         f"{(o[2] - o[0]) * 1000:.0f} x {(o[3] - o[1]) * 1000:.0f} mm "
                         f"in a box of {b_w * 1000:.0f} x {b_h * 1000:.0f} mm")

            if folded:
                heights = []
                for _, v in folded:
                    o = self._get(v, "GetOutline")
                    heights.append(o[3] - o[1])
                total = sum(heights) + vgap * (len(folded) - 1)
                cursor = (f_in[1] + f_in[3]) / 2.0 + total / 2.0
                for (_, v), h in zip(folded, heights):
                    self._move_view(v, (f_in[0] + f_in[2]) / 2.0, cursor - h / 2.0)
                    cursor -= h + vgap
            self._move_view(flat_view, (b_in[0] + b_in[2]) / 2.0, (b_in[1] + b_in[3]) / 2.0)

            txt = ", ".join(f"{k} {self._scale_text(v)}" for k, v in scales.items())
            steps.append(f"layout scales: {txt}")
            return scales["flat"]
        except Exception as exc:
            steps.append(f"layout failed: {exc}")
            return None

    def _set_sheet_scale(self, drawing, scale, steps):
        """Title block 'SCALE' field = scale of the flat pattern (as in the PDF version)."""
        text = self._scale_text(scale)
        num, den = (1.0, float(round(1 / scale))) if scale < 1 else (float(round(scale)), 1.0)
        try:
            sheet = self._get(drawing, "GetCurrentSheet")
            sheet.SetScale(num, den, False, False)      # False, False = do not rescale the views
            steps.append(f"sheet scale set to {text}")
        except Exception as exc:
            steps.append(f"sheet scale not set ({exc})")
        try:
            drawing.Extension.CustomPropertyManager("").Add3("Scale", 30, text, 2)
        except Exception:
            pass

    # ---------------- overall dimensions ----------------
    def _view_vertices(self, view):
        """Visible vertices of a view (the entity type is detected, so no enum value is assumed)."""
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        for entity_type in (2, 1, 3, 4):
            try:
                entities = view.GetVisibleEntities2(nothing, entity_type)
            except Exception:
                continue
            if entities and self._get(entities[0], "GetPoint") is not None:
                return list(entities)
        return []

    def _vertex_points(self, view, vertices, steps):
        """[(x, y, vertex)] in sheet coordinates, used only to find the extreme corners."""
        try:
            to_sheet = self._view_mapper(view, steps)
        except Exception as exc:
            steps.append(f"point transform unavailable ({exc}) - using model coordinates")
            to_sheet = None
        points = []
        for vertex in vertices:
            p = self._get(vertex, "GetPoint")
            if not p:
                continue
            xyz = (float(p[0]), float(p[1]), float(p[2]))
            try:
                x, y = to_sheet(xyz) if to_sheet else (xyz[0], xyz[1])
            except Exception:
                x, y = xyz[0], xyz[1]
            points.append((x, y, vertex))
        return points

    def _select_vertex(self, view, vertex, append, sel):
        """Vertex.Select4 first; IView.SelectEntity as fallback."""
        try:
            if vertex.Select4(append, sel):
                return True, "Select4"
        except Exception:
            pass
        try:
            if view.SelectEntity(vertex, append):
                return True, "SelectEntity"
        except Exception:
            pass
        return False, None

    def _add_overall_dimensions(self, drawing, view, steps, label, length=True, width=True, width_side="right"):
        """Length (horizontal, above the view) and/or width (vertical, left or right of the view)."""
        stage = "vertices"
        try:
            points = self._vertex_points(view, self._view_vertices(view), steps)
            if len(points) < 2:
                steps.append(f"dimensions {label}: no vertices found")
                return
            left, right = min(points, key=lambda t: t[0]), max(points, key=lambda t: t[0])
            bottom, top = min(points, key=lambda t: t[1]), max(points, key=lambda t: t[1])

            outline = self._get(view, "GetOutline")
            off = self.dimension_offset_mm / 1000.0
            try:
                sel = drawing.SelectionManager.CreateSelectData()
                sel.View = view
            except Exception:
                sel = VARIANT(pythoncom.VT_DISPATCH, None)

            jobs = []
            if length:
                jobs.append(("length", left, right, "AddHorizontalDimension2",
                             (outline[0] + outline[2]) / 2.0, outline[3] + off))
            if width:
                px = outline[2] + off if width_side == "right" else outline[0] - off
                jobs.append(("width", bottom, top, "AddVerticalDimension2",
                             px, (outline[1] + outline[3]) / 2.0))

            made = []
            for name, a_pt, b_pt, func, px, py in jobs:
                stage = f"{name}: select"
                drawing.ClearSelection2(True)
                ok1, how1 = self._select_vertex(view, a_pt[2], False, sel)
                ok2, _ = self._select_vertex(view, b_pt[2], True, sel)
                if not (ok1 and ok2):
                    made.append(f"{name}: could not select corners")
                    continue
                stage = f"{name}: {func}"
                dim = getattr(drawing, func)(px, py, 0.0)
                made.append(f"{name}: {'ok' if dim is not None else 'failed'}")
            drawing.ClearSelection2(True)
            steps.append(f"dimensions {label} -> " + ", ".join(made))
        except Exception as exc:
            steps.append(f"dimensions {label} failed at '{stage}': {exc}")

    # ---------------- bend dimensions ----------------
    def _view_edges(self, view, steps=None):
        """Return visible edge entities, tolerating pywin32 SAFEARRAY/tuple variants.

        swViewEntityType_e: Edge=1, SilhouetteEdge=3. Some SolidWorks versions
        return a tuple/SAFEARRAY wrapper, so don't assume entities[0] is the edge.
        """
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        found = []
        diagnostics = []

        def as_items(value):
            if value is None:
                return []
            # Some COM wrappers return (SAFEARRAY, status) instead of the array.
            if isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], (bool, int)):
                value = value[0]
            try:
                return list(value)
            except Exception:
                return [value]

        def looks_like_edge(obj):
            if obj is None:
                return False
            try:
                start = self._get(obj, "GetStartVertex")
                end = self._get(obj, "GetEndVertex")
                if start is not None and end is not None:
                    return True
            except Exception:
                pass
            # Fallback for wrappers exposing only curve/start/end point methods.
            return (self._get(obj, "GetCurve") is not None and
                    (self._get(obj, "GetStartPoint2") is not None or
                     self._get(obj, "GetEndPoint2") is not None))

        # Edge and silhouette-edge collections are the only relevant types here.
        for entity_type in (1, 3):
            try:
                raw = view.GetVisibleEntities2(nothing, entity_type)
                items = as_items(raw)
                edges = [obj for obj in items if looks_like_edge(obj)]
                diagnostics.append(f"type {entity_type}: {len(items)} entities/{len(edges)} edges")
                for edge in edges:
                    if not any(edge is prior for prior in found):
                        found.append(edge)
            except Exception as exc:
                diagnostics.append(f"type {entity_type}: {exc}")
        if steps is not None:
            steps.append("visible edge scan: " + "; ".join(diagnostics))
        return found

    def _make_mapper(self, view, steps):
        """model space -> view space. Uses SolidWorks' MathUtility; if that fails, the transform matrix."""
        transform = self._get(view, "ModelToViewTransform")
        if transform is None:
            raise RuntimeError("ModelToViewTransform not available")
        math_util = self._get(self.sw_app, "GetMathUtility")        # an object, not a function call
        if math_util is not None:
            def com_map(xyz, via=None):
                mp = math_util.CreatePoint(VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                                   tuple(float(v) for v in xyz)))
                if via is not None:
                    mp = mp.MultiplyTransform(via)
                a = mp.MultiplyTransform(transform).ArrayData
                return float(a[0]), float(a[1])
            try:
                com_map((0.0, 0.0, 0.0))
                return com_map, "com"
            except Exception as exc:
                steps.append(f"MathUtility transform failed ({exc}) - using the transform matrix")

        data = self._get(transform, "ArrayData")
        if not data or len(data) < 13:
            raise RuntimeError("transform matrix not available")
        d = [float(v) for v in data]
        rows = [d[0:3], d[3:6], d[6:9]]
        state = {"transpose": False, "k": d[12] or 1.0}

        def apply_transform(xyz, xform, transpose=None, scale_override=None):
            """Apply a SolidWorks MathTransform matrix, including translation.

            The same ArrayData orientation selected from ModelToViewTransform is used
            for chained sketch transforms as well; mixing conventions here can move
            bend segments outside the drawing-view bounds.
            """
            if xform is None:
                return tuple(float(v) for v in xyz)
            arr = self._get(xform, "ArrayData")
            if arr is None or len(arr) < 13:
                raise RuntimeError("transform matrix data unavailable")
            a = [float(v) for v in arr]
            rr = [a[0:3], a[3:6], a[6:9]]
            use_transpose = state["transpose"] if transpose is None else transpose
            if use_transpose:
                rr = [[rr[j][i] for j in range(3)] for i in range(3)]
            x, y, z = (float(v) for v in xyz)
            sc = float(scale_override if scale_override is not None else (a[12] or 1.0))
            return (sc * (rr[0][0]*x + rr[0][1]*y + rr[0][2]*z) + a[9],
                    sc * (rr[1][0]*x + rr[1][1]*y + rr[1][2]*z) + a[10],
                    sc * (rr[2][0]*x + rr[2][1]*y + rr[2][2]*z) + a[11])

        def mat_map(xyz, via=None):
            # Bend-line endpoints may be in sketch space. Convert sketch -> model
            # first, then model -> view. Previously this path always raised and
            # caused every bend line to be discarded when MathUtility was unavailable.
            model_xyz = apply_transform(xyz, via) if via is not None else tuple(float(v) for v in xyz)
            x, y, z = model_xyz
            r = rows if not state["transpose"] else [[rows[j][i] for j in range(3)] for i in range(3)]
            sc = state["k"]
            return (sc * (r[0][0]*x + r[0][1]*y + r[0][2]*z),
                    sc * (r[1][0]*x + r[1][1]*y + r[1][2]*z))
        mat_map.state = state
        return mat_map, "matrix"

    def _view_mapper(self, view, steps, verbose=False):
        """Function xyz -> (x, y) in SHEET coordinates, calibrated against the view's outline.
        .vb = bounding box of the view's vertices in sheet coordinates."""
        to_view, kind = self._make_mapper(view, steps)
        points = [p for p in (self._get(v, "GetPoint") for v in self._view_vertices(view)) if p]
        if len(points) < 2:
            raise RuntimeError("no vertices in the view")
        outline = self._get(view, "GetOutline")
        ow, oh = outline[2] - outline[0], outline[3] - outline[1]

        def bbox():
            xy = [to_view(p) for p in points]
            return (min(x for x, _ in xy), min(y for _, y in xy), max(x for x, _ in xy), max(y for _, y in xy))

        if kind == "matrix":                         # choose the matrix convention and the scale by the outline
            state, best = to_view.state, None
            for transpose in (False, True):
                state["transpose"], state["k"] = transpose, 1.0
                b = bbox()
                w, h = b[2] - b[0], b[3] - b[1]
                err = abs(w / h - ow / oh) if h > 1e-9 and oh > 1e-9 else abs(w - ow)
                if best is None or err < best[0]:
                    best = (err, transpose, w)
            state["transpose"] = best[1]
            state["k"] = ow / best[2] if best[2] > 1e-9 else 1.0
        vb = bbox()
        off = ((outline[0] + outline[2]) / 2.0 - (vb[0] + vb[2]) / 2.0,
               (outline[1] + outline[3]) / 2.0 - (vb[1] + vb[3]) / 2.0)
        if abs((vb[2] - vb[0]) - ow) > 0.05 * max(ow, 1e-9):
            steps.append("warning: view transform and outline sizes differ")
        if verbose:
            steps.append(f"point transform: {kind}")

        def to_sheet(xyz, via=None):
            x, y = to_view(xyz, via)
            return x + off[0], y + off[1]
        to_sheet.vb = (vb[0] + off[0], vb[1] + off[1], vb[2] + off[0], vb[3] + off[1])
        return to_sheet

    def _segment_xyz(self, segment):
        points = []
        for name in ("GetStartPoint2", "GetEndPoint2"):
            pt = self._get(segment, name)
            points.append(tuple(float(self._get(pt, c)) for c in ("X", "Y", "Z")))
        return points

    def _add_bend_dimensions(self, drawing, view, scale, limit_bottom, steps):
        """One dimension per bend: from the bend line to the nearest outer edge of the flat pattern.
        Sorted by value (largest first by default) and stacked on separate levels so numbers never overlap."""
        stage = "bend lines"
        try:
            lines = list(self._get(view, "GetBendLines") or ())
            steps.append(f"bend lines in the flat pattern view: {len(lines)}")
            if not lines:
                return
            scale = float(scale or self._get(view, "ScaleDecimal"))

            # ---- geometry in sheet coordinates ----
            stage = "geometry"
            to_sheet = self._view_mapper(view, steps, verbose=True)
            vb = to_sheet.vb
            outline = self._get(view, "GetOutline")
            tol = 0.02 * max(vb[2] - vb[0], vb[3] - vb[1])

            stage = "edges"
            edges = []
            for e in self._view_edges(view, steps):
                curve = self._get(e, "GetCurve")
                is_line = self._get(curve, "IsLine") if curve is not None else None
                if is_line is False:
                    continue
                sv = self._get(e, "GetStartVertex")
                ev = self._get(e, "GetEndVertex")
                p = self._get(sv, "GetPoint") if sv is not None else None
                q = self._get(ev, "GetPoint") if ev is not None else None
                # Point2 endpoints are a fallback for edge wrappers that do not
                # expose IVertex objects consistently through pywin32.
                if p is None or q is None:
                    p = self._get(e, "GetStartPoint2")
                    q = self._get(e, "GetEndPoint2")
                if p is None or q is None:
                    continue
                try:
                    p = tuple(float(p[i]) for i in range(3))
                    q = tuple(float(q[i]) for i in range(3))
                    edges.append((to_sheet(p), to_sheet(q), e))
                except Exception:
                    continue
            steps.append(f"straight edges used: {len(edges)}")

            stage = "bend line positions"
            bends = []
            for i, seg in enumerate(lines):
                pts = self._segment_xyz(seg)
                candidates = [[to_sheet(pt) for pt in pts]]
                try:                                                    # sketch space -> model space
                    sketch = self._get(seg, "GetSketch")
                    model_to_sketch = self._get(sketch, "ModelToSketchTransform") if sketch is not None else None
                    inv = self._get(model_to_sketch, "Inverse") if model_to_sketch is not None else None
                    if inv is not None:
                        candidates.append([to_sheet(pt, via=inv) for pt in pts])
                except Exception as exc:
                    steps.append(f"bend line {i}: sketch transform fallback failed ({exc})")
                found = None
                for k, cand in enumerate(candidates):
                    c2 = cand
                    if all(vb[0] - tol <= x <= vb[2] + tol and vb[1] - tol <= y <= vb[3] + tol for x, y in c2):
                        found = c2
                        break
                if found is None:
                    # Keep a useful diagnostic: distinguishes a coordinate-space problem
                    # from a genuinely missing/invalid bend segment.
                    try:
                        raw_box = (min(p[0] for p in pts), min(p[1] for p in pts),
                                   max(p[0] for p in pts), max(p[1] for p in pts))
                        cand_boxes = []
                        for cand in candidates:
                            cand_boxes.append((round(min(p[0] for p in cand), 6),
                                               round(min(p[1] for p in cand), 6),
                                               round(max(p[0] for p in cand), 6),
                                               round(max(p[1] for p in cand), 6)))
                        steps.append(f"bend line {i}: outside view bounds; raw={tuple(round(v,6) for v in raw_box)}, candidates={cand_boxes}, view={tuple(round(v,6) for v in vb)}")
                    except Exception:
                        steps.append(f"bend line {i}: position does not match the view - skipped")
                    continue
                bends.append((seg, found))

            stage = "distances"
            items = []
            for seg, (A, B) in bends:
                dx, dy = B[0] - A[0], B[1] - A[1]
                length = math.hypot(dx, dy)
                if length < 1e-6:
                    continue
                ux, uy = dx / length, dy / length
                nx, ny = -uy, ux
                s0 = nx * A[0] + ny * A[1]
                pos = neg = None
                for p, q, e in edges:
                    le = math.hypot(q[0] - p[0], q[1] - p[1])
                    if le < 1e-6 or abs(ux * (q[1] - p[1]) - uy * (q[0] - p[0])) / le > 0.01:
                        continue                                        # not parallel to the bend line
                    delta = nx * p[0] + ny * p[1] - s0
                    if abs(delta) < 1e-4:
                        continue                                        # the bend line's own edge
                    if delta > 0 and (pos is None or delta > pos[0]):
                        pos = (delta, e, 1.0)                           # farthest edge on that side = outer edge
                    if delta < 0 and (neg is None or -delta > neg[0]):
                        neg = (-delta, e, -1.0)
                options = [o for o in (pos, neg) if o]
                if not options:
                    steps.append("a bend line has no parallel outer edge - skipped")
                    continue
                dist, edge, sign = min(options, key=lambda o: o[0])    # nearest outer edge
                tx, ty = A[0] + sign * dist * nx, A[1] + sign * dist * ny
                if abs(nx) >= abs(ny):
                    group, lo, hi = "h", min(A[0], tx), max(A[0], tx)
                else:
                    group, lo, hi = "v", min(A[1], ty), max(A[1], ty)
                items.append({"seg": seg, "edge": edge, "group": group, "lo": lo, "hi": hi,
                              "mid": (lo + hi) / 2.0, "value": dist / scale * 1000.0})

            # ---- order + levels (no overlapping numbers) ----
            stage = "levels"
            # The supplied FlatPatternView logic places the smallest dimension nearest the
            # geometry and progressively larger dimensions farther out, reducing visual clutter.
            # "largest_first" retains the requested value priority for the outer rows.
            items.sort(key=lambda it: it["value"], reverse=(self.bend_dim_order == "smallest_first"))
            tw = self.bend_dim_text_width_mm / 1000.0
            gap = 0.001
            for group in ("h", "v"):
                placed = []
                for it in (i for i in items if i["group"] == group):
                    lo, hi = min(it["lo"], it["mid"] - tw / 2.0), max(it["hi"], it["mid"] + tw / 2.0)
                    level = 0
                    while any(l == level and lo < h2 + gap and hi > l2 - gap for l, l2, h2 in placed):
                        level += 1
                    it["level"] = level
                    placed.append((level, lo, hi))

            # ---- create the dimensions ----
            stage = "dimensions"
            start = self.bend_dim_offset_mm / 1000.0
            step = self.bend_dim_step_mm / 1000.0
            width, _ = self._sheet_size(drawing)
            try:
                sel = drawing.SelectionManager.CreateSelectData()
                sel.View = view
            except Exception:
                sel = VARIANT(pythoncom.VT_DISPATCH, None)

            created, too_low = 0, False
            for it in items:
                if it["group"] == "h":                       # measured along x -> stacked BELOW the view
                    px, py = it["mid"], outline[1] - start - it["level"] * step
                    too_low = too_low or py < limit_bottom
                else:                                        # measured along y -> stacked at the RIGHT
                    px, py = outline[2] + start + it["level"] * step, it["mid"]
                    if px > width - self.layout["margin_mm"] / 1000.0:       # no room: use the left side
                        px = outline[0] - start - 0.010 - it["level"] * step
                drawing.ClearSelection2(True)
                ok1, _ = self._select_vertex(view, it["seg"], False, sel)
                ok2, _ = self._select_vertex(view, it["edge"], True, sel)
                if ok1 and ok2 and drawing.AddDimension2(px, py, 0.0) is not None:
                    created += 1
            drawing.ClearSelection2(True)
            levels = {g: 1 + max([i["level"] for i in items if i["group"] == g] or [-1]) for g in ("h", "v")}
            steps.append(f"bend dimensions: {created}/{len(items)} created, values "
                         f"{[round(i['value'], 2) for i in items]}, levels {levels}")
            if too_low:
                steps.append("warning: bend dimensions reach the bend table area - raise bend_dim_pad_mm")
        except Exception as exc:
            steps.append(f"bend dimensions failed at '{stage}': {exc}")

    # ---------------- QTY note ----------------
    def _add_qty_note(self, drawing, item, steps, table_bottom=None, row_h=None):
        qty = item.get("quantity", 1)
        try:
            x = (self.qty_note_x_mm if self.qty_note_x_mm is not None else self.layout["margin_mm"] + 12) / 1000.0
            if self.qty_note_y_mm is not None:
                y = self.qty_note_y_mm / 1000.0
            else:
                base = table_bottom if table_bottom is not None else self.bend_table_bottom_mm / 1000.0
                y = base + 1.2 * (row_h or 0.007)
            note = drawing.InsertNote(f"QTY={qty}")
            annotation = self._get(note, "GetAnnotation")
            annotation.SetPosition2(x, y, 0.0)
            steps.append(f"QTY note added: QTY={qty}")
        except Exception as exc:
            steps.append(f"QTY note failed: {exc}")

    # ---------------- bend table ----------------
    def _default_bend_templates(self):
        found = []
        try:
            exe = Path(str(self._get(self.sw_app, "GetExecutablePath")))
            for base in (exe, exe.parent):
                found += [str(p) for p in base.glob("lang/*/*.sldbndtbt")]
        except Exception:
            pass
        return found

    @staticmethod
    def _cell(grid, row, col):
        try:
            return str(grid.Text(row, col)).strip()
        except Exception:
            return ""

    def _insert_table_once(self, view, x, y, steps):
        templates = [str(self.bend_table_template)] if self.bend_table_template else []
        templates += self._default_bend_templates() + [""]
        for tpl in templates:
            name = Path(tpl).name if tpl else "default"
            try:
                table = view.InsertBendTable(False, x, y, BEND_TABLE_ANCHOR_TOP_LEFT,
                                             self.bend_table_start_value, tpl)
                if table is not None:
                    steps.append(f"bend table inserted (template: {name}, start tag: {self.bend_table_start_value})")
                    return table
                steps.append(f"bend table returned None ({name})")
            except Exception as exc:
                steps.append(f"bend table failed ({name}): {exc}")
        return None

    def _table_grid(self, table, steps):
        try:
            table.ColumnCount
            return table
        except Exception:
            for name in ("GetTableAnnotation", "TableAnnotation"):
                inner = self._get(table, name)
                if inner is not None:
                    steps.append(f"table accessed via {name}")
                    return inner
        return table

    def _clean_columns(self, grid, steps):
        try:
            cols = int(self._get(grid, "ColumnCount"))
            headers = [self._cell(grid, 0, c) for c in range(cols)]
            if len(set(headers)) == 1 and cols > 1:           # row 0 is a merged title row
                headers = [self._cell(grid, 1, c) for c in range(cols)]
            steps.append(f"bend table columns: {headers}")
            removed = []
            for c in range(cols - 1, -1, -1):
                if any(w in headers[c].lower() for w in self.bend_table_remove_words):
                    try:
                        grid.DeleteColumn(c)
                    except Exception:
                        grid.DeleteColumn2(c, False)
                    removed.append(headers[c])
            steps.append(f"bend table columns removed: {removed}")
            steps.append(f"bend table rows: {self._get(grid, 'RowCount')}")
        except Exception as exc:
            steps.append(f"bend table column cleanup failed: {exc}")

    def _table_size(self, grid, steps):
        cols = int(self._get(grid, "ColumnCount") or 0)
        rows = int(self._get(grid, "RowCount") or 0)
        try:
            return (sum(float(grid.GetColumnWidth(c)) for c in range(cols)),
                    sum(float(grid.GetRowHeight(r)) for r in range(rows)))
        except Exception as exc:
            steps.append(f"bend table size estimated ({exc})")
            return 0.070, 0.0072 * max(rows, 1)

    def _annotation_of(self, grid, view=None):
        annotation = self._get(grid, "GetAnnotation")
        if annotation is not None:
            return annotation
        if view is not None:                                  # fallback: the tables of the view
            tables = self._get(view, "GetTableAnnotations")
            for table in reversed(list(tables or ())):
                annotation = self._get(table, "GetAnnotation")
                if annotation is not None:
                    return annotation
        return None

    def _move_table(self, grid, left, top, steps, view=None):
        """SetPosition2 = top-left corner. Returns True when the table really is at the target."""
        try:
            annotation = self._annotation_of(grid, view)
            if annotation is None:
                steps.append("bend table: annotation object not found")
                return False
            before = self._get(annotation, "GetPosition")
            annotation.SetPosition2(left, top, 0.0)
            after = self._get(annotation, "GetPosition")
            steps.append(
                f"bend table position {tuple(round(v * 1000, 1) for v in (before or ()))} -> "
                f"{tuple(round(v * 1000, 1) for v in (after or ()))} mm (target {left * 1000:.1f}, {top * 1000:.1f})")
            return bool(after) and abs(after[0] - left) < 0.0005 and abs(after[1] - top) < 0.0005
        except Exception as exc:
            steps.append(f"bend table not moved: {exc}")
            return False

    def _delete_table(self, drawing, grid, view=None):
        try:
            nothing = VARIANT(pythoncom.VT_DISPATCH, None)
            annotation = self._annotation_of(grid, view)
            drawing.ClearSelection2(True)
            if annotation is not None and annotation.Select3(False, nothing):
                drawing.Extension.DeleteSelection2(6)
                return True
        except Exception:
            pass
        return False

    def _insert_bend_table(self, drawing, view, steps):
        """Native bend table: every bend gets a letter tag on the flat pattern (A, B, C...)
        and the same tag is listed in the table (Tag / Direction / Angle / Inner Radius).
        The table ends up in the bottom-right corner, right above the title block."""
        width, height = self._sheet_size(drawing)

        table = self._insert_table_once(view, width * 0.1, height * 0.5, steps)   # inside the sheet first
        if table is None:
            steps.append("bend table NOT created")
            return None
        try:
            # with a bend table present the bend notes of the flat pattern show the TAG (A, B, C...)
            view.ShowSheetMetalBendNotes = True
            steps.append(f"bend notes (tags) on view: {self._get(view, 'ShowSheetMetalBendNotes')}")
        except Exception as exc:
            steps.append(f"could not show bend tags: {exc}")

        grid = self._table_grid(table, steps)
        self._clean_columns(grid, steps)

        table_w, table_h = self._table_size(grid, steps)
        left = width - self.bend_table_right_margin_mm / 1000.0 - table_w
        bottom = (self.bend_table_bottom_mm + self.layout["table_gap_mm"]) / 1000.0
        top = bottom + table_h
        rows = max(int(self._get(grid, "RowCount") or 1), 1)
        steps.append(f"bend table size {table_w * 1000:.0f} x {table_h * 1000:.0f} mm")

        if not self._move_table(grid, left, top, steps, view):
            # moving did not work: delete it and insert again directly at the final position
            if self._delete_table(drawing, grid, view):
                steps.append("bend table re-inserted at the final position")
                table = self._insert_table_once(view, left, top, steps)
                if table is not None:
                    grid = self._table_grid(table, steps)
                    self._clean_columns(grid, steps)
                    self._move_table(grid, left, top, steps, view)
            else:
                steps.append("bend table could not be moved or deleted - check its position manually")
        return bottom, top, table_h / rows

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
                Path(slddrw_path).unlink(missing_ok=True)   # do not leave a master copy without views
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
        unsuppressed_by_us = False
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
            steps.append(self._describe_flat(flat))

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

            # ---- views ----
            flat_view, method = self._create_flat_view(drawing, model_path, active_cfg,
                                                       width / 2, height / 3, steps, rounds=1)
            if flat_view is None and bool(self._get(flat, "IsSuppressed", False)):
                try:                                              # second chance: unsuppress the flat pattern
                    flat.SetSuppression2(SW_UNSUPPRESS, SW_THIS_CONFIGURATION, None)
                    unsuppressed_by_us = True
                    model.ForceRebuild3(True)
                    steps.append("flat pattern unsuppressed (retry)")
                except Exception as exc:
                    steps.append(f"could not unsuppress flat pattern: {exc}")
                flat_view, method = self._create_flat_view(drawing, model_path, active_cfg,
                                                           width / 2, height / 3, steps, rounds=2)
            if flat_view is None:
                steps.append(self._describe_flat(flat))
                steps.append(self._flat_error_code(flat))
                return False, "could not create the flat pattern view", steps
            steps.append(f"flat view created ({method})")

            views = []
            if self.add_folded_views:
                top_view = self._create_model_view(drawing, model_path, "*Top", width / 2, height * 0.75, steps)
                front_view = self._create_model_view(drawing, model_path, "*Front", width / 2, height * 0.55, steps)
                if top_view is not None:
                    views.append(("top", top_view))
                if front_view is not None:
                    views.append(("front", front_view))
                steps.append(f"folded views: {[n for n, _ in views]}")
            views.append(("flat", flat_view))

            # ---- bend table first: its top edge defines the free area for the views ----
            table_info = None
            table_top = (self.bend_table_bottom_mm + self.layout["table_gap_mm"]) / 1000.0
            if self.insert_bend_table:
                table_info = self._insert_bend_table(drawing, flat_view, steps)
                if table_info:
                    table_top = table_info[1]

            # ---- views fitted into the A4 areas ----
            folded = [(n, v) for n, v in views if n != "flat"]
            flat_scale = self._layout_views(drawing, folded, flat_view, table_top, steps)
            if flat_scale:
                self._set_sheet_scale(drawing, flat_scale, steps)

            # ---- dimensions ----
            if self.add_dimensions:
                for name, v in views:
                    if name == "top":
                        self._add_overall_dimensions(drawing, v, steps, "top", True, True, "right")
                    elif name == "front":
                        # Front view also receives both overall dimensions (horizontal + vertical).
                        self._add_overall_dimensions(drawing, v, steps, "front", True, True, "right")
                    else:
                        self._add_overall_dimensions(drawing, v, steps, "flat", True, True, "left")

            if self.add_bend_dimensions:
                self._add_bend_dimensions(drawing, flat_view, flat_scale, table_top, steps)

            if self.add_qty_note:
                self._add_qty_note(drawing, item, steps,
                                   table_bottom=table_info[0] if table_info else None,
                                   row_h=table_info[2] if table_info else None)

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
            if flat is not None and unsuppressed_by_us:
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