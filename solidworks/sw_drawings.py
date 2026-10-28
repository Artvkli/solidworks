"""One SLDDRW (+ PDF) per sheet-metal part, saved in one folder per sheet thickness:
master.slddrw + folded views (Top / Front / Right) + Flat pattern view + overall dimensions
+ bend-line-to-outer-edge dimensions (one per bend tag) + QTY note + tagged bend table.

Bend dimensions (see bend_geometry.py for the coordinate systems):
  values are decided in the true-size frame of the bend-line sketch (sketch -> model via
  ModelToSketchTransform), the drawing-view transform is only used to place text and pair tags,
  and every created dimension is read back from SolidWorks and compared with the expected value.

Layout: every part is scaled by its own size (small parts get a larger scale), nothing may overlap:
  views + dimension room are checked against each other, and a final readability pass moves tags,
  notes and bend dimension texts on the flat pattern until no text touches another one.
"""
import gc
import itertools
import json
import math
import re
import shutil
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

from .bend_geometry import (Xform, assign_tags, classify_outer, detect_convention, measure_bends,
                            order_key, plan_placements, point_segment_distance, same_segment)

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
    "folded_band_ratio": 0.38,     # old fixed split; now only the reference for sharing the free space
    "band_gap_mm": 2,
    "view_gap_mm": 12,             # minimum gap between views (grows to fit the dimensions between them)
    "dim_pad_mm": 12,              # minimum room kept free around each view for its dimension lines
    "common_scale": False,         # True = all views share one scale, False = per group
    "bend_dim_pad_mm": 18,         # initial free space below the flat pattern for the bend dimensions
    "auto_rotate_flat": True,      # rotate a tall flat pattern by 90 deg so its long side is horizontal
    "scale_mode": "fine",          # "fine" = 1:n for any whole n (largest view that fits), "standard" = fixed list
    "max_scale": 2.0,              # small parts are enlarged up to 2:1
    "folded_ratio": 0.6,           # longest side of Top/Front/Right drawn at least 0.6x the flat pattern's longest side
    "folded_arrange": "auto",      # "auto" = per part the arrangement that draws the views largest, "standard" = fixed grid
    "folded_rotate": True,         # "auto" may turn a folded view by 90 deg (e.g. a tall Top view)
}

# Grid cell (column, row) of each folded view inside the folded-view band; row 0 is the top row.
# Top sits above Front, Right sits to the right of Front.
FOLDED_GRID = {"top": (0, 0), "front": (0, 1), "right": (1, 1), "left": (-1, 1), "bottom": (0, 2), "back": (2, 1)}
SIDES = ("left", "right", "top", "bottom")
DIM_VALUE_TOL = 1e-5               # 0.01 mm: created dimension must match the computed value within this


class SwDrawingExporter:
    """Per part: copy master.slddrw -> Edit Sheet Format -> remove old views -> Top / Front / Right / Flat pattern
    -> dimensions -> QTY note -> bend table with letter tags -> save .SLDDRW + .PDF in the thickness folder."""

    def __init__(self, sw_app, assembly_path=None, reconnect=None, template_path=None,
                 sheet_format_path=None, text_replacements=None, master_drawing=None,
                 layout=None, add_folded_views=True,
                 insert_bend_table=True, bend_table_template=None,
                 bend_table_right_margin_mm=10, bend_table_bottom_mm=66,
                 bend_table_remove_words=(), bend_table_start_value="A", bend_table_scale=1 / 1.5,
                 add_dimensions=True, dimension_offset_mm=15,
                 add_qty_note=True, qty_note_x_mm=None, qty_note_y_mm=None,
                 add_bend_dimensions=True, bend_dim_offset_mm=8, bend_dim_step_mm=7,
                 bend_dim_text_width_mm=14, bend_dim_order="largest_outside",
                 folded_view_names=("*Top", "*Front"), bend_dim_right_pad_mm=18,
                 bend_dim_text_height_mm=3.5, bend_dim_max_levels=12, bend_dim_merge_collinear=True,
                 bend_dim_verify="flag", bend_tag_max_distance_mm=15, bend_debug_json=False,
                 group_by_thickness=True, export_pdf=True,
                 thickness_folder_format="{thickness:g} mm", unknown_thickness_folder="no thickness"):
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
        self.bend_table_scale = float(bend_table_scale or 1.0)
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
        self.folded_view_names = tuple(folded_view_names or ())
        self.bend_dim_right_pad_mm = bend_dim_right_pad_mm
        self.bend_dim_text_height_mm = bend_dim_text_height_mm
        self.bend_dim_max_levels = int(bend_dim_max_levels)
        self.bend_dim_merge_collinear = bend_dim_merge_collinear
        self.bend_dim_verify = bend_dim_verify
        self.bend_tag_max_distance_mm = bend_tag_max_distance_mm
        self.bend_debug_json = bend_debug_json
        self.group_by_thickness = group_by_thickness
        self.export_pdf = export_pdf
        self.thickness_folder_format = thickness_folder_format
        self.unknown_thickness_folder = unknown_thickness_folder
        self._convention = None                          # MathTransform convention, verified once per session
        self._convention_source = None
        self._convention_check_failed = False
        self._bend_tags = []
        self._current_item = {}

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
            if hasattr(value, "_oleobj_"):
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
        try:
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
        if not self.text_replacements:
            return 0
        count = 0
        view = self._get(drawing, "GetFirstView")
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
                    manager.Add3(key, 30, str(value), 2)
            steps.append("drawing properties written")
        except Exception as exc:
            steps.append(f"drawing properties skipped: {exc}")

        try:
            drawing.EditTemplate()
            replaced = self._replace_sheet_texts(drawing, values)
            steps.append(f"Edit Sheet Format: {replaced} text(s) replaced")
        except Exception as exc:
            steps.append(f"Edit Sheet Format failed: {exc}")
        finally:
            try:
                drawing.EditSheet()
                steps.append("back to sheet (EditSheet)")
            except Exception as exc:
                steps.append(f"EditSheet failed: {exc}")

    # ---------------- master cleanup ----------------
    def _remove_existing_views(self, drawing, steps):
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)

        views = []
        view = self._get(self._get(drawing, "GetFirstView"), "GetNextView")
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
        for _ in range(2):
            outline = self._get(view, "GetOutline")
            cx = (outline[0] + outline[2]) / 2.0
            cy = (outline[1] + outline[3]) / 2.0
            pos = self._get(view, "Position")
            view.Position = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                    (pos[0] + target_cx - cx, pos[1] + target_cy - cy))

    def _fit_scale(self, limit):
        if not limit or limit <= 0:
            return SCALES[-1]
        if self.layout.get("scale_mode", "fine") == "fine":
            if limit >= 1:
                return next(v for v in (10, 5, 4, 2, 1) if v <= limit * 1.001)
            return 1.0 / max(1, math.ceil(1.0 / limit - 1e-9))
        return next((v for v in SCALES if v <= limit * 1.001), SCALES[-1])

    @staticmethod
    def _scale_text(scale):
        return f"1:{1 / scale:.0f}" if scale < 1 else f"{scale:g}:1"

    def _unit_size_and_border(self, view):
        """(width, height) of the view geometry at scale 1, plus the outline border (x and y).
        GetOutline = scale * size + constant border, so two probe scales separate the two."""
        probes = []
        for probe in (0.05, 0.1):
            view.ScaleDecimal = probe
            o = self._get(view, "GetOutline")
            probes.append((probe, o[2] - o[0], o[3] - o[1]))
        (s1, w1, h1), (s2, w2, h2) = probes
        w, h = (w2 - w1) / (s2 - s1), (h2 - h1) / (s2 - s1)
        if w <= 0 or h <= 0:
            return w2 / s2, h2 / s2, 0.0, 0.0
        return w, h, max(0.0, w2 - s2 * w), max(0.0, h2 - s2 * h)

    def _size_at_unit_scale(self, view):
        w, h, _, _ = self._unit_size_and_border(view)
        return w, h

    def _settle_scale(self, view, scale, box_w, box_h):
        w = h = 0.0
        for _ in range(8):
            view.ScaleDecimal = scale
            o = self._get(view, "GetOutline")
            w, h = o[2] - o[0], o[3] - o[1]
            if w <= box_w * 1.002 and h <= box_h * 1.002:
                break
            scale = self._fit_scale(scale * 0.93)
        return scale, (w, h)

    def _initial_flat_pads(self):
        """Extra room (metres) around the flat pattern for bend dimensions: below and to the right."""
        if not self.add_bend_dimensions:
            return dict.fromkeys(SIDES, 0.0)
        return {"left": 0.0, "top": 0.0,
                "bottom": self.layout.get("bend_dim_pad_mm", 18) / 1000.0,
                "right": self.bend_dim_right_pad_mm / 1000.0}

    def _dim_room(self):
        """(room above a view for its horizontal dimension, room beside a view for its vertical dimension), metres.
        = dimension offset + text size + 3 mm, so the numbers never reach the next view."""
        mm = 0.001
        base = self.layout["dim_pad_mm"] * mm
        if not self.add_dimensions:
            return base, base
        off = self.dimension_offset_mm * mm
        th, tw = self.bend_dim_text_height_mm * mm, self.bend_dim_text_width_mm * mm
        return max(base, off + th + 3 * mm), max(base, off + tw / 2.0 + 3 * mm)

    @staticmethod
    def _folded_cells(folded):
        """(name, view, column, row) for each folded view, using FOLDED_GRID; unknown names go right of row 1."""
        cells, used = [], set()
        for name, view in folded:
            cell = FOLDED_GRID.get(name)
            if cell is None or cell in used:
                col = max([c for c, r in used if r == 1] + [0]) + 1
                cell = (col, 1)
            used.add(cell)
            cells.append((name, view, cell[0], cell[1]))
        return cells

    def _layout_problems(self, folded, flat_view, pads, table_top, width, height):
        """Every view + the room of its dimensions, checked against each other, the sheet border and the
        bend table / title block. Returns a list of problems (empty = nothing overlaps)."""
        mm = 0.001
        dim_h, dim_v = self._dim_room()
        m = self.layout["margin_mm"] * mm
        tol = 0.2 * mm
        boxes = []
        for name, v in folded:
            o = self._get(v, "GetOutline")
            if o:
                boxes.append((name, (o[0], o[1], o[2] + dim_v, o[3] + dim_h)))
        o = self._get(flat_view, "GetOutline")
        if o:
            boxes.append(("flat", (o[0] - dim_v - pads.get("left", 0.0), o[1] - pads.get("bottom", 0.0),
                                   o[2] + pads.get("right", 0.0), o[3] + dim_h + pads.get("top", 0.0))))
        problems = []
        for name, b in boxes:
            if b[0] < m - tol or b[2] > width - m + tol or b[3] > height - m + tol:
                problems.append(f"{name} outside the sheet border")
            if b[1] < table_top - tol:
                problems.append(f"{name} reaches the bend table / title block")
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                (n1, a), (n2, b) = boxes[i], boxes[j]
                if a[0] < b[2] - tol and b[0] < a[2] - tol and a[1] < b[3] - tol and b[1] < a[3] - tol:
                    problems.append(f"{n1} overlaps {n2}")
        return problems

    def _set_angle(self, view, angle):
        try:
            if abs(float(self._get(view, "Angle") or 0.0) - angle) > 1e-6:
                view.Angle = angle
            return True
        except Exception:
            return False

    def _scale_up(self, s):
        """Next larger scale than s (None if s is already the largest allowed)."""
        max_scale = float(self.layout.get("max_scale", 2.0) or 2.0)
        if self.layout.get("scale_mode", "fine") == "fine" and s < 0.999:
            n = round(1.0 / s)
            up = 1.0 / (n - 1) if n > 1 else None
        else:
            bigger = [v for v in SCALES if v > s * 1.001]
            up = min(bigger) if bigger else None
        return up if up is not None and up <= max_scale * 1.001 else None

    def _arrangements(self, names):
        """Candidate arrangements of the folded views: (label, rows, rotated names).
        rows = [[name, ...], ...] top to bottom, views in a row left to right."""
        grid = {}
        for n in names:
            c, r = FOLDED_GRID.get(n, (9, 1))
            grid.setdefault(r, []).append((c, n))
        standard = [[n for _, n in sorted(grid[r])] for r in sorted(grid)]
        out = [("standard", standard, frozenset())]
        if self.layout.get("folded_arrange", "auto") != "auto" or not 1 < len(names) <= 4:
            return out
        rot_sets = [frozenset()]
        if self.layout.get("folded_rotate", True):
            rot_sets = [frozenset(c) for k in range(len(names) + 1) for c in itertools.combinations(names, k)]
        seen = set()
        for perm in itertools.permutations(names):
            for cuts in range(2 ** (len(names) - 1)):
                rows, row = [], [perm[0]]
                for i in range(1, len(perm)):
                    if (cuts >> (i - 1)) & 1:
                        rows.append(row)
                        row = []
                    row.append(perm[i])
                rows.append(row)
                key = tuple(tuple(r) for r in rows)
                if key in seen:
                    continue
                seen.add(key)
                for rot in rot_sets:
                    out.append(("free", rows, rot))
        return out

    @staticmethod
    def _arrangement_text(cand):
        label, rows, rot = cand
        return f"{label}: " + " / ".join(
            "[" + ", ".join(n + (" (rotated 90)" if n in rot else "") for n in row) + "]" for row in rows)

    def _layout_views(self, drawing, folded, flat_view, table_top, steps, flat_pads=None):
        """Scale every part by its own size and pick, per part, the arrangement of the folded views
        (order in rows + optional 90 deg rotation) that lets them be drawn largest; nothing may overlap."""
        width, height = self._sheet_size(drawing)
        lay = self.layout
        mm = 0.001
        m, pad = lay["margin_mm"] * mm, lay["side_pad_mm"] * mm
        dim_h, dim_v = self._dim_room()
        cg = 2 * mm                                  # extra gap between two view cells
        band_gap = lay["band_gap_mm"] * mm
        max_scale = float(lay.get("max_scale", 2.0) or 2.0)
        k_ratio = float(lay.get("folded_ratio", 0.6) or 0.6)
        common = bool(lay["common_scale"])
        pads = dict(self._initial_flat_pads() if flat_pads is None else flat_pads)
        left, right = m + pad, width - m - pad
        low = table_top + lay["table_clearance_mm"] * mm
        high = height - m - lay["top_pad_mm"] * mm
        avail_w, avail_h = right - left, high - low
        try:
            if lay.get("auto_rotate_flat", True):
                o = self._get(flat_view, "GetOutline")
                if (o[3] - o[1]) > (o[2] - o[0]) * 1.05:
                    try:
                        flat_view.Angle = math.pi / 2.0
                        steps.append("flat pattern rotated 90 deg (long side horizontal)")
                    except Exception as exc:
                        steps.append(f"flat pattern rotation failed: {exc}")

            # ---- unit sizes (scale 1); folded views measured unrotated ----
            views = dict(folded)
            names = [n for n, _ in folded]
            unit = {}
            for n in names:
                self._set_angle(views[n], 0.0)
                unit[n] = self._unit_size_and_border(views[n])
            fw, fh, fbw, fbh = self._unit_size_and_border(flat_view)

            def flat_w(s): return fw * s + fbw + dim_v + pads.get("left", 0.0) + pads.get("right", 0.0)
            def flat_h(s): return fh * s + fbh + dim_h + pads.get("top", 0.0) + pads.get("bottom", 0.0)

            def block(rows, rot, s):
                """(width, height) of the folded block at scale s, incl. outline borders and dimension room."""
                bw_ = bh_ = 0.0
                for row in rows:
                    rw = rh = 0.0
                    for n in row:
                        w, h, ow, oh = unit[n]
                        if n in rot:
                            w, h, ow, oh = h, w, oh, ow
                        rw += w * s + ow + dim_v
                        rh = max(rh, h * s + oh + dim_h)
                    bw_ = max(bw_, rw + cg * (len(row) - 1))
                    bh_ += rh
                return bw_, bh_ + cg * (len(rows) - 1)

            def fits(cand, sf, sb):
                if flat_w(sb) > avail_w + 1e-6:
                    return False
                if not names:
                    return flat_h(sb) <= avail_h + 1e-6
                bw_, bh_ = block(cand[1], cand[2], sf)
                return bw_ <= avail_w + 1e-6 and bh_ + band_gap + flat_h(sb) <= avail_h + 1e-6

            Lb = max(fw, fh, 1e-9)                                   # longest side of the flat pattern
            Lf = max([max(u[0], u[1]) for u in unit.values()] + [1e-9])  # longest side of the folded part
            ratio = 1.0 if common else k_ratio * Lb / Lf             # folded scale / flat scale (target)

            def shrink(sf, sb):
                sb = self._fit_scale(sb * 0.97)
                sf = sb if (common or not names) else self._fit_scale(sf * 0.97)
                return sf, sb

            def solve(cand):
                lo, hi = 0.0, max_scale
                for _ in range(50):                      # largest flat scale with the folded views at target
                    mid = (lo + hi) / 2.0
                    if fits(cand, min(ratio * mid, max_scale), mid):
                        lo = mid
                    else:
                        hi = mid
                sb = self._fit_scale(lo)
                sf = self._fit_scale(min(ratio * lo, max_scale)) if names else sb
                if common:
                    sf = sb = min(sf, sb)
                for _ in range(300):
                    if fits(cand, sf, sb):
                        break
                    sf, sb = shrink(sf, sb)
                for _ in range(300):                     # use the space left by rounding: grow the one behind
                    grown = False
                    order = ("f", "b") if (names and sf < ratio * sb) else ("b", "f")
                    for g in order:
                        if common or not names:
                            up = self._scale_up(sb)
                            if up and fits(cand, up, up):
                                sf = sb = up
                                grown = True
                            break
                        up = self._scale_up(sf if g == "f" else sb)
                        ok = bool(up) and (fits(cand, up, sb) if g == "f" else fits(cand, sf, up))
                        if ok:
                            if g == "f":
                                sf = up
                            else:
                                sb = up
                            grown = True
                            break
                    if not grown:
                        break
                return sf, sb

            candidates = self._arrangements(names) if names else [("none", [], frozenset())]
            best = None
            for cand in candidates:
                sf, sb = solve(cand)
                if names:   # the smaller of (folded vs its target, flat), then total size as tie-break
                    drawn = min(sf * Lf / max(k_ratio, 1e-6), sb * Lb) + 0.05 * (sf * Lf + sb * Lb)
                else:
                    drawn = sb * Lb
                pref = (1.0 if cand[0] in ("standard", "none") else 0.9) * (0.88 ** len(cand[2]))   # rotate only for a clear gain
                if best is None or drawn * pref > best[0] + 1e-12:
                    best = (drawn * pref, cand, sf, sb)
            _, cand, sf, sb = best
            _, rows, rot = cand
            for n in names:
                if n in rot and not self._set_angle(views[n], math.pi / 2.0):
                    steps.append(f"{n} view could not be rotated - standard arrangement used")
                    cand = candidates[0]
                    _, rows, rot = cand
                    sf, sb = solve(cand)
                    break

            problems = []
            for attempt in range(10):
                if names:
                    bw_, bh_ = block(rows, rot, sf)
                    extra = max(0.0, avail_h - bh_ - band_gap - flat_h(sb))
                    f_bot = high - bh_ - extra / 2.0
                    folded_box = (left, f_bot, right, high)
                    flat_box = (left, low, right, f_bot - band_gap)
                else:
                    folded_box, flat_box = None, (left, low, right, high)
                b_in = (flat_box[0] + dim_v + pads.get("left", 0.0), flat_box[1] + pads.get("bottom", 0.0),
                        flat_box[2] - pads.get("right", 0.0), flat_box[3] - dim_h - pads.get("top", 0.0))

                scales = {"flat": sb}
                if names:
                    scales["folded"] = sf
                b_w, b_h = b_in[2] - b_in[0], b_in[3] - b_in[1]
                scales["flat"], _ = self._settle_scale(flat_view, scales["flat"], b_w, b_h)
                if common and names:
                    scales["folded"] = scales["flat"] = min(scales["folded"], scales["flat"])
                    flat_view.ScaleDecimal = scales["flat"]
                for n in names:
                    views[n].ScaleDecimal = scales["folded"]

                if names:
                    size = {}
                    for n in names:
                        o = self._get(views[n], "GetOutline")
                        size[n] = (o[2] - o[0], o[3] - o[1])
                    row_dims = [(sum(size[n][0] + dim_v for n in row) + cg * (len(row) - 1),
                                 max(size[n][1] + dim_h for n in row)) for row in rows]
                    tot_w = max(rw for rw, _ in row_dims)
                    tot_h = sum(rh for _, rh in row_dims) + cg * (len(rows) - 1)
                    x0 = left + (avail_w - tot_w) / 2.0
                    y = (folded_box[1] + folded_box[3]) / 2.0 + tot_h / 2.0
                    for row, (_rw, rh) in zip(rows, row_dims):
                        x = x0
                        for n in row:
                            w_, _h = size[n]
                            # view at the bottom-left of its cell: dimension room above and to the right
                            self._move_view(views[n], x + w_ / 2.0, y - dim_h - (rh - dim_h) / 2.0)
                            x += w_ + dim_v + cg
                        y -= rh + cg
                self._move_view(flat_view, (b_in[0] + b_in[2]) / 2.0, (b_in[1] + b_in[3]) / 2.0)

                problems = self._layout_problems(folded, flat_view, pads, table_top, width, height)
                if not problems:
                    break
                steps.append(f"overlap check (attempt {attempt + 1}): {problems} -> views reduced one step")
                sf, sb = shrink(scales.get("folded", scales["flat"]), scales["flat"])

            steps.append(f"sheet {width * 1000:.0f} x {height * 1000:.0f} mm | table top {table_top * 1000:.0f} mm | "
                         f"flat area {flat_box[1] * 1000:.0f}-{flat_box[3] * 1000:.0f} mm"
                         + (f" | folded area {folded_box[1] * 1000:.0f}-{folded_box[3] * 1000:.0f} mm" if folded_box else "")
                         + " | bend dim room " + ", ".join(f"{k} {v * 1000:.0f}" for k, v in pads.items() if v))
            if names:
                steps.append(f"folded arrangement ({len(candidates)} tried) -> {self._arrangement_text(cand)}")
            o = self._get(flat_view, "GetOutline")
            steps.append(f"flat pattern at {self._scale_text(scales['flat'])}: "
                         f"{(o[2] - o[0]) * 1000:.0f} x {(o[3] - o[1]) * 1000:.0f} mm "
                         f"in a box of {b_w * 1000:.0f} x {b_h * 1000:.0f} mm")
            if names:
                steps.append("part-size scaling: longest side drawn "
                             f"folded {scales['folded'] * Lf * 1000:.0f} mm, flat {scales['flat'] * Lb * 1000:.0f} mm")
            steps.append("overlap check: " + ("OK - nothing overlaps" if not problems
                                              else f"STILL OVERLAPPING {problems} - check manually"))
            steps.append("layout scales: " + ", ".join(f"{k} {self._scale_text(v)}" for k, v in scales.items()))
            return scales["flat"]
        except Exception as exc:
            steps.append(f"layout failed: {type(exc).__name__}: {exc}")
            return None

    def _set_sheet_scale(self, drawing, scale, steps):
        text = self._scale_text(scale)
        num, den = (1.0, float(round(1 / scale))) if scale < 1 else (float(round(scale)), 1.0)
        try:
            sheet = self._get(drawing, "GetCurrentSheet")
            sheet.SetScale(num, den, False, False)
            steps.append(f"sheet scale set to {text}")
        except Exception as exc:
            steps.append(f"sheet scale not set ({exc})")
        try:
            drawing.Extension.CustomPropertyManager("").Add3("Scale", 30, text, 2)
        except Exception:
            pass

    # ---------------- coordinate transforms ----------------
    def _check_convention(self, com_xform, data, label, steps):
        """Verify the MathTransform matrix convention ONCE per session by letting SolidWorks transform a probe
        point (IMathUtility.CreatePoint + IMathPoint.MultiplyTransform). Returns the convention or None."""
        if self._convention or self._convention_check_failed:
            return self._convention
        probe = (0.1234, -0.5678, 0.9012)
        failures = []

        def attempt(stage, *calls):
            for how, call in calls:
                try:
                    value = call()
                    if value is not None:
                        return value
                    failures.append(f"{stage}[{how}]: returned nothing")
                except Exception as exc:
                    failures.append(f"{stage}[{how}]: {exc}")
            raise RuntimeError(stage)

        try:
            math_util = attempt("GetMathUtility", ("property/method", lambda: self._get(self.sw_app, "GetMathUtility")),
                                ("call", lambda: self.sw_app.GetMathUtility()))
            point = attempt("CreatePoint",
                            ("VARIANT", lambda: math_util.CreatePoint(
                                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, probe))),
                            ("list", lambda: math_util.CreatePoint(list(probe))))
            moved = attempt("MultiplyTransform", ("MultiplyTransform", lambda: point.MultiplyTransform(com_xform)),
                            ("IMultiplyTransform", lambda: point.IMultiplyTransform(com_xform)))
            out = attempt("ArrayData", ("property", lambda: self._get(moved, "ArrayData")))
        except RuntimeError:
            self._convention_check_failed = True
            steps.append(f"{label}: SolidWorks convention check not possible ({'; '.join(failures)}) - "
                         "the bend geometry will decide between row and column form")
            return None
        convention = detect_convention(data, probe, [float(v) for v in out])
        if convention is None:
            self._convention_check_failed = True
            steps.append(f"{label}: MathPoint result matches neither matrix form - geometry will decide")
            return None
        self._convention, self._convention_source = convention, "IMathPoint.MultiplyTransform"
        steps.append(f"transform convention verified with IMathPoint.MultiplyTransform: {convention} vectors")
        return convention

    def _to_xform(self, com_xform, label, steps):
        if com_xform is None:
            raise RuntimeError(f"{label} not available")
        data = self._get(com_xform, "ArrayData")
        if not data or len(data) < 13:
            raise RuntimeError(f"{label}: ArrayData not available")
        data = [float(v) for v in data]
        convention = self._check_convention(com_xform, data, label, steps) or "row"
        return Xform.from_array(data, convention)

    def _sheet_mapper(self, view, steps, points=None):
        """model xyz -> sheet xy for this view, built from IView.ModelToViewTransform and checked against
        IView.GetOutline."""
        mv = self._to_xform(self._get(view, "ModelToViewTransform"), "ModelToViewTransform", steps)
        if points is None:
            points = self._visible_model_points(view)
        outline = self._get(view, "GetOutline")
        k, dx, dy, status = 1.0, 0.0, 0.0, "not checked (no vertices or outline)"
        to_sheet_ok = False
        if points and outline:
            mapped = [mv.apply(p)[:2] for p in points]
            bx0, by0 = min(p[0] for p in mapped), min(p[1] for p in mapped)
            bx1, by1 = max(p[0] for p in mapped), max(p[1] for p in mapped)
            bw, bh = bx1 - bx0, by1 - by0
            ow, oh = outline[2] - outline[0], outline[3] - outline[1]
            pad_w, pad_h = ow - bw, oh - bh
            inside = (bx0 >= outline[0] - 0.001 and by0 >= outline[1] - 0.001
                      and bx1 <= outline[2] + 0.001 and by1 <= outline[3] + 0.001)
            if inside and -0.001 <= min(pad_w, pad_h) and max(pad_w, pad_h) <= 0.03:
                status = "inside the view outline (arcs / hidden edges make the outline larger)"
                if abs(pad_w - pad_h) <= 0.0005 + 0.01 * max(ow, oh):
                    status = "matches the view outline" + (f" (outline border {(pad_w + pad_h) / 4 * 1000:.1f} mm "
                                                           "per side)" if (pad_w + pad_h) > 0.0008 else "")
                to_sheet_ok = True
            elif (abs(pad_w - pad_h) <= 0.0005 + 0.01 * max(ow, oh)
                    and -0.001 <= min(pad_w, pad_h) and max(pad_w, pad_h) <= 0.03):
                border = (pad_w + pad_h) / 4.0
                status = "matches the view outline" + (f" (outline border {border * 1000:.1f} mm per side)"
                                                       if border > 0.0002 else "")
            else:
                kw = ow / bw if bw > 1e-9 else None
                kh = oh / bh if bh > 1e-9 else None
                if kw and kh and abs(kw - kh) <= 0.03 * max(kw, kh):
                    k = (kw + kh) / 2.0
                    status = f"scale corrected by {k:.6g} to match the view outline"
                elif -0.001 <= min(pad_w, pad_h) and max(pad_w, pad_h) <= 0.03:
                    status = "smaller than the view outline (arcs / hidden edges) - centred on it"
                else:
                    raise RuntimeError(
                        f"drawing-view transform does not match the view outline (mapped {bw * 1000:.1f} x "
                        f"{bh * 1000:.1f} mm, outline {ow * 1000:.1f} x {oh * 1000:.1f} mm)")
            cx, cy = k * (bx0 + bx1) / 2.0, k * (by0 + by1) / 2.0
            ox, oy = (outline[0] + outline[2]) / 2.0, (outline[1] + outline[3]) / 2.0
            if not to_sheet_ok and (abs(ox - cx) > 0.002 + 0.03 * ow or abs(oy - cy) > 0.002 + 0.03 * oh):
                dx, dy = ox - cx, oy - cy
                status += f"; position corrected by ({dx * 1000:.1f}, {dy * 1000:.1f}) mm"

        def to_sheet(xyz):
            x, y, _z = mv.apply(xyz)
            return (k * x + dx, k * y + dy)

        to_sheet.status = status
        to_sheet.xform = mv
        to_sheet.points = points
        return to_sheet

    def _view_mapper(self, view, steps, verbose=False):
        mapper = self._sheet_mapper(view, steps)
        if not mapper.points:
            raise RuntimeError("no vertices in the view")
        if verbose:
            steps.append(f"view transform: {mapper.status}")

        def to_sheet(xyz, via=None):
            return mapper(via.apply(xyz) if via is not None else xyz)
        pts = [mapper(p) for p in mapper.points]
        to_sheet.vb = (min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts))
        to_sheet.status = mapper.status
        return to_sheet

    # ---------------- overall dimensions ----------------
    def _view_vertices(self, view):
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        for entity_type in (2, 1, 3, 4):
            try:
                entities = view.GetVisibleEntities2(nothing, entity_type)
            except Exception:
                continue
            if entities and self._get(entities[0], "GetPoint") is not None:
                return list(entities)
        return []

    def _visible_model_points(self, view):
        points = []
        for vertex in self._view_vertices(view):
            p = self._get(vertex, "GetPoint")
            if p:
                points.append(tuple(float(c) for c in p[:3]))
        if not points:
            for edge in self._view_edges(view):
                points.extend(p for p in self._edge_points(edge) if p is not None)
        return points

    def _vertex_points(self, view, vertices, steps):
        to_sheet = self._view_mapper(view, steps)
        points = []
        for vertex in vertices:
            p = self._get(vertex, "GetPoint")
            if not p:
                continue
            x, y = to_sheet((float(p[0]), float(p[1]), float(p[2])))
            points.append((x, y, vertex))
        return points

    def _select_vertex(self, view, vertex, append, sel):
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

    _select_entity = _select_vertex

    def _select_data(self, drawing, view):
        try:
            sel = drawing.SelectionManager.CreateSelectData()
            sel.View = view
            return sel
        except Exception:
            return VARIANT(pythoncom.VT_DISPATCH, None)

    def _add_overall_dimensions(self, drawing, view, steps, label, length=True, width=True, width_side="right"):
        stage = "vertices"
        try:
            try:
                points = self._vertex_points(view, self._view_vertices(view), steps)
            except Exception as exc:
                steps.append(f"dimensions {label}: NOT created - view transform unusable ({exc})")
                return
            if len(points) < 2:
                steps.append(f"dimensions {label}: no vertices found")
                return
            left, right = min(points, key=lambda t: t[0]), max(points, key=lambda t: t[0])
            bottom, top = min(points, key=lambda t: t[1]), max(points, key=lambda t: t[1])

            outline = self._get(view, "GetOutline")
            off = self.dimension_offset_mm / 1000.0
            sel = self._select_data(drawing, view)

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
                ok1, _ = self._select_vertex(view, a_pt[2], False, sel)
                ok2, _ = self._select_vertex(view, b_pt[2], True, sel)
                if not (ok1 and ok2):
                    made.append(f"{name}: could not select corners")
                    continue
                stage = f"{name}: {func}"
                dim = getattr(drawing, func)(px, py, 0.0)
                if dim is None:
                    made.append(f"{name}: failed")
                    continue
                axis = 0 if name == "length" else 1
                scale = float(self._get(view, "ScaleDecimal") or 1.0)
                expected = abs(b_pt[axis] - a_pt[axis]) / scale
                actual = self._dimension_value(dim)
                if actual is None:
                    made.append(f"{name}: ok ({expected * 1000:.1f} mm, value not readable)")
                elif abs(actual - expected) <= max(0.0002, 0.002 * expected):
                    made.append(f"{name}: ok {actual * 1000:.1f} mm verified")
                else:
                    made.append(f"{name}: {actual * 1000:.1f} mm - MISMATCH, expected {expected * 1000:.1f} mm")
            drawing.ClearSelection2(True)
            steps.append(f"dimensions {label} -> " + ", ".join(made))
        except Exception as exc:
            steps.append(f"dimensions {label} failed at '{stage}': {exc}")

    # ---------------- bend dimensions: geometry ----------------
    def _view_edges(self, view, steps=None):
        """Visible model edges of the view (classified by their methods, not by entity type codes)."""
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        found = []
        diagnostics = []

        def as_items(value):
            if value is None:
                return []
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
            return (self._get(obj, "GetCurve") is not None and
                    (self._get(obj, "GetStartPoint2") is not None or
                     self._get(obj, "GetEndPoint2") is not None))

        for entity_type in (1, 2, 3, 4):
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

    def _edge_points(self, edge):
        """(start, end) model points of an edge, or (None, None) for closed edges such as full circles."""
        out = []
        for vertex_name, point_name in (("GetStartVertex", "GetStartPoint2"), ("GetEndVertex", "GetEndPoint2")):
            p = self._get(self._get(edge, vertex_name), "GetPoint")
            if p is None:
                p = self._get(edge, point_name)
            try:
                out.append(tuple(float(p[i]) for i in range(3)) if p is not None else None)
            except Exception:
                out.append(None)
        return tuple(out)

    def _segment_xyz(self, segment):
        points = []
        for name in ("GetStartPoint2", "GetEndPoint2"):
            pt = self._get(segment, name)
            points.append(tuple(float(self._get(pt, c)) for c in ("X", "Y", "Z")))
        return points

    def _read_bend_data(self, view, steps):
        """Everything the bend analysis needs from SolidWorks, read once (no coordinate math here)."""
        raw = {"lines": list(self._get(view, "GetBendLines") or ()), "bends": [], "edges": [], "faces": []}
        steps.append(f"bend lines in the flat pattern view: {len(raw['lines'])}")
        if not raw["lines"]:
            return raw
        for index, seg in enumerate(raw["lines"]):
            record = {"index": index, "id": f"#{index + 1}", "seg": seg}
            try:
                sketch = self._get(seg, "GetSketch")
                if sketch is None:
                    raise RuntimeError("GetSketch returned nothing")
                com_xform = self._get(sketch, "ModelToSketchTransform")
                data = self._get(com_xform, "ArrayData")
                if not data or len(data) < 13:
                    raise RuntimeError("ModelToSketchTransform.ArrayData not available")
                record["sketch_data"] = [float(v) for v in data]
                record["sketch_points"] = self._segment_xyz(seg)
                self._check_convention(com_xform, record["sketch_data"], "ModelToSketchTransform", steps)
            except Exception as exc:
                record["error"] = f"{type(exc).__name__}: {exc}"
            raw["bends"].append(record)
        for entity in self._view_edges(view, steps):
            p, q = self._edge_points(entity)
            if p is None or q is None:
                continue
            curve = self._get(entity, "GetCurve")
            is_line = self._get(curve, "IsLine") if curve is not None else None
            raw["edges"].append({"entity": entity, "m0": p, "m1": q,
                                 "is_line": True if is_line is None else bool(is_line)})
        for record in raw["edges"][:300]:
            for face in (self._get(record["entity"], "GetTwoAdjacentFaces2") or ()):
                if face is None:
                    continue
                normal, area = self._get(face, "Normal"), self._get(face, "GetArea")
                try:
                    if normal and area and math.sqrt(sum(float(c) ** 2 for c in normal)) > 0.5:
                        raw["faces"].append((float(area), tuple(float(c) for c in normal), face))
                except Exception:
                    continue
        return raw

    def _outer_loop_segments(self, face, model_to_frame, cache):
        key = id(face)
        if key not in cache:
            model_segments = []
            for loop in (self._get(face, "GetLoops") or ()):
                if not self._get(loop, "IsOuter"):
                    continue
                for edge in (self._get(loop, "GetEdges") or ()):
                    p, q = self._edge_points(edge)
                    if p is not None and q is not None:
                        model_segments.append((p, q))
            cache[key] = model_segments
        return [(model_to_frame.apply(p)[:2], model_to_frame.apply(q)[:2]) for p, q in cache[key]]

    @staticmethod
    def _nearest_parallel_edge(p0, p1, edges, outer):
        """(edge id, distance, 'outer'/'inner') of the nearest edge parallel to the bend line p0-p1 (frame xy).
        Edges that run alongside the bend line are preferred; outer boundary before any other edge."""
        bx, by = p1[0] - p0[0], p1[1] - p0[1]
        bl = math.hypot(bx, by)
        if bl < 1e-9:
            return None
        ux, uy = bx / bl, by / bl
        best = {}
        for e in edges:
            ex, ey = e["p1"][0] - e["p0"][0], e["p1"][1] - e["p0"][1]
            el = math.hypot(ex, ey)
            if el < 1e-6 or not e.get("is_line", True) or abs(ux * ey - uy * ex) / el > 1e-3:
                continue
            mx, my = (e["p0"][0] + e["p1"][0]) / 2.0, (e["p0"][1] + e["p1"][1]) / 2.0
            d = abs(ux * (my - p0[1]) - uy * (mx - p0[0]))
            if d < 1e-5:
                continue
            t0 = ux * (e["p0"][0] - p0[0]) + uy * (e["p0"][1] - p0[1])
            t1 = ux * (e["p1"][0] - p0[0]) + uy * (e["p1"][1] - p0[1])
            alongside = min(max(t0, t1), bl) - max(min(t0, t1), 0.0) > 1e-4
            group = "outer" if e["id"] in outer else "inner"
            key = (0 if alongside else 1, d)
            if group not in best or key < best[group][0]:
                best[group] = (key, e["id"], d)
        for group in ("outer", "inner"):
            if group in best:
                return best[group][1], best[group][2], group
        return None

    def _complete_bends(self, results, records, edges, outer):
        """One dimension per bend (= per tag): shared bends get their own dimension, bends without a reference
        are measured to the nearest parallel edge."""
        by_id = {r["id"]: r for r in results}
        rec = {r["id"]: r for r in records if "error" not in r}
        for r in results:
            if r["status"] in ("merged", "duplicate"):
                src = by_id.get(r.get("merged_into"))
                if src and src["status"] == "ok" and src.get("edge") is not None:
                    r.update(status="ok", edge=src["edge"], value=src["value"],
                             note=f"own dimension (was shared with bend {src['id']})")
        for r in results:
            if r["status"] == "ok" or r["id"] not in rec:
                continue
            found = self._nearest_parallel_edge(rec[r["id"]]["p0"], rec[r["id"]]["p1"], edges, outer)
            if found is not None:
                edge_id, dist, group = found
                r.update(status="ok", edge=edge_id, value=dist,
                         note=f"measured to the nearest parallel {group} edge (was {r['status']}: {r.get('reason')})")
                continue
            corner = self._flange_corner(rec[r["id"]]["p0"], rec[r["id"]]["p1"], edges, outer)
            if corner is not None:
                edge_id, end, dist = corner
                r.update(status="ok", edge=edge_id, vertex=end, value=dist,
                         note=f"no parallel outer edge - measured to the outer corner of the flange "
                              f"(was {r['status']}: {r.get('reason')})")

    @staticmethod
    def _flange_corner(p0, p1, edges, outer):
        """(edge id, 0/1 end, distance) of the outer corner farthest from the bend line on the flange side
        (= the side that reaches less far). Used when no outer edge is parallel to the bend line."""
        bx, by = p1[0] - p0[0], p1[1] - p0[1]
        bl = math.hypot(bx, by)
        if bl < 1e-9:
            return None
        ux, uy = bx / bl, by / bl
        sides = {1: None, -1: None}
        for e in edges:
            if e["id"] not in outer:
                continue
            for end, q in ((0, e["p0"]), (1, e["p1"])):
                t = ux * (q[0] - p0[0]) + uy * (q[1] - p0[1])
                if t < -0.25 * bl or t > 1.25 * bl:        # only corners alongside this bend line
                    continue
                d = ux * (q[1] - p0[1]) - uy * (q[0] - p0[0])
                if abs(d) < 1e-5:
                    continue
                side = 1 if d > 0 else -1
                if sides[side] is None or abs(d) > sides[side][2]:
                    sides[side] = (e["id"], end, abs(d))
        found = [v for v in sides.values() if v is not None]
        return min(found, key=lambda v: v[2]) if found else None

    def _bend_analysis_for(self, raw, convention, item, steps, loop_cache):
        """Bend analysis for one MathTransform convention. Raises when the frames are inconsistent."""
        analysis = {"error": None, "convention": convention, "bends": [], "edges": [], "results": [],
                    "outer": set(), "outer_info": {}, "bend_count": len(raw["lines"])}
        # --- bend lines in the measuring frame (= sketch frame of the first readable bend line) ---
        reference = reference_data = None
        records = []
        for src in raw["bends"]:
            record = {k: v for k, v in src.items() if k not in ("sketch_data", "sketch_points")}
            if "error" not in src:
                model_to_sketch = Xform.from_array(src["sketch_data"], convention)
                if reference is None:
                    reference, reference_data = model_to_sketch, src["sketch_data"]
                pts = src["sketch_points"]
                same = all(abs(a - b) <= 1e-12 for a, b in zip(src["sketch_data"], reference_data))
                if not same:
                    to_frame = model_to_sketch.inverse().then(reference)
                    pts = [to_frame.apply(p) for p in pts]
                record.update(p0=pts[0][:2], p1=pts[1][:2], z=(pts[0][2], pts[1][2]), own_sketch=not same)
            records.append(record)
        if reference is None:
            raise RuntimeError("no bend line has a readable sketch transform (ISketch.ModelToSketchTransform)")
        model_to_frame = reference
        analysis["model_to_frame"], analysis["bends"] = model_to_frame, records
        bad = [r for r in records if "error" in r]
        if bad:
            steps.append("bend lines unreadable (reported, not dimensioned): "
                         + ", ".join(f"{r['id']} ({r['error']})" for r in bad))

        # --- edges in the frame ---
        edges, zs = [], []
        for src in raw["edges"]:
            f0, f1 = model_to_frame.apply(src["m0"]), model_to_frame.apply(src["m1"])
            zs += [f0[2], f1[2]]
            record = dict(src, id=len(edges), f0=f0, f1=f1, p0=f0[:2], p1=f1[:2])
            if math.hypot(f1[0] - f0[0], f1[1] - f0[1]) < 1e-6:
                continue
            if any(same_segment(record["p0"], record["p1"], e["p0"], e["p1"], 1e-5) for e in edges):
                continue
            edges.append(record)
        analysis["edges"] = edges
        steps.append(f"sheet edges in the bend frame: {len(edges)}")
        if not edges:
            raise RuntimeError("no visible edges with readable end points")

        thickness = float((item or {}).get("thickness") or 0.0) / 1000.0
        allowed = (2.0 * thickness + 0.0005) if thickness > 0 else 0.01
        spread = max(zs) - min(zs)
        analysis["frame_check"] = {"z_spread_mm": round(spread * 1000, 4), "allowed_mm": round(allowed * 1000, 4)}
        if spread > allowed:
            raise RuntimeError(f"edges are not in the bend-sketch plane (spread {spread * 1000:.2f} mm > "
                               f"{allowed * 1000:.2f} mm) - coordinate frames do not match")

        # --- outer boundary: SolidWorks' outer loop of the largest face parallel to the sheet ---
        hint = None
        faces = [f for f in raw["faces"] if abs(model_to_frame.normal_z(f[1])) >= 0.999]
        if faces:
            area, _normal, face = max(faces, key=lambda f: f[0])
            hint = self._outer_loop_segments(face, model_to_frame, loop_cache) or None
            steps.append(f"outer boundary: face of {area * 1e6:.0f} mm2, outer loop with {len(hint or ())} edge(s)")
        else:
            steps.append("outer boundary: no sheet face found through the API - using the edge loops")
        outer, info = classify_outer([(e["p0"], e["p1"]) for e in edges], hint)
        analysis["outer"], analysis["outer_info"] = outer, info
        steps.append(f"outer boundary: {len(outer)} edge(s) via {info['method']}"
                     + (f" - WARNING: {info['warning']}" if info.get("warning") else ""))
        if not outer:
            raise RuntimeError("outer boundary of the flat pattern could not be determined")

        valid = [r for r in records if "error" not in r]
        results = measure_bends([{"id": r["id"], "p0": r["p0"], "p1": r["p1"]} for r in valid],
                                edges, outer, info.get("polygon"), merge_collinear=self.bend_dim_merge_collinear)
        for r in bad:
            results.append({"id": r["id"], "status": "invalid", "reason": r["error"], "value": None})
        analysis["results"] = results
        outside = [r for r in results if r["status"] == "outside_boundary"]
        if valid and len(outside) * 2 > len(valid):
            raise RuntimeError(f"{len(outside)} of {len(valid)} bend lines fall outside the outer boundary - "
                               "bend-line and edge coordinates do not match")
        # every bend (= every tag) gets its own dimension
        self._complete_bends(results, valid, edges, outer)
        touch = max(0.0005, 2 * thickness)
        analysis["endpoints_on_edges"] = sum(
            1 for r in valid for p in (r["p0"], r["p1"])
            if min(point_segment_distance(p, e["p0"], e["p1"]) for e in edges) <= touch)
        return analysis

    @staticmethod
    def _result_signature(analysis):
        return [(r["id"], r["status"], round(r["value"], 7) if r.get("value") is not None else None)
                for r in analysis["results"]]

    def _analyze_bends(self, view, steps, item=None):
        """Measure every bend line to its outer reference edge. Never raises: problems go to analysis['error']."""
        analysis = {"error": None, "bends": [], "edges": [], "results": [], "outer": set(), "outer_info": {}}
        try:
            raw = self._read_bend_data(view, steps)
            analysis["bend_count"] = len(raw["lines"])
            if not raw["lines"]:
                return analysis
            conventions = [self._convention] if self._convention else ["row", "column"]
            outcomes, loop_cache = [], {}
            for convention in conventions:
                sub = []
                try:
                    outcomes.append((convention, self._bend_analysis_for(raw, convention, item, sub, loop_cache), sub, None))
                except Exception as exc:
                    outcomes.append((convention, None, sub, exc))
            good = [o for o in outcomes if o[1] is not None]
            if not good:
                convention, _, sub, exc = outcomes[0]
                steps.extend(sub)
                raise exc
            chosen = good[0]
            if len(conventions) == 2:
                if len(good) == 2:
                    if self._result_signature(good[0][1]) == self._result_signature(good[1][1]):
                        steps.append("matrix convention: both forms give identical bend results")
                    else:
                        scores = [g[1]["endpoints_on_edges"] for g in good]
                        if scores[0] == scores[1]:
                            raise RuntimeError("matrix convention ambiguous: row and column forms both fit the "
                                               "geometry but give different bend results - not guessing")
                        chosen = good[scores.index(max(scores))]
                        self._convention, self._convention_source = chosen[0], "geometry (bend frame)"
                else:
                    self._convention, self._convention_source = chosen[0], "geometry (bend frame)"
                    rejected = next(o for o in outcomes if o[1] is None)
                    steps.append(f"matrix convention from geometry: {chosen[0]} vectors "
                                 f"({rejected[0]} rejected: {rejected[3]})")
            steps.extend(chosen[2])
            analysis = chosen[1]
        except Exception as exc:
            analysis["error"] = f"{type(exc).__name__}: {exc}"
            steps.append(f"bend dimensions failed: {analysis['error']}")
        return analysis

    # ---------------- annotations: real text boxes ----------------
    def _view_notes(self, view):
        notes = self._get(view, "GetNotes")
        items = list(notes) if notes else []
        if not items:
            note = self._get(view, "GetFirstNote")
            for _ in range(5000):
                if note is None:
                    break
                items.append(note)
                note = self._get(note, "GetNext")
        out = []
        for note in items:
            text = self._get(note, "GetText")
            position = self._get(self._get(note, "GetAnnotation"), "GetPosition")
            if text and position and len(position) >= 2:
                out.append((str(text).strip(), (float(position[0]), float(position[1]))))
        return out

    def _text_box(self, annotation, text=""):
        """Sheet box (x0, y0, x1, y1) of an annotation's text, from IDisplayData; estimated if not readable."""
        h0 = self.bend_dim_text_height_mm / 1000.0
        boxes = []
        try:
            dd = annotation.GetDisplayData()
            for i in range(int(self._get(dd, "GetTextCount") or 0)):
                t = str(dd.GetTextAtIndex(i) or "")
                pos = dd.GetTextPositionAtIndex(i)
                h = float(dd.GetTextHeightAtIndex(i) or h0)
                ang = float(dd.GetTextAngleAtIndex(i) or 0.0)
                w = max(len(t), 1) * 0.75 * h
                ux, uy = math.cos(ang), math.sin(ang)
                x, y = float(pos[0]), float(pos[1])
                pts = [(x, y), (x + ux * w, y + uy * w), (x - uy * h, y + ux * h),
                       (x + ux * w - uy * h, y + uy * w + ux * h)]
                boxes.append((min(p[0] for p in pts), min(p[1] for p in pts),
                              max(p[0] for p in pts), max(p[1] for p in pts)))
        except Exception:
            boxes = []
        if boxes:
            return (min(b[0] for b in boxes), min(b[1] for b in boxes),
                    max(b[2] for b in boxes), max(b[3] for b in boxes))
        pos = self._get(annotation, "GetPosition") or (0.0, 0.0)
        w = max(len(text), 1) * 0.75 * h0
        return (pos[0] - w / 2.0, pos[1] - h0 / 2.0, pos[0] + w / 2.0, pos[1] + h0 / 2.0)

    def _note_box(self, note, annotation, text):
        ext = self._get(note, "GetExtent")
        try:
            if ext and len(ext) >= 5 and ext[3] > ext[0] and ext[4] > ext[1]:
                return (float(ext[0]), float(ext[1]), float(ext[3]), float(ext[4]))
        except Exception:
            pass
        return self._text_box(annotation, text)

    def _view_annotation_items(self, view):
        """Notes (incl. bend tags) and dimensions of a view with their real text boxes on the sheet."""
        anns = []
        a = self._get(view, "GetFirstAnnotation3")
        for _ in range(5000):
            if a is None:
                break
            anns.append(a)
            a = self._get(a, "GetNext3")
        if not anns:
            anns = list(self._get(view, "GetAnnotations") or ())
        items = []
        for ann in anns:
            try:
                kind = {4: "dim", 6: "note"}.get(int(self._get(ann, "GetType") or 0))
                if kind is None:
                    continue
                spec = self._get(ann, "GetSpecificAnnotation")
                pos = self._get(ann, "GetPosition")
                if kind == "note":
                    text = str(self._get(spec, "GetText") or "").strip()
                    box = self._note_box(spec, ann, text)
                else:
                    v = self._dimension_value(spec) if spec is not None else None
                    text = f"{v * 1000:.2f}" if v is not None else "0000.00"
                    box = self._text_box(ann, text)
                items.append({"ann": ann, "kind": kind, "text": text, "box": box,
                              "pos": (float(pos[0]), float(pos[1])) if pos else None})
            except Exception:
                continue
        return items

    def _shift_annotation(self, annotation, dx, dy):
        try:
            pos = self._get(annotation, "GetPosition")
            annotation.SetPosition2(pos[0] + dx, pos[1] + dy, pos[2] if len(pos) > 2 else 0.0)
            return True
        except Exception:
            return False

    def _declutter_flat(self, drawing, view, analysis, plan, steps):
        """Readability pass on the flat pattern: no tag, note or number may touch another one.
        Tags slide along their own bend line, other notes move to the nearest free spot, bend dimension texts
        slide along their dimension line; overall dimensions stay where they are."""
        mm = 0.001
        gap = 0.8 * mm
        items = self._view_annotation_items(view)
        if not items:
            steps.append("readability check: annotations of the flat view could not be read")
            return
        width, height = self._sheet_size(drawing)
        m = self.layout["margin_mm"] * mm
        tag_line = {}
        for bid, tag in ((plan or {}).get("tags") or {}).items():
            if bid in (plan.get("sheet_lines") or {}):
                tag_line[tag] = plan["sheet_lines"][bid]
        for r in (analysis or {}).get("results", []):
            disp = r.get("display")
            pos = self._get(self._get(disp, "GetAnnotation"), "GetPosition") if disp is not None else None
            if not pos:
                continue
            for it in items:
                if (it["kind"] == "dim" and it["pos"] and abs(it["pos"][0] - pos[0]) < 1e-7
                        and abs(it["pos"][1] - pos[1]) < 1e-7):
                    it["slide"] = "x" if r.get("side") in ("top", "bottom") else "y"

        def overlap(a, b):
            return a[0] < b[2] + gap and b[0] < a[2] + gap and a[1] < b[3] + gap and b[1] < a[3] + gap

        def free(box, me):
            inside = box[0] >= m and box[1] >= m and box[2] <= width - m and box[3] <= height - m
            return inside and not any(o is not me and overlap(box, o["box"]) for o in items)

        def candidates(it):
            b = it["box"]
            cx, cy, w, h = (b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0, b[2] - b[0], b[3] - b[1]
            if it["kind"] == "note" and it["text"] in tag_line:
                (x0, y0), (x1, y1) = tag_line[it["text"]]
                length = math.hypot(x1 - x0, y1 - y0) or 1.0
                nx, ny = -(y1 - y0) / length, (x1 - x0) / length
                step = max(w, h) + 1 * mm
                for f in (0.5, 0.35, 0.65, 0.2, 0.8, 0.1, 0.9):
                    for o in (0.0, step, -step, 2 * step, -2 * step):
                        yield (x0 + f * (x1 - x0) + nx * o - cx, y0 + f * (y1 - y0) + ny * o - cy)
            elif it["kind"] == "note":
                for r in (2, 4, 6, 9, 12, 16):
                    for a in range(8):
                        yield (math.cos(a * math.pi / 4) * r * mm, math.sin(a * math.pi / 4) * r * mm)
            elif it.get("slide"):
                for k in (1, -1, 2, -2, 3, -3, 4, -4):
                    yield (k * (w + 1 * mm), 0.0) if it["slide"] == "x" else (0.0, k * (h + 1 * mm))

        movable = [it for it in items if it["kind"] == "note" or it.get("slide")]
        movable.sort(key=lambda it: 0 if it["text"] in tag_line else 1 if it["kind"] == "note" else 2)
        moved = []
        for it in movable:
            if free(it["box"], it):
                continue
            for dx, dy in candidates(it):
                nb = (it["box"][0] + dx, it["box"][1] + dy, it["box"][2] + dx, it["box"][3] + dy)
                if free(nb, it) and self._shift_annotation(it["ann"], dx, dy):
                    it["box"] = nb
                    moved.append(it["text"])
                    break
        residual = [f"'{a['text']}' / '{b['text']}'" for i, a in enumerate(items) for b in items[i + 1:]
                    if overlap(a["box"], b["box"])]
        steps.append("readability check: " + (f"moved {moved}; " if moved else "")
                     + ("OK - no tag / number overlaps" if not residual
                        else f"STILL OVERLAPPING {residual} - check manually"))

    # ---------------- bend dimensions: tags + placement ----------------
    @staticmethod
    def _assign_tags_global(sheet_lines, notes, tag_names):
        """One tag per bend line, chosen so that the sum of tag-to-bend-line distances is smallest."""
        pos = {}
        for text, p in notes:
            if text in tag_names:
                pos.setdefault(text, p)
        names = [t for t in tag_names if t in pos]
        bids = list(sheet_lines)
        result = {}
        if names and bids:
            cost = {(t, b): point_segment_distance(pos[t], *sheet_lines[b]) for t in names for b in bids}
            if len(names) <= 8 and len(bids) <= 8:
                best = None
                if len(names) <= len(bids):
                    for perm in itertools.permutations(bids, len(names)):
                        c = sum(cost[(t, b)] for t, b in zip(names, perm))
                        if best is None or c < best[0]:
                            best = (c, dict(zip(perm, names)))
                else:
                    for perm in itertools.permutations(names, len(bids)):
                        c = sum(cost[(t, b)] for t, b in zip(perm, bids))
                        if best is None or c < best[0]:
                            best = (c, dict(zip(bids, perm)))
                result = best[1]
            else:
                for _c, t, b in sorted((c, t, b) for (t, b), c in cost.items()):
                    if b not in result and t not in result.values():
                        result[b] = t
        return result, [t for t in tag_names if t not in result.values()]

    @staticmethod
    def _pair_leftover_tags(tags, unassigned, sheet_lines, notes):
        """Tags the normal pairing left out -> nearest bend line that has no tag yet (no distance limit)."""
        pos = {}
        for text, p in notes:
            pos.setdefault(text, p)
        pairs = sorted((point_segment_distance(pos[t], a, b), t, bid)
                       for t in unassigned if t in pos
                       for bid, (a, b) in sheet_lines.items() if bid not in tags)
        paired = []
        for _d, t, bid in pairs:
            if t in unassigned and bid not in tags:
                tags[bid] = t
                unassigned.remove(t)
                paired.append(t)
        return paired

    def _bends_from_tag_lines(self, analysis, to_sheet, tags, unassigned, sheet_lines, notes):
        """A tag of the bend table with no bend line in the view: the straight line (not on the outer boundary)
        nearest to the tag is used as its bend line and measured to the outer edge like any other bend."""
        pos = {}
        for text, p in notes:
            pos.setdefault(text, p)
        edges, outer = analysis["edges"], analysis["outer"]
        known = {r["id"] for r in analysis["results"]}
        used = []
        for t in list(unassigned):
            bid = f"tag {t}"
            if bid not in known:
                if t not in pos:
                    continue
                best = None
                for e in edges:
                    if e["id"] in outer or not e.get("is_line", True):
                        continue
                    a, b = to_sheet(e["m0"]), to_sheet(e["m1"])
                    d = point_segment_distance(pos[t], a, b)
                    if best is None or d < best[0]:
                        best = (d, e)
                if best is None or best[0] > 0.03:
                    continue
                e = best[1]
                result = {"id": bid, "status": "none", "value": None, "reason": "solid line next to the tag"}
                found = self._nearest_parallel_edge(e["p0"], e["p1"], edges, outer)
                if found is not None:
                    result.update(status="ok", edge=found[0], value=found[1],
                                  note=f"no bend line for tag {t} - solid line next to the tag measured to the "
                                       f"nearest parallel {found[2]} edge")
                else:
                    corner = self._flange_corner(e["p0"], e["p1"], edges, outer)
                    if corner is None:
                        continue
                    result.update(status="ok", edge=corner[0], vertex=corner[1], value=corner[2],
                                  note=f"no bend line for tag {t} - solid line next to the tag measured to the "
                                       "outer corner of the flange")
                analysis["bends"].append({"id": bid, "seg": e["entity"], "p0": e["p0"], "p1": e["p1"],
                                          "z": (e["f0"][2], e["f1"][2]), "from_edge": True})
                analysis["results"].append(result)
                known.add(bid)
                sheet_lines[bid] = (to_sheet(e["m0"]), to_sheet(e["m1"]))
            if bid not in sheet_lines:
                continue
            tags[bid] = t
            unassigned.remove(t)
            used.append(t)
        return used

    def _plan_bend_dims(self, drawing, view, analysis, table_top, steps, folded=(), final=True):
        """Map the measured bends to the sheet, pair bend tags, and plan text positions.
        Returns the plan dict or None (reason logged)."""
        if analysis.get("error") or not analysis.get("results"):
            return None
        try:
            edges = analysis["edges"]
            to_sheet = self._sheet_mapper(view, steps, points=[e["m0"] for e in edges] + [e["m1"] for e in edges])
            frame_to_model = analysis["model_to_frame"].inverse()
            bends = {r["id"]: r for r in analysis["bends"] if "error" not in r}

            def bend_sheet(bid):
                b = bends[bid]
                return (to_sheet(frame_to_model.apply((b["p0"][0], b["p0"][1], b["z"][0]))),
                        to_sheet(frame_to_model.apply((b["p1"][0], b["p1"][1], b["z"][1]))))

            sheet_lines = {bid: bend_sheet(bid) for bid in bends}
            tag_names = list(self._bend_tags)
            notes = self._view_notes(view)
            tags, unassigned, ambiguous = assign_tags(sheet_lines, notes, tag_names,
                                                      self.bend_tag_max_distance_mm / 1000.0)
            tags, unassigned = dict(tags), list(unassigned or [])
            if ambiguous or unassigned:
                # tags midway between bend lines: pair ALL tags at once with the smallest total distance
                tags, unassigned = self._assign_tags_global(sheet_lines, notes, tag_names)
                if final:
                    steps.append(f"tags paired together by smallest total distance (ambiguous before: "
                                 f"{sorted(ambiguous.values()) if isinstance(ambiguous, dict) else ambiguous})")
                ambiguous = {}
            late = self._pair_leftover_tags(tags, unassigned, sheet_lines, notes)
            solid = self._bends_from_tag_lines(analysis, to_sheet, tags, unassigned, sheet_lines, notes)
            if final and (late or solid):
                steps.append("tag completion: "
                             + (f"paired to the nearest free bend line {late}; " if late else "")
                             + (f"no bend line - solid line next to the tag used {solid}" if solid else ""))

            mm = 0.001
            width, height = self._sheet_size(drawing)
            outline = self._get(view, "GetOutline")
            off, th, tw = self.dimension_offset_mm * mm, self.bend_dim_text_height_mm * mm, self.bend_dim_text_width_mm * mm
            note_boxes = [(b[0] - 0.5 * mm, b[1] - 0.5 * mm, b[2] + 0.5 * mm, b[3] + 0.5 * mm)
                          for b in (it["box"] for it in self._view_annotation_items(view) if it["kind"] == "note")]
            obstacles = note_boxes or [(x - 2.5 * mm, y - 2.5 * mm, x + 2.5 * mm, y + 2.5 * mm)
                                       for text, (x, y) in notes if text in tag_names or not tag_names]
            if outline and self.add_dimensions:  # overall length (above) and width (left) of the flat view
                obstacles.append((outline[0], outline[3], outline[2], outline[3] + off + th))
                obstacles.append((outline[0] - off - tw / 2.0, outline[1], outline[0], outline[3]))
            dim_h, dim_v = self._dim_room()
            for _name, folded_view in folded:
                o = self._get(folded_view, "GetOutline")
                if o:
                    obstacles.append((o[0], o[1], o[2] + dim_v, o[3] + dim_h))

            outer_ids = analysis["outer"] or set(range(len(edges)))
            part_points = [to_sheet(edges[i][k]) for i in outer_ids for k in ("m0", "m1")]
            ok = [r for r in analysis["results"] if r["status"] == "ok"]
            if tag_names:                                 # only bends that have a tag in the bend table
                untagged = [r["id"] for r in ok if r["id"] not in tags]
                ok = [r for r in ok if r["id"] in tags]
                if final and untagged:
                    steps.append(f"bends without a tag are not dimensioned: {untagged}")
            ok = sorted(ok, key=order_key(self.bend_dim_order))
            dims = []
            for r in ok:
                b0, b1 = sheet_lines[r["id"]]
                e = edges[r["edge"]]
                if r.get("vertex") is not None:          # corner reference: short line through the corner
                    q = to_sheet(e["m0"] if r["vertex"] == 0 else e["m1"])
                    bl = math.hypot(b1[0] - b0[0], b1[1] - b0[1]) or 1.0
                    e0, e1 = q, (q[0] + (b1[0] - b0[0]) / bl * mm, q[1] + (b1[1] - b0[1]) / bl * mm)
                else:
                    e0, e1 = to_sheet(e["m0"]), to_sheet(e["m1"])
                dims.append({"key": r["id"], "b0": b0, "b1": b1, "e0": e0, "e1": e1})
            bounds = (self.layout["margin_mm"] * mm, table_top, width - self.layout["margin_mm"] * mm,
                      height - self.layout["margin_mm"] * mm) if final else None
            placements, needs = plan_placements(
                dims, part_points, self.bend_dim_offset_mm * mm, self.bend_dim_step_mm * mm, tw, th,
                max_levels=self.bend_dim_max_levels, bounds=bounds, obstacles=obstacles) if dims else ([], dict.fromkeys(SIDES, 0.0))
            if final:
                steps.append(f"bend view transform: {to_sheet.status}")
                if tag_names or notes:
                    steps.append(f"bend tags paired: {len(tags)} of {len(tag_names) or len(notes)}"
                                 + (f" (unpaired: {unassigned})" if unassigned else "")
                                 + (f" - AMBIGUOUS (tag is as close to another bend line, check manually): "
                                    f"{ambiguous}" if ambiguous else ""))
            return {"placements": placements, "needs": needs, "tags": tags, "unassigned_tags": unassigned,
                    "ambiguous_tags": ambiguous,
                    "sheet_lines": sheet_lines, "transform_status": to_sheet.status,
                    "edge_lines": {d["key"]: (d["e0"], d["e1"]) for d in dims}}
        except Exception as exc:
            steps.append(f"bend dimension placement failed: {type(exc).__name__}: {exc}")
            return None

    # ---------------- bend dimensions: create + verify ----------------
    def _dimension_value(self, display_dimension):
        dimension = None
        try:
            dimension = display_dimension.GetDimension2(0)
        except Exception:
            dimension = self._get(display_dimension, "GetDimension")
        value = self._get(dimension, "SystemValue")
        try:
            return float(value)
        except Exception:
            return None

    def _selected_count(self, drawing):
        try:
            return int(drawing.SelectionManager.GetSelectedObjectCount2(-1))
        except Exception:
            return None

    def _selected_object(self, drawing, index):
        try:
            return drawing.SelectionManager.GetSelectedObject6(index, -1)
        except Exception:
            return None

    @staticmethod
    def _same_pair(a, b, tol=2e-5):
        """True if the two 3D point pairs are the same segment (either direction)."""
        if not a or not b or None in a or None in b:
            return False

        def close(p, q):
            return all(abs(p[i] - q[i]) <= tol for i in range(3))
        return (close(a[0], b[0]) and close(a[1], b[1])) or (close(a[0], b[1]) and close(a[1], b[0]))

    def _is_bend_line(self, obj, bend_xyz):
        try:
            return self._same_pair(self._segment_xyz(obj), bend_xyz)
        except Exception:
            return False

    def _is_edge(self, obj, edge):
        try:
            return self._same_pair(self._edge_points(obj), (edge["m0"], edge["m1"]))
        except Exception:
            return False

    def _select_pair_by_entities(self, drawing, view, bend_seg, edge_entity, sel):
        drawing.ClearSelection2(True)
        ok1, _ = self._select_entity(view, bend_seg, False, sel)
        ok2, _ = self._select_entity(view, edge_entity, True, sel)
        count = self._selected_count(drawing)
        return ok1 and ok2 and (count is None or count >= 2)

    def _dimension_loose(self, drawing, bend, bend_line, edge_line, corner, func, px, py):
        """Pick the bend line (checked if possible) and the reference edge / corner at its position (not checked)
        and create exactly ONE dimension. Returns the display dimension or None. Nothing is left behind on failure."""
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        ext = drawing.Extension

        def at(line, f):
            return (line[0][0] + f * (line[1][0] - line[0][0]), line[0][1] + f * (line[1][1] - line[0][1]))

        def pick(x, y, types, append):
            for t in types:
                try:
                    if ext.SelectByID2("", t, x, y, 0.0, append, 0, nothing, 0):
                        return True
                except Exception:
                    pass
            return False

        try:
            bend_xyz = self._segment_xyz(bend["seg"])
        except Exception:
            bend_xyz = None
        bend_pt = None
        for f in (0.4, 0.6, 0.25, 0.75):                 # prefer a pick that really is this bend line
            drawing.ClearSelection2(True)
            x, y = at(bend_line, f)
            if pick(x, y, ("EXTSKETCHSEGMENT", "SKETCHSEGMENT"), False) and (
                    bend_xyz is None or self._is_bend_line(self._selected_object(drawing, 1), bend_xyz)):
                bend_pt = (x, y)
                break
        if bend_pt is None:
            bend_pt = at(bend_line, 0.4)
        ref_pts = [edge_line[0]] if corner else [at(edge_line, f) for f in (0.5, 0.3, 0.7)]
        ref_types = ("VERTEX", "EDGE") if corner else ("EDGE", "SILHOUETTE")
        for f in [func] + (["AddDimension2"] if func != "AddDimension2" else []):
            for ex, ey in ref_pts:
                drawing.ClearSelection2(True)
                if not pick(bend_pt[0], bend_pt[1], ("EXTSKETCHSEGMENT", "SKETCHSEGMENT", ""), False):
                    break
                if not pick(ex, ey, ref_types, True):
                    continue
                try:
                    display = getattr(drawing, f)(px, py, 0.0)
                except Exception:
                    display = None
                if display is not None:
                    drawing.ClearSelection2(True)
                    return display
        drawing.ClearSelection2(True)
        return None

    def _select_bend_and_corner(self, drawing, view, bend_seg, edge_entity, which, sel):
        """Bend line + one end point (vertex) of the reference edge: a perpendicular dimension then measures
        exactly the bend-line-to-edge distance."""
        vertex = self._get(edge_entity, "GetStartVertex" if which == 0 else "GetEndVertex")
        if vertex is None:
            return False
        drawing.ClearSelection2(True)
        ok1, _ = self._select_entity(view, bend_seg, False, sel)
        ok2, _ = self._select_vertex(view, vertex, True, sel)
        count = self._selected_count(drawing)
        return ok1 and ok2 and (count is None or count >= 2)

    def _select_pair_by_position(self, drawing, bend, edge, bend_line, edge_line):
        """Pick bend line and edge by sheet position, and keep the pick ONLY if the selected objects really are
        this bend line and this outer edge (checked by their end points). Several points along each are tried."""
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        ext = drawing.Extension
        try:
            bend_xyz = self._segment_xyz(bend["seg"])
        except Exception:
            bend_xyz = None
        fractions = (0.5, 0.3, 0.7, 0.15, 0.85)

        def at(line, f):
            return (line[0][0] + f * (line[1][0] - line[0][0]), line[0][1] + f * (line[1][1] - line[0][1]))

        def pick(x, y, types, append):
            for t in types:
                try:
                    if ext.SelectByID2("", t, x, y, 0.0, append, 0, nothing, 0):
                        return True
                except Exception:
                    pass
            return False

        for fb in fractions:
            bx, by = at(bend_line, fb)
            drawing.ClearSelection2(True)
            if not pick(bx, by, ("EXTSKETCHSEGMENT", "SKETCHSEGMENT", ""), False):
                continue
            if bend_xyz is not None and not self._is_bend_line(self._selected_object(drawing, 1), bend_xyz):
                continue
            for fe in fractions:
                ex, ey = at(edge_line, fe)
                if not pick(ex, ey, ("EDGE", "SILHOUETTE", ""), True):
                    continue
                count = self._selected_count(drawing) or 2
                if count >= 2 and self._is_edge(self._selected_object(drawing, count), edge):
                    return True
                drawing.ClearSelection2(True)
                pick(bx, by, ("EXTSKETCHSEGMENT", "SKETCHSEGMENT", ""), False)
            break
        drawing.ClearSelection2(True)
        return False

    @staticmethod
    def _perpendicular_dim_function(bend_line):
        """Dimension type perpendicular to the bend line, so it can only measure the bend-to-edge distance."""
        (x0, y0), (x1, y1) = bend_line
        dx, dy = abs(x1 - x0), abs(y1 - y0)
        length = math.hypot(dx, dy)
        if length < 1e-9:
            return "AddDimension2"
        if dy <= 0.01 * length:
            return "AddVerticalDimension2"
        if dx <= 0.01 * length:
            return "AddHorizontalDimension2"
        return "AddDimension2"

    def _override_dimension(self, display, value):
        """Show the computed bend dimension value on the dimension. Returns how it was done, or None."""
        text = f"{value * 1000:.2f}".rstrip("0").rstrip(".")
        try:
            display.SetOverride(True, float(value))
            return f"value overridden to {text} mm"
        except Exception:
            pass
        try:
            display.SetText(1, text)
            return f"text set to {text}"
        except Exception:
            return None

    def _create_bend_dims(self, drawing, view, analysis, plan, steps):
        """Width dimension of every bend (bend line -> nearest outer edge). The value written on the sheet is
        ALWAYS the computed bend dimension: if SolidWorks measures something else, the dimension is redone with
        a verified selection, and as a last resort its text is overridden with the computed value."""
        results = {r["id"]: r for r in analysis["results"]}
        bends = {r["id"]: r for r in analysis["bends"]}
        sel = self._select_data(drawing, view)
        created, mismatched = 0, []
        for p in plan["placements"]:
            r = results[p["key"]]
            edge = analysis["edges"][r["edge"]]
            bend = bends[r["id"]]
            tag = plan["tags"].get(r["id"], "?")
            expected = r["value"]
            label = f"bend {r['id']} [tag {tag}] width {expected * 1000:.2f} mm"
            bend_line = plan["sheet_lines"].get(r["id"])
            edge_line = plan.get("edge_lines", {}).get(r["id"])
            px, py = p["text_point"]

            func = self._perpendicular_dim_function(bend_line) if bend_line else "AddDimension2"
            funcs = [func] + (["AddDimension2"] if func != "AddDimension2" else [])
            if r.get("vertex") is not None:              # reference is an outer corner
                attempts = [(f"corner{r['vertex']}", f) for f in funcs]
            else:
                attempts = [("entities", f) for f in funcs]
                if bend_line and edge_line:
                    attempts += [("position", f) for f in funcs]
                attempts += [("corner0", func), ("corner1", func)]   # bend line -> end point of the outer edge

            r["created"], tried = False, []
            for i, (how, f) in enumerate(attempts):
                last = i == len(attempts) - 1
                try:
                    if how == "entities":
                        ok = self._select_pair_by_entities(drawing, view, bend["seg"], edge["entity"], sel)
                    elif how.startswith("corner"):
                        ok = self._select_bend_and_corner(drawing, view, bend["seg"], edge["entity"],
                                                          int(how[-1]), sel)
                    else:
                        ok = self._select_pair_by_position(drawing, bend, edge, bend_line, edge_line)
                    if not ok:
                        tried.append(f"{how}: correct bend line + edge not selected")
                        if not last:
                            continue
                        # nothing verified could be selected: dimension what can be selected, value overridden below
                        drawing.ClearSelection2(True)
                        self._select_entity(view, bend["seg"], False, sel)
                        self._select_entity(view, edge["entity"], True, sel)
                        if (self._selected_count(drawing) or 0) < 2:
                            continue
                    display = getattr(drawing, f)(px, py, 0.0)
                except Exception as exc:
                    tried.append(f"{how}/{f}: {exc}")
                    continue
                if display is None:
                    tried.append(f"{how}/{f}: no dimension created")
                    continue
                actual = self._dimension_value(display)
                if actual is None or abs(actual - expected) <= DIM_VALUE_TOL:
                    r["created"], r["solidworks_value"] = True, actual
                    r["display"], r["side"] = display, p["side"]
                    created += 1
                    check = "verified" if actual is not None else "value not readable"
                    steps.append(f"{label} -> {p['side']} level {p['level']}, {check} ({how}/{f})"
                                 + (f", {p['collisions']} unavoidable overlap(s)" if p["collisions"] else ""))
                    break
                if not last and self._delete_annotation(drawing, display):
                    tried.append(f"{how}/{f}: measured {actual * 1000:.2f} mm -> removed, retrying")
                    continue
                how_fixed = self._override_dimension(display, expected)
                r["created"], r["solidworks_value"] = True, actual
                r["display"], r["side"] = display, p["side"]
                if how_fixed:
                    created += 1
                    r["overridden"] = True
                    steps.append(f"{label} -> SolidWorks measured {actual * 1000:.2f} mm, {how_fixed} "
                                 f"(computed bend dimension used; check the extension lines). " + "; ".join(tried))
                else:
                    mismatched.append(r["id"])
                    steps.append(f"{label}: MISMATCH {actual * 1000:.2f} mm and value could not be overridden - "
                                 "check manually; " + "; ".join(tried))
                break
            if not r["created"] and bend_line and edge_line:
                # ONE dimension from the bend line to the edge at its position; wrong value -> number overridden
                display = self._dimension_loose(drawing, bend, bend_line, edge_line, r.get("vertex") is not None,
                                                func, px, py)
                if display is None:
                    tried.append("position pick without check: no dimension created")
                else:
                    actual = self._dimension_value(display)
                    r["created"], r["solidworks_value"] = True, actual
                    r["display"], r["side"] = display, p["side"]
                    if actual is not None and abs(actual - expected) <= DIM_VALUE_TOL:
                        created += 1
                        steps.append(f"{label} -> {p['side']} level {p['level']}, verified (position pick)")
                    else:
                        how_fixed = self._override_dimension(display, expected)
                        if how_fixed:
                            created += 1
                            r["overridden"] = True
                            steps.append(f"{label} -> {p['side']} level {p['level']}, {how_fixed} (SolidWorks "
                                         f"measured {'?' if actual is None else f'{actual * 1000:.2f}'} mm; "
                                         "check the extension lines)")
                        else:
                            steps.append(f"{label}: MISMATCH and value could not be overridden - check manually")
            if not r["created"]:
                mismatched.append(r["id"])
                steps.append(f"{label}: NOT dimensioned - " + "; ".join(tried))
        drawing.ClearSelection2(True)
        return created, mismatched

    def _delete_annotation(self, drawing, display):
        try:
            nothing = VARIANT(pythoncom.VT_DISPATCH, None)
            annotation = self._get(display, "GetAnnotation")
            drawing.ClearSelection2(True)
            if annotation is not None and annotation.Select3(False, nothing):
                drawing.Extension.DeleteSelection2(6)
                return True
        except Exception:
            pass
        return False

    def _report_bends(self, analysis, plan, steps):
        """One line per bend that is NOT dimensioned (and why) or was measured in a special way."""
        for r in analysis.get("results", []):
            tag = (plan or {}).get("tags", {}).get(r["id"])
            who = f"bend {r['id']}" + (f" [tag {tag}]" if tag else "")
            if r["status"] == "ok":
                if r.get("note"):
                    steps.append(f"{who}: {r['note']}")
                continue
            if r["status"] in ("merged", "duplicate"):
                steps.append(f"{who}: shares the dimension of bend {r.get('merged_into')} ({r['reason']})")
            else:
                steps.append(f"{who}: NOT dimensioned - {r['status']}: {r['reason']} (manual review)")

    def _add_bend_dimensions(self, drawing, view, scale, limit_bottom, steps, folded=()):
        """Measure, place and create the bend dimensions of a flat pattern view (no relayout)."""
        analysis = self._analyze_bends(view, steps, self._current_item)
        plan = self._plan_bend_dims(drawing, view, analysis, limit_bottom, steps, folded, final=True)
        self._finish_bend_dims(drawing, view, analysis, plan, steps)
        return analysis, plan

    def _finish_bend_dims(self, drawing, view, analysis, plan, steps):
        """Create the planned dimensions, log a summary, report every bend that was not dimensioned and
        check that every tag of the bend table has a dimension."""
        if plan is not None:
            created, mismatched = self._create_bend_dims(drawing, view, analysis, plan, steps)
            ok = [r for r in analysis["results"] if r["status"] == "ok"]
            steps.append(f"bend dimensions: {created}/{len(ok)} created and verified, values "
                         f"{[round(r['value'] * 1000, 2) for r in ok]}, levels "
                         f"{sorted({p['level'] for p in plan['placements']})}"
                         + (f", MISMATCH on {mismatched}" if mismatched else ""))
        elif analysis.get("results") and not analysis.get("error"):
            steps.append("bend dimensions NOT created: placement could not be planned (see above)")
        self._report_bends(analysis, plan, steps)
        if plan is not None and self._bend_tags:
            done = {plan["tags"].get(r["id"]) for r in analysis.get("results", []) if r.get("created")}
            missing = [t for t in self._bend_tags if t not in done]
            why = []
            for t in missing:
                bid = next((b for b, x in plan["tags"].items() if x == t), None)
                if bid is None:
                    why.append(f"{t}: no bend line or solid line found next to the tag")
                else:
                    r = next((x for x in analysis["results"] if x["id"] == bid), {})
                    why.append(f"{t}: bend {bid} " + ("dimension could not be created" if r.get("status") == "ok"
                                                       else f"{r.get('status')} - {r.get('reason')}"))
            steps.append("tags with a dimension: " + (f"all {len(self._bend_tags)}" if not missing else
                         f"{len(self._bend_tags) - len(missing)}/{len(self._bend_tags)} - MISSING: " + "; ".join(why)))

    def _write_bend_probe(self, path, analysis, plan, steps):
        """Diagnostic JSON next to the drawing: frames, bends, edges, results, tags, placements."""
        def clean(value):
            if isinstance(value, dict):
                return {str(k): clean(v) for k, v in value.items() if k not in ("seg", "entity", "display")}
            if isinstance(value, (list, tuple, set)):
                return [clean(v) for v in value]
            if isinstance(value, float):
                return round(value, 9)
            if isinstance(value, (str, int, bool)) or value is None:
                return value
            if isinstance(value, Xform):
                return {"array": [round(v, 12) for v in value.to_array()]}
            return str(value)
        data = {"part": self._current_item.get("name"), "convention": self._convention,
                "analysis": clean(analysis or {}), "plan": clean(plan or {})}
        try:
            Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
            steps.append(f"bend probe written: {Path(path).name}")
        except Exception as exc:
            steps.append(f"bend probe not written: {exc}")

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
            if len(set(headers)) == 1 and cols > 1:
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

    def _shrink_table(self, grid, steps):
        """Scale the bend table (text, column widths, row heights) by bend_table_scale."""
        f = self.bend_table_scale
        if abs(f - 1.0) < 1e-6:
            return
        cols = int(self._get(grid, "ColumnCount") or 0)
        rows = int(self._get(grid, "RowCount") or 0)
        done = []
        cells = 0
        for r in range(rows):
            for c in range(cols):
                try:
                    tf = grid.GetCellTextFormat(r, c)
                    if tf is None:
                        continue
                    tf.CharHeight = float(tf.CharHeight) * f
                    grid.SetCellTextFormat(r, c, False, tf)
                    cells += 1
                except Exception:
                    pass
        done.append(f"cell text {cells}/{rows * cols}")
        try:
            tf = grid.GetTextFormat()
            h = float(tf.CharHeight)
            tf.CharHeight = h * f
            grid.SetTextFormat(False, tf)
            done.append(f"text {h * 1000:.2f} -> {h * f * 1000:.2f} mm")
        except Exception as exc:
            done.append(f"table text not changed ({exc})")

        def resize(get, put, n, label):
            k = 0
            for i in range(n):
                try:
                    v = float(get(i)) * f
                    try:
                        put(i, v, 0)
                    except Exception:
                        put(i, v)
                    k += 1
                except Exception:
                    pass
            done.append(f"{label} {k}/{n}")
        resize(grid.GetColumnWidth, grid.SetColumnWidth, cols, "columns")
        resize(grid.GetRowHeight, grid.SetRowHeight, rows, "rows")
        steps.append(f"bend table scaled x{f:.3f}: " + ", ".join(done))

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
        if view is not None:
            tables = self._get(view, "GetTableAnnotations")
            for table in reversed(list(tables or ())):
                annotation = self._get(table, "GetAnnotation")
                if annotation is not None:
                    return annotation
        return None

    def _move_table(self, grid, left, top, steps, view=None):
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
        width, height = self._sheet_size(drawing)
        table = self._insert_table_once(view, width * 0.1, height * 0.5, steps)
        if table is None:
            steps.append("bend table NOT created")
            return None
        try:
            view.ShowSheetMetalBendNotes = True
            steps.append(f"bend notes (tags) on view: {self._get(view, 'ShowSheetMetalBendNotes')}")
        except Exception as exc:
            steps.append(f"could not show bend tags: {exc}")

        grid = self._table_grid(table, steps)
        self._clean_columns(grid, steps)
        self._shrink_table(grid, steps)

        table_w, table_h = self._table_size(grid, steps)
        left = width - self.bend_table_right_margin_mm / 1000.0 - table_w
        bottom = (self.bend_table_bottom_mm + self.layout["table_gap_mm"]) / 1000.0
        top = bottom + table_h
        rows = max(int(self._get(grid, "RowCount") or 1), 1)
        steps.append(f"bend table size {table_w * 1000:.0f} x {table_h * 1000:.0f} mm")

        if not self._move_table(grid, left, top, steps, view):
            if self._delete_table(drawing, grid, view):
                steps.append("bend table re-inserted at the final position")
                table = self._insert_table_once(view, left, top, steps)
                if table is not None:
                    grid = self._table_grid(table, steps)
                    self._clean_columns(grid, steps)
                    self._shrink_table(grid, steps)
                    self._move_table(grid, left, top, steps, view)
            else:
                steps.append("bend table could not be moved or deleted - check its position manually")
        self._bend_tags = self._read_tags(grid, steps)
        return bottom, top, table_h / rows, list(self._bend_tags)

    def _read_tags(self, grid, steps):
        """Tag texts (A, B, ...) from the bend table's tag column; header row(s) are skipped."""
        try:
            cols = int(self._get(grid, "ColumnCount") or 0)
            rows = int(self._get(grid, "RowCount") or 0)
            header_row, tag_col = 0, 0
            for r in range(min(2, rows)):
                found = next((c for c in range(cols) if "tag" in self._cell(grid, r, c).lower()), None)
                if found is not None:
                    header_row, tag_col = r, found
                    break
            tags = [self._cell(grid, r, tag_col) for r in range(header_row + 1, rows)]
            tags = [t for t in tags if t]
            steps.append(f"bend table tags: {tags}")
            return tags
        except Exception as exc:
            steps.append(f"bend table tags not read: {exc}")
            return []

    # ---------------- output folders / PDF ----------------
    def _thickness_folder(self, item):
        """Folder name for the part's sheet thickness (mm), e.g. '1.5 mm'."""
        try:
            t = float(item.get("thickness"))
            if not t > 0:
                raise ValueError
            return self._safe_name(self.thickness_folder_format.format(thickness=round(t, 3)))
        except Exception:
            return self._safe_name(self.unknown_thickness_folder)

    def output_paths(self, item, slddrw_path):
        """(slddrw path, pdf path) - inside the thickness folder when group_by_thickness is on."""
        p = Path(slddrw_path)
        if self.group_by_thickness:
            folder = self._thickness_folder(item)
            if p.parent.name != folder:
                p = p.parent / folder / p.name
        return p, p.with_suffix(".pdf")

    def _save_pdf(self, drawing, pdf_path, steps):
        """Export the open drawing (all sheets) as PDF next to the SLDDRW."""
        pdf_path = Path(pdf_path)
        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except Exception as exc:
            steps.append(f"PDF NOT saved: old PDF is locked (open in a viewer?) - {exc}")
            return False
        nothing = VARIANT(pythoncom.VT_DISPATCH, None)
        export_data = None
        try:
            export_data = self.sw_app.GetExportFileData(1)          # swExportPdfData
            export_data.SetSheets(1, nothing)                        # swExportData_ExportAllSheets
        except Exception:
            pass
        tried = []
        try:
            errors, warnings = self._byref_i4(), self._byref_i4()
            drawing.Extension.SaveAs(str(pdf_path), SW_SAVE_AS_CURRENT_VERSION, SW_SAVE_SILENT,
                                     export_data if export_data is not None else nothing, errors, warnings)
            if self._written(pdf_path):
                steps.append(f"PDF saved: {pdf_path.name}")
                return True
            tried.append(f"Extension.SaveAs errors={getattr(errors, 'value', errors)}")
        except Exception as exc:
            tried.append(f"Extension.SaveAs: {exc}")
        try:
            drawing.SaveAs3(str(pdf_path), SW_SAVE_AS_CURRENT_VERSION, SW_SAVE_SILENT)
            if self._written(pdf_path):
                steps.append(f"PDF saved: {pdf_path.name} (SaveAs3)")
                return True
            tried.append("SaveAs3: no file written")
        except Exception as exc:
            tried.append(f"SaveAs3: {exc}")
        steps.append("PDF NOT saved - " + "; ".join(tried))
        return False

    # ---------------- public ----------------
    def export_part(self, item, slddrw_path):
        slddrw_path, pdf_path = self.output_paths(item, slddrw_path)
        item["slddrw_path"] = str(slddrw_path)
        item["pdf_path"] = str(pdf_path) if self.export_pdf else None
        ok, message, steps = self._export_once(item, slddrw_path)
        if not ok and self._is_connection_error(message):
            steps.append("connection lost -> reconnecting and retrying once")
            time.sleep(3)
            if self._reconnect():
                ok, message, steps2 = self._export_once(item, slddrw_path)
                steps += steps2
        if not ok:
            for p in (slddrw_path, pdf_path):
                try:
                    Path(p).unlink(missing_ok=True)
                except Exception:
                    pass
        return ok, message, steps

    def _export_once(self, item, slddrw_path):
        steps = []
        slddrw_path = Path(slddrw_path)
        slddrw_path.parent.mkdir(parents=True, exist_ok=True)
        steps.append(f"output folder: {slddrw_path.parent}")
        if slddrw_path.exists():
            try:
                slddrw_path.unlink()
            except Exception as exc:
                return False, f"cannot remove old file: {exc}", steps

        part_path = str(Path(item["path"]).resolve())
        configuration = item.get("configuration")
        self._current_item = item
        self._bend_tags = []
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

            self._apply_sheet_format(drawing, item, steps)
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
                try:
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
                for i, view_name in enumerate(self.folded_view_names):
                    label = str(view_name).strip().lstrip("*").strip().lower() or f"view{i + 1}"
                    v = self._create_model_view(drawing, model_path, view_name,
                                                width / 2, height * (0.75 - 0.1 * i), steps)
                    if v is not None:
                        views.append((label, v))
                steps.append(f"folded views: {[n for n, _ in views]}")
            views.append(("flat", flat_view))

            table_info = None
            table_top = (self.bend_table_bottom_mm + self.layout["table_gap_mm"]) / 1000.0
            if self.insert_bend_table:
                table_info = self._insert_bend_table(drawing, flat_view, steps)
                if table_info:
                    table_top = table_info[1]

            folded = [(n, v) for n, v in views if n != "flat"]
            analysis = plan = None
            if self.add_bend_dimensions:
                # bend values do not depend on the layout, so they are measured once
                analysis = self._analyze_bends(flat_view, steps, item)
            pads = self._initial_flat_pads()
            flat_scale = self._layout_views(drawing, folded, flat_view, table_top, steps, pads)
            if flat_scale and analysis and not analysis.get("error") and analysis.get("results"):
                dim_h, dim_v = self._dim_room()
                for i in range(4):  # give the bend dimensions exactly the room they need, then re-plan
                    trial = self._plan_bend_dims(drawing, flat_view, analysis, table_top, [], folded, final=False)
                    if trial is None:
                        break
                    need = {k: v + 0.002 for k, v in trial["needs"].items()}
                    need["left"] = max(0.0, need.get("left", 0.0) - dim_v)
                    need["top"] = max(0.0, need.get("top", 0.0) - dim_h)
                    for k in SIDES:
                        need.setdefault(k, 0.0)
                    if all(need[k] <= pads[k] + 0.0005 and pads[k] - need[k] <= 0.001 for k in SIDES):
                        break
                    # first rounds: exact room (unused space is given back); last round: only grow
                    pads = dict(need) if i < 2 else {k: max(pads[k], need[k]) for k in SIDES}
                    steps.append("relayout: room for bend dimensions " +
                                 ", ".join(f"{k} {v * 1000:.0f} mm" for k, v in pads.items() if v))
                    flat_scale = self._layout_views(drawing, folded, flat_view, table_top, steps, pads) or flat_scale
                plan = self._plan_bend_dims(drawing, flat_view, analysis, table_top, steps, folded, final=True)
            if flat_scale:
                self._set_sheet_scale(drawing, flat_scale, steps)

            if self.add_dimensions:
                for name, v in views:
                    if name == "flat":
                        self._add_overall_dimensions(drawing, v, steps, "flat", True, True, "left")
                    else:
                        self._add_overall_dimensions(drawing, v, steps, name, True, True, "right")

            if analysis is not None:
                self._finish_bend_dims(drawing, flat_view, analysis, plan, steps)
                if self.bend_debug_json:
                    self._write_bend_probe(slddrw_path.with_name(slddrw_path.stem + "_bend_probe.json"),
                                           analysis, plan, steps)

            if self.add_dimensions or self.add_bend_dimensions:
                self._declutter_flat(drawing, flat_view, analysis, plan, steps)

            if self.add_qty_note:
                self._add_qty_note(drawing, item, steps,
                                   table_bottom=table_info[0] if table_info else None,
                                   row_h=table_info[2] if table_info else None)

            try:
                drawing.ForceRebuild3(False)
            except Exception:
                pass

            err = drawing.SaveAs3(str(slddrw_path), SW_SAVE_AS_CURRENT_VERSION, SW_SAVE_SILENT)
            if not self._written(slddrw_path):
                return False, f"SaveAs3 failed (code {err})", steps
            steps.append(f"SLDDRW saved: {slddrw_path.name}")
            if self.export_pdf and not self._save_pdf(drawing, slddrw_path.with_suffix(".pdf"), steps):
                return True, "ok (PDF not saved - see steps)", steps
            return True, "ok", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
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
