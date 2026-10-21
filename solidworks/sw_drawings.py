import csv
import math
import os
import re
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

from .dwg_exporter import DwgExporter

# ---------------------------------------------------------
# SolidWorks constants
# ---------------------------------------------------------
SW_DOC_PART = 1
SW_DWG_PAPER_A4_VERTICAL = 7  # swDwgPaperA4sizeVertical
SW_DWG_PAPERS_USER_DEFINED = 12  # swDwgPapersUserDefined = custom size, NO sheet format (blank sheet)
ANCHOR_BOTTOM_RIGHT = 4  # bend table anchor corner (bottom right)
FIRST_BEND_TAG = "A"  # first tag shown on the flat-pattern bend lines and in the table

# ---------------------------------------------------------
# Layout, copied from the reference PDF (A4 portrait, metres on the sheet)
# ---------------------------------------------------------
SHEET_W, SHEET_H = 0.210, 0.297
FRAME = (0.008, 0.008, 0.202, 0.289)  # outer frame
TITLE_TOP = 0.027  # top of the title block
TITLE_COLS = (0.008, 0.105, 0.140, 0.171, 0.202)  # DWG NO. | THICKNESS | SCALE | SHEET
TABLE_BOTTOM = 0.033  # bend table: bottom right, just above the title block
ROW_HEIGHT_GUESS = 0.0065
TABLE_WIDTH_GUESS = 0.093

VIEW_ROW_X = (0.016, 0.194)  # the Front / Top / Right row
VIEW_ROW_Y = (0.222, 0.270)  # vertical band of the views in that row
VIEW_LABEL_TOP = 0.216  # "FRONT VIEW  SCALE 1:x" labels, all on one line
VIEW_CELL_GAP = 0.010
VIEW_DIM_ROOM = 0.016  # room right of a folded view for its height dimension

FLAT_X = (0.034, 0.178)  # flat pattern area (side room is for the dimensions)
FLAT_TOP = 0.190
FLAT_BOTTOM_ROOM = 0.020  # between the bend table and the flat pattern

QTY_X, QTY_TOP = 0.0196, 0.0465
MARGIN_GUESS = 0.0084  # a view outline is wider than its geometry by about this much on the sheet
DIM_OFFSET = 0.008  # first dimension line distance from a view
PICK_FRACTIONS = (0.5, 0.35, 0.65, 0.2, 0.8, 0.1, 0.9)  # where along an edge a pick is tried


class SwDrawingPdfExporter(DwgExporter):
    """
    Creates a real SolidWorks drawing for every sheet metal part (default drawing template):
    flat pattern + folded views + bend table + native dimensions + info note,
    saves the .slddrw and exports the PDF.  Every optional step is best effort and logged.
    """

    def __init__(
        self,
        sw_app,
        assembly_path=None,
        reconnect=None,
        template=None,
        paper_size=SW_DWG_PAPER_A4_VERTICAL,
        folded_views=True,
        bend_table=True,
        dimensions=True,
        view_names=("Front", "Top", "Right"),
        qty_note=True,
        blank_sheet=True,
        frame=True,
    ):
        super().__init__(sw_app, assembly_path=assembly_path, reconnect=reconnect)
        self.template = template
        self.paper_size = paper_size
        self.folded_views = folded_views
        self.view_names = tuple(view_names)
        self.qty_note = qty_note
        self.bend_table = bend_table
        self.dimensions = dimensions
        self._template_path = None
        self._select_error = ""
        self.blank_sheet = blank_sheet
        self.frame = frame

    # =====================================================
    # TEMPLATE
    # =====================================================

    def find_template(self):
        if self._template_path:
            return self._template_path

        candidates = []
        if self.template:
            candidates.append(str(self.template))

        # 1) the default drawing template from the user's SolidWorks options
        for index in range(0, 100):
            try:
                value = self.sw_app.GetUserPreferenceStringValue(index)
            except Exception:
                continue
            if value and str(value).lower().endswith(".drwdot"):
                candidates.append(str(value))

        # 2) documented helper
        try:
            value = self.sw_app.GetDocumentTemplate(3, "", self.paper_size, 0.0, 0.0)
            if value:
                candidates.append(str(value))
        except Exception:
            pass

        # 3) standard install folder
        program_data = os.environ.get("ProgramData")
        if program_data:
            base = Path(program_data) / "SolidWorks"
            if base.exists():
                candidates += [
                    str(p)
                    for p in sorted(base.glob("SOLIDWORKS*/templates/*.drwdot"), reverse=True)
                ]

        for candidate in candidates:
            if Path(candidate).exists():
                self._template_path = candidate
                return candidate
        return None

    def _find_bend_table_template(self):
        roots = []
        try:
            executable = self.sw_app.GetExecutablePath()
            if executable:
                roots.append(Path(str(executable)))
        except Exception:
            pass

        program_data = os.environ.get("ProgramData")
        if program_data:
            roots.append(Path(program_data) / "SolidWorks")

        for root in roots:
            try:
                if root.exists():
                    found = sorted(root.glob("**/*.sldbndtbt"), reverse=True)
                    if found:
                        return str(found[0])
            except Exception:
                continue
        return ""

    # =====================================================
    # SMALL SOLIDWORKS HELPERS
    # =====================================================

    @staticmethod
    def _r8_array(*values):
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in values))

    def _sheet_size(self, drawing):
        sheet = self._call(drawing, "GetCurrentSheet")
        props = self._call(sheet, "GetProperties2")
        try:
            width, height = float(props[5]), float(props[6])
            if width > 0 and height > 0:
                return sheet, width, height
        except Exception:
            pass
        return sheet, 0.210, 0.297

    def _outline(self, view):
        values = self._call(view, "GetOutline")
        try:
            x0, y0, x1, y1 = (float(v) for v in values[:4])
            return x0, y0, x1, y1
        except Exception:
            return None

    def _position(self, view):
        values = self._call(view, "Position")
        try:
            return float(values[0]), float(values[1])
        except Exception:
            return None

    def _view_name(self, view):
        for getter in (lambda: view.GetName2(), lambda: view.Name):
            try:
                value = getter()
                if value:
                    return str(value)
            except Exception:
                continue
        return None

    def _move_view_center(self, view, cx, cy):
        outline = self._outline(view)
        position = self._position(view)
        if outline is None or position is None:
            return False

        dx = cx - (outline[0] + outline[2]) / 2
        dy = cy - (outline[1] + outline[3]) / 2
        target = (position[0] + dx, position[1] + dy)

        for value in (self._r8_array(*target), target):
            try:
                view.Position = value
                break
            except Exception:
                continue
        else:
            return False

        after = self._outline(view)
        if after is None:
            return True
        return (
            abs((after[0] + after[2]) / 2 - cx) < 0.002
            and abs((after[1] + after[3]) / 2 - cy) < 0.002
        )

    def _create_view(self, drawing, kind, model_path, configuration, x, y):
        try:
            if kind == "flat":
                view = drawing.CreateFlatPatternViewFromModelView3(
                    str(model_path), str(configuration or ""), x, y, 0.0, False, False
                )
                if view is not None:
                    return view
                return drawing.CreateDrawViewFromModelView3(
                    str(model_path), "*Flat pattern", x, y, 0.0
                )
            return drawing.CreateDrawViewFromModelView3(
                str(model_path), f"*{kind}", x, y, 0.0
            )
        except Exception:
            if kind == "flat":
                try:
                    return drawing.CreateDrawViewFromModelView3(
                        str(model_path), "*Flat pattern", x, y, 0.0
                    )
                except Exception:
                    return None
            return None

    # =====================================================
    # BEND TABLE
    # =====================================================

    @staticmethod
    def _method_signature(obj, method_name):
        """(parameter names, parameter variant types) of a COM method, read from its type library."""
        try:
            typeinfo = obj._oleobj_.GetTypeInfo()
            attr = typeinfo.GetTypeAttr()
            for index in range(attr.cFuncs):
                desc = typeinfo.GetFuncDesc(index)
                names = typeinfo.GetNames(desc.memid)
                if names and names[0].lower() == method_name.lower():
                    args = getattr(desc, "args", None) or desc[2]
                    types = []
                    for arg in args:
                        try:
                            types.append(int(arg[0][0]))
                        except Exception:
                            types.append(None)
                    return list(names[1:]), types
        except Exception:
            pass
        return None, None

    @staticmethod
    def _guess_arguments(names, types, x, y, template):
        values = []
        for position, name in enumerate(names):
            key = str(name).lower()
            vt = types[position] if position < len(types) else None
            if "anchorpoint" in key or "useanchor" in key:
                values.append(False)
            elif key == "x":
                values.append(x)
            elif key == "y":
                values.append(y)
            elif "anchortype" in key or key == "anchor":
                values.append(ANCHOR_BOTTOM_RIGHT)
            elif "template" in key:
                values.append(template)
            elif "startvalue" in key or "starttag" in key or "firsttag" in key:
                values.append(FIRST_BEND_TAG)
            elif vt == 11:  # VT_BOOL
                values.append(False)
            elif vt == 5:  # VT_R8
                values.append(0.0)
            elif vt == 8:  # VT_BSTR
                values.append("")
            else:
                values.append(0)
        return values

    def _insert_bend_table(self, flat_view, sheet_w, steps):
        x = FRAME[2]
        y = TABLE_BOTTOM
        template = self._find_bend_table_template()

        attempts = []

        names, types = self._method_signature(flat_view, "InsertBendTable")
        if names:
            arguments = self._guess_arguments(names, types, x, y, template)
            steps.append(f"bend table signature: InsertBendTable({', '.join(map(str, names))})")
            attempts.append(lambda: flat_view.InsertBendTable(*arguments))
        else:
            steps.append("bend table signature: could not be read from SolidWorks")

        attempts += [
            lambda: flat_view.InsertBendTable(False, x, y, ANCHOR_BOTTOM_RIGHT, FIRST_BEND_TAG, template),
            lambda: flat_view.InsertBendTable(False, x, y, ANCHOR_BOTTOM_RIGHT, FIRST_BEND_TAG, ""),
            lambda: flat_view.InsertBendTable(False, x, y, ANCHOR_BOTTOM_RIGHT, template),
            lambda: flat_view.InsertBendTable(False, x, y, ANCHOR_BOTTOM_RIGHT, ""),
            lambda: flat_view.InsertBendTable(False, ANCHOR_BOTTOM_RIGHT, template),
            lambda: flat_view.InsertBendTable(False, ANCHOR_BOTTOM_RIGHT, ""),
        ]
        errors = []
        for number, attempt in enumerate(attempts, start=1):
            try:
                table = attempt()
                if table is not None:
                    steps.append(f"bend table: inserted (variant {number}; first bend tag={FIRST_BEND_TAG})")
                    return table
                errors.append(f"v{number}: returned nothing")
            except Exception as exc:
                errors.append(f"v{number}: {str(exc)[:70]}")
        steps.append("bend table: FAILED (" + " | ".join(errors) + ")")
        return None

    def _set_bend_table_headers(self, table, steps):
        headers = ("Tag", "Direction", "Angle", "Inner Radius")
        try:
            columns = int(self._get(table, "ColumnCount") or 0)
        except Exception:
            columns = 0
        if columns < len(headers):
            steps.append(f"bend table headers: expected 4 columns, found {columns}")
            return False

        changed = 0
        for column, header in enumerate(headers):
            written = False
            for setter in (
                lambda c=column, h=header: table.SetText(0, c, h),
                lambda c=column, h=header: table.Text(0, c, h),
            ):
                try:
                    setter()
                    written = True
                    break
                except Exception:
                    pass
            if written:
                changed += 1
            else:
                steps.append(f"bend table header not changed: column {column} -> {header}")

        try:
            self._call(table, "Update")
        except Exception:
            pass
        steps.append(f"bend table headers: {changed}/4 updated (Tag / Direction / Angle / Inner Radius)")
        return changed == 4

    def _drop_blank_row(self, table, steps):
        """The default bend table has an empty row under the header; the reference has none."""
        try:
            rows = int(self._get(table, "RowCount"))
            columns = int(self._get(table, "ColumnCount"))
            if rows < 3:
                return
            def row_text(r):
                return "".join(str(self._call(table, "Text", r, c) or "").strip() for c in range(columns))
            if row_text(1) or not row_text(2):
                return
            for args in ((1, True), (1,)):
                try:
                    table.DeleteRow2(*args)
                    break
                except Exception:
                    continue
            after = int(self._get(table, "RowCount"))
            steps.append("bend table: empty row removed" if after < rows else "bend table: empty row could not be removed")
        except Exception as exc:
            steps.append(f"bend table: empty row not removed ({str(exc)[:50]})")

    def _table_size(self, table):
        rows = self._get(table, "RowCount")
        columns = self._get(table, "ColumnCount")
        try:
            rows = int(rows)
        except Exception:
            rows = 0
        try:
            columns = int(columns)
        except Exception:
            columns = 0

        height = 0.0
        for row in range(rows):
            value = self._call(table, "GetRowHeight", row)
            try:
                height += float(value)
            except Exception:
                height += ROW_HEIGHT_GUESS
        if height <= 0:
            height = max(rows, 1) * ROW_HEIGHT_GUESS

        width = 0.0
        for column in range(columns):
            value = self._call(table, "GetColumnWidth", column)
            try:
                width += float(value)
            except Exception:
                width = 0.0
                break
        if width <= 0:
            width = TABLE_WIDTH_GUESS
        return width, height

    def _place_bend_table(self, table, sheet_w, steps):
        """Moves the table to the bottom right corner, just above the title block. Returns its height."""
        width, height = self._table_size(table)
        x_left = FRAME[2] - width
        y_top = TABLE_BOTTOM + height

        try:
            annotation = self._call(table, "GetAnnotation")
            annotation.SetPosition2(x_left, y_top, 0.0)
            steps.append(f"bend table: placed bottom right ({width * 1000:.0f} x {height * 1000:.0f} mm)")
        except Exception as exc:
            steps.append(f"bend table: could not be moved, left where SolidWorks put it ({str(exc)[:60]})")
        return height

    def _describe_table(self, table, steps):
        try:
            rows = int(self._get(table, "RowCount"))
            columns = int(self._get(table, "ColumnCount"))
        except Exception:
            steps.append("bend table content: could not be read")
            return

        steps.append(f"bend table content ({max(rows - 1, 0)} rows after the header):")
        for row in range(min(rows, 40)):
            cells = []
            for column in range(columns):
                value = self._call(table, "Text", row, column)
                cells.append("" if value is None else str(value).strip())
            steps.append("    " + " | ".join(cells))

    # =====================================================
    # SELECTION / GEOMETRY HELPERS
    # =====================================================

    @staticmethod
    def _xyz(value):
        """Read a SolidWorks point (object with X/Y/Z) or an array into an XYZ tuple."""
        if value is None:
            return None
        try:
            return (float(value.X), float(value.Y), float(value.Z))
        except Exception:
            pass
        try:
            values = list(value)
            if len(values) >= 3:
                return (float(values[0]), float(values[1]), float(values[2]))
        except Exception:
            pass
        return None

    def _line_segment_points(self, segment):
        start = self._xyz(self._call(segment, "GetStartPoint2"))
        end = self._xyz(self._call(segment, "GetEndPoint2"))
        if start is None or end is None:
            return None
        return start, end

    def _get_bend_lines(self, view):
        try:
            return list(self._call(view, "GetBendLines") or [])
        except Exception:
            return []

    def _select_at(self, drawing, x, y, append, kinds=("EDGE", "SKETCHSEGMENT")):
        """
        Select an entity at sheet coordinates (metres).
        NOTE: pywin32 needs an explicit VT_DISPATCH null for the Callout argument; a plain
        None raises 'type mismatch', which silently made every earlier pick fail.
        """
        try:
            ext = drawing.Extension
        except Exception as exc:
            self._select_error = f"no Extension: {str(exc)[:60]}"
            return False
        for kind in kinds:
            try:
                ok = ext.SelectByID2(
                    "", kind, float(x), float(y), 0.0, bool(append), 0,
                    VARIANT(pythoncom.VT_DISPATCH, None), 0,
                )
                if ok:
                    return True
            except Exception as exc:
                self._select_error = str(exc)[:80]
        return False

    def _select_object(self, drawing, view, obj, append):
        """Select a COM entity (e.g. a bend line segment) inside a drawing view."""
        try:
            sel_mgr = drawing.SelectionManager
            data = sel_mgr.CreateSelectData()
            data.View = view
            return bool(obj.Select4(bool(append), data))
        except Exception as exc:
            self._select_error = f"Select4: {str(exc)[:70]}"
            return False

    def _clear(self, drawing):
        try:
            drawing.ClearSelection2(True)
        except Exception:
            pass

    def _dim_value(self, dim):
        for getter in (lambda: dim.GetDimension2(0).SystemValue, lambda: dim.GetDimension2(0).Value):
            try:
                return float(getter())
            except Exception:
                continue
        return None

    def _delete_dimension(self, drawing, dim):
        try:
            self._clear(drawing)
            dim.GetAnnotation().Select3(False, VARIANT(pythoncom.VT_DISPATCH, None))
            drawing.Extension.DeleteSelection2(0)
        except Exception as exc:
            self._select_error = f"delete dim: {str(exc)[:50]}"
        self._clear(drawing)

    def _add_checked_dimension(self, drawing, label, expected, tol=None):
        """AddDimension2 for the current selection. Kept only if its value matches `expected` (metres)."""
        try:
            dim = drawing.AddDimension2(float(label[0]), float(label[1]), 0.0)
        except Exception as exc:
            self._select_error = f"AddDimension2: {str(exc)[:60]}"
            self._clear(drawing)
            return False
        if dim is None:
            self._clear(drawing)
            return False
        value = self._dim_value(dim)
        self._clear(drawing)
        if value is not None and expected is not None:
            if abs(value - expected) > (tol if tol is not None else max(0.0015, 0.01 * expected)):
                self._select_error = (
                    f"wrong value {value * 1000:.1f} mm (expected {expected * 1000:.1f} mm) - removed"
                )
                self._delete_dimension(drawing, dim)
                return False
        return True

    def _dimension_two_picks(self, drawing, first_pts, second_pts, label, expected=None):
        """Pick an edge near any of first_pts, another near any of second_pts, add a checked dimension."""
        for p in first_pts:
            self._clear(drawing)
            if not self._select_at(drawing, p[0], p[1], False, kinds=("EDGE",)):
                continue
            for q in second_pts:
                if self._select_at(drawing, q[0], q[1], True, kinds=("EDGE",)):
                    if self._add_checked_dimension(drawing, label, expected):
                        return True
                    break
        self._clear(drawing)
        return False

    # =====================================================
    # FLAT PATTERN DIMENSIONS
    # =====================================================

    # ---------- edge scan: find where the outer edges REALLY are ----------
    def _scan_side(self, drawing, side, bbox, span=0.012, step=0.0005, wanted=3):
        """Walk from outside to inside across a side of the bbox; the first successful pick is a real edge point."""
        x0, y0, x1, y1 = bbox
        hits = []
        for f in (0.5, 0.35, 0.65, 0.2, 0.8):
            if side in ("left", "right"):
                y = y0 + (y1 - y0) * f
                base, sign = (x0, 1.0) if side == "left" else (x1, -1.0)
                points = [(base - sign * span + sign * step * i, y) for i in range(int(2 * span / step) + 1)]
            else:
                x = x0 + (x1 - x0) * f
                base, sign = (y0, 1.0) if side == "bottom" else (y1, -1.0)
                points = [(x, base - sign * span + sign * step * i) for i in range(int(2 * span / step) + 1)]
            for px, py in points:
                if self._select_at(drawing, px, py, False, kinds=("EDGE",)):
                    hits.append((px, py))
                    break
            self._clear(drawing)
            if len(hits) >= wanted:
                break
        return hits

    def _scan_edges(self, drawing, view, bbox, steps, tag):
        sides = {s: self._scan_side(drawing, s, bbox) for s in ("left", "right", "bottom", "top")}
        if not any(sides.values()):
            name = self._view_name(view)
            try:
                if name and drawing.ActivateView(name):
                    sides = {s: self._scan_side(drawing, s, bbox) for s in ("left", "right", "bottom", "top")}
                    steps.append(f"{tag}: edges only pickable with the view activated ({name})")
            except Exception as exc:
                steps.append(f"{tag}: ActivateView failed ({str(exc)[:40]})")
        steps.append(f"{tag}: edge scan hits " + ", ".join(f"{k}={len(v)}" for k, v in sides.items()))
        return sides

    def _dimension_overall(self, drawing, view, bbox, den, steps, tag, sides=None):
        """Overall length (above the view) and height (right of the view)."""
        x0, y0, x1, y1 = bbox
        w, h = x1 - x0, y1 - y0
        if sides is None:
            sides = self._scan_edges(drawing, view, bbox, steps, tag)
        fallback_h = [(x0, y0 + h * f) for f in PICK_FRACTIONS], [(x1, y0 + h * f) for f in PICK_FRACTIONS]
        fallback_v = [(x0 + w * f, y0) for f in PICK_FRACTIONS], [(x0 + w * f, y1) for f in PICK_FRACTIONS]
        made = 0
        if self._dimension_two_picks(
            drawing, sides["left"] or fallback_h[0], sides["right"] or fallback_h[1],
            ((x0 + x1) / 2, y1 + DIM_OFFSET), w * den,
        ):
            made += 1
        else:
            steps.append(f"{tag}: overall length not created ({self._select_error or 'no edge picked'})")
        if self._dimension_two_picks(
            drawing, sides["bottom"] or fallback_v[0], sides["top"] or fallback_v[1],
            (x1 + DIM_OFFSET, (y0 + y1) / 2), h * den,
        ):
            made += 1
        else:
            steps.append(f"{tag}: overall height not created ({self._select_error or 'no edge picked'})")
        return made

    # ---------- bend lines: positions come from the A/B/C tag notes ----------
    def _view_tags(self, view):
        """[(letter, x, y)] of the single-letter notes (bend tags) that SolidWorks put on the flat pattern."""
        found = []

        def add(text, pos):
            text = str(text or "").strip()
            try:
                if len(text) == 1 and text.isalpha() and pos is not None:
                    found.append((text.upper(), float(pos[0]), float(pos[1])))
            except Exception:
                pass

        try:
            ann = view.GetFirstAnnotation3()
            guard = 0
            while ann is not None and guard < 500:
                guard += 1
                try:
                    if int(ann.GetType()) == 6:  # swNote
                        note = ann.GetSpecificAnnotation()
                        try:
                            text = note.GetText()
                        except Exception:
                            text = note.Text
                        add(text, ann.GetPosition())
                except Exception:
                    pass
                ann = ann.GetNext3()
        except Exception:
            pass
        if not found:
            try:
                for note in list(view.GetNotes() or []):
                    try:
                        add(note.GetText(), note.GetAnnotation().GetPosition())
                    except Exception:
                        continue
            except Exception:
                pass
        unique = {}
        for t in found:
            unique.setdefault(t[0], t)
        return sorted(unique.values(), key=lambda t: t[0])

    def _fit_bend_mapping(self, pairs, tags, bbox, den, steps):
        """
        Sketch coordinates (real metres) -> sheet coordinates:  sheet = T + (1/den) * M * p.
        M is one of the 8 rotations/flips; T comes from the tag positions (a tag sits on its bend line).
        """
        n = len(pairs)
        if n == 0 or len(tags) < n:
            steps.append(f"flat dimensions: {len(tags)} tags found for {n} bend lines - cannot locate the bend lines")
            return None
        tags = tags[:n]
        s = 1.0 / float(den)
        mids = [((a[0] + b[0]) / 2, (a[1] + b[1]) / 2) for a, b in pairs]
        x0, y0, x1, y1 = bbox
        mats = {
            "rot0": ((1, 0), (0, 1)), "rot90": ((0, -1), (1, 0)), "rot180": ((-1, 0), (0, -1)),
            "rot270": ((0, 1), (-1, 0)), "flipX": ((-1, 0), (0, 1)), "flipY": ((1, 0), (0, -1)),
            "swap": ((0, 1), (1, 0)), "swap180": ((0, -1), (-1, 0)),
        }

        def apply(M, p):
            return (M[0][0] * p[0] + M[0][1] * p[1], M[1][0] * p[0] + M[1][1] * p[1])

        best = None
        for name, M in mats.items():
            rm = [apply(M, m) for m in mids]
            tx = sum(t[1] for t in tags) / n - s * sum(r[0] for r in rm) / n
            ty = sum(t[2] for t in tags) / n - s * sum(r[1] for r in rm) / n
            residual = sum(math.hypot(t[1] - (tx + s * r[0]), t[2] - (ty + s * r[1])) for t, r in zip(tags, rm)) / n
            outside, fill = 0.0, 0.0
            for a, b in pairs:
                pa, pb = apply(M, a), apply(M, b)
                pts = [(tx + s * pa[0], ty + s * pa[1]), (tx + s * pb[0], ty + s * pb[1])]
                for x, y in pts:
                    outside += max(x0 - x, 0) + max(x - x1, 0) + max(y0 - y, 0) + max(y - y1, 0)
                dx, dy = abs(pts[0][0] - pts[1][0]), abs(pts[0][1] - pts[1][1])
                fill += (dy / max(y1 - y0, 1e-9)) if dy >= dx else (dx / max(x1 - x0, 1e-9))
            fill /= n
            score = residual + outside + 0.003 * (1.0 - min(fill, 1.0))
            if best is None or score < best[0]:
                best = (score, name, M, tx, ty, residual, outside, fill)
        score, name, M, tx, ty, residual, outside, fill = best
        if residual > 0.008 or outside > 0.004 * n:
            steps.append(
                f"flat dimensions: tags do not match the bend lines (best {name}, residual {residual * 1000:.1f} mm, "
                f"outside {outside * 1000:.1f} mm)"
            )
            return None
        steps.append(
            f"flat dimensions: bend line mapping = {name}, tag residual {residual * 1000:.1f} mm, fill {fill:.2f}"
            + (" (single bend: orientation guessed from its length)" if n == 1 else "")
        )

        def mapper(p):
            q = apply(M, p)
            return (tx + s * q[0], ty + s * q[1])

        return mapper

    def _dimension_chain(self, drawing, view, entities, bbox, den, bottom_row):
        """edge -> bend -> bend -> ... -> edge, one dimension between every neighbouring pair."""
        x0, y0, x1, y1 = bbox
        made = 0
        entities = sorted(entities, key=lambda e: e["coord"])
        for e1, e2 in zip(entities, entities[1:]):
            gap = e2["coord"] - e1["coord"]
            if gap < 0.0004:
                continue
            self._clear(drawing)
            if not self._select_entity(drawing, view, e1, False):
                continue
            if not self._select_entity(drawing, view, e2, True):
                self._clear(drawing)
                continue
            mid = (e1["coord"] + e2["coord"]) / 2
            label = (mid, y0 - DIM_OFFSET) if bottom_row else (x0 - DIM_OFFSET, mid)
            exact = e1["kind"] == "bend" and e2["kind"] == "bend"
            tol = None if exact else 0.003 * den  # edge distances depend on the tag position (about 3 mm on the sheet)
            if self._add_checked_dimension(drawing, label, gap * den, tol):
                made += 1
        self._clear(drawing)
        return made

    def _select_entity(self, drawing, view, ent, append):
        if ent["kind"] == "bend":
            if self._select_object(drawing, view, ent["seg"], append):
                return True
            for p in ent["along"]:
                if self._select_at(drawing, p[0], p[1], append, kinds=("SKETCHSEGMENT", "EDGE")):
                    return True
            return False
        for p in ent["picks"]:
            if self._select_at(drawing, p[0], p[1], append, kinds=("EDGE",)):
                return True
        return False

    def _insert_flat_pattern_dimensions(self, drawing, view, bbox, den, steps):
        """
        Like the reference: overall length (top) and width (right) + chain dimensions
        edge | bend | bend | edge along the bottom (vertical bend lines) and the left (horizontal bend lines).
        Needs the bend table to exist already (its A/B/C tags tell where the bend lines are).
        """
        self._select_error = ""
        x0, y0, x1, y1 = bbox
        sides = self._scan_edges(drawing, view, bbox, steps, "flat dimensions")
        created = self._dimension_overall(drawing, view, bbox, den, steps, "flat dimensions", sides)

        segments = []
        for seg in self._get_bend_lines(view):
            pts = self._line_segment_points(seg)
            if pts:
                segments.append((seg, pts))
        tags = self._view_tags(view)
        steps.append(
            f"flat dimensions: bend lines readable {len(segments)}, tags found {''.join(t[0] for t in tags) or 'none'}"
        )

        if segments:
            pairs = [((p[0][0], p[0][1]), (p[1][0], p[1][1])) for _, p in segments]
            mapper = self._fit_bend_mapping(pairs, tags, bbox, den, steps)
            if mapper is not None:
                v_bends, h_bends = [], []
                for (seg, _), (ra, rb) in zip(segments, pairs):
                    a, b = mapper(ra), mapper(rb)
                    along = [(a[0] + (b[0] - a[0]) * f, a[1] + (b[1] - a[1]) * f) for f in PICK_FRACTIONS]
                    if abs(a[0] - b[0]) <= abs(a[1] - b[1]):
                        v_bends.append({"kind": "bend", "seg": seg, "coord": (a[0] + b[0]) / 2, "along": along})
                    else:
                        h_bends.append({"kind": "bend", "seg": seg, "coord": (a[1] + b[1]) / 2, "along": along})
                v_bends = [e for e in v_bends if x0 < e["coord"] < x1]
                h_bends = [e for e in h_bends if y0 < e["coord"] < y1]
                n_v = n_h = 0
                if v_bends:
                    chain = list(v_bends)
                    if sides["left"]:
                        chain.append({"kind": "edge", "coord": x0, "picks": sides["left"]})
                    if sides["right"]:
                        chain.append({"kind": "edge", "coord": x1, "picks": sides["right"]})
                    n_v = self._dimension_chain(drawing, view, chain, bbox, den, bottom_row=True)
                if h_bends:
                    chain = list(h_bends)
                    if sides["bottom"]:
                        chain.append({"kind": "edge", "coord": y0, "picks": sides["bottom"]})
                    if sides["top"]:
                        chain.append({"kind": "edge", "coord": y1, "picks": sides["top"]})
                    n_h = self._dimension_chain(drawing, view, chain, bbox, den, bottom_row=False)
                steps.append(
                    f"flat dimensions: chain along the bottom {n_v}/{len(v_bends) + 1 if v_bends else 0}, "
                    f"along the left {n_h}/{len(h_bends) + 1 if h_bends else 0}"
                )
                created += n_v + n_h
        if self._select_error:
            steps.append(f"flat dimensions: last problem: {self._select_error}")
        steps.append(f"flat dimensions: {created} native dimensions in total")
        return created

    # =====================================================
    # VIEW SCALE / TRUE GEOMETRY
    # =====================================================

    def _set_view_scale(self, view, den):
        """Own (custom) scale for a view, independent of the sheet scale."""
        try:
            view.UseSheetScale = False
        except Exception:
            pass
        try:
            view.ScaleDecimal = 1.0 / float(den)
            return True
        except Exception:
            pass
        try:
            view.ScaleRatio = self._r8_array(1.0, float(den))
            return True
        except Exception:
            return False

    def _geometry_bbox(self, drawing, view, raw_size, den, apply_scale, steps, tag):
        """
        A view outline = geometry + a fixed margin on the sheet (about 8 mm, independent of the scale).
        Measuring at 1:1 (raw_size) and at 1:den separates the two. Returns (bbox, (length, width) in metres).
        """
        outline = self._outline(view)
        if outline is None:
            return None, None
        n = float(den)
        temp = n < 1.5
        measured_n = 2.0 if temp else n
        if temp:
            apply_scale(2.0)
            self._call(drawing, "ForceRebuild3", False)
        measured = self._outline(view)
        if temp:
            apply_scale(n)
            self._call(drawing, "ForceRebuild3", False)
            outline = self._outline(view) or outline
        if measured is None:
            return outline, raw_size
        k = 1.0 - 1.0 / measured_n
        length = (raw_size[0] - (measured[2] - measured[0])) / k
        width = (raw_size[1] - (measured[3] - measured[1])) / k
        if not (0 < length <= raw_size[0] and 0 < width <= raw_size[1]):
            steps.append(f"{tag}: margin calibration failed, using the raw outline")
            return outline, raw_size
        cx, cy = (outline[0] + outline[2]) / 2, (outline[1] + outline[3]) / 2
        hw, hh = length / n / 2, width / n / 2
        steps.append(f"{tag}: {length * 1000:.1f} x {width * 1000:.1f} mm at 1:{n:g}")
        return (cx - hw, cy - hh, cx + hw, cy + hh), (length, width)

    @staticmethod
    def _fit_denominator(raw_w, raw_h, avail_w, avail_h):
        aw, ah = max(avail_w - MARGIN_GUESS, 0.005), max(avail_h - MARGIN_GUESS, 0.005)
        den = max((raw_w - MARGIN_GUESS) / aw, (raw_h - MARGIN_GUESS) / ah, 1.0)
        return math.ceil(den * 10) / 10

    @staticmethod
    def _scale_text(den):
        return "1:" + f"{den:.2f}".rstrip("0").rstrip(".")

    # =====================================================
    # NOTES / FRAME / TITLE BLOCK
    # =====================================================

    def _insert_note(self, drawing, text, x, y, height=None, bold=False):
        """Note with its top-left corner at (x, y). Returns True when it was created."""
        try:
            note = drawing.InsertNote(text)
            if note is None:
                return False
            try:
                note.SetTextPoint(x, y, 0.0)
            except Exception:
                try:
                    note.GetAnnotation().SetPosition2(x, y, 0.0)
                except Exception:
                    pass
            if height or bold:
                try:
                    ann = note.GetAnnotation()
                    fmt = ann.GetTextFormat(0)
                    if height:
                        fmt.CharHeight = float(height)
                    fmt.Bold = bool(bold)
                    ann.SetTextFormat(0, False, fmt)
                except Exception:
                    pass
            return True
        except Exception as exc:
            self._select_error = f"note: {str(exc)[:60]}"
            return False

    def _line(self, drawing, x1, y1, x2, y2):
        try:
            return drawing.CreateLine2(x1, y1, 0.0, x2, y2, 0.0) is not None
        except Exception as exc:
            self._select_error = f"line: {str(exc)[:60]}"
            return False

    def _draw_frame_and_title(self, drawing, part_name, thickness_text, scale_text, steps):
        fx0, fy0, fx1, fy1 = FRAME
        lines = [
            (fx0, fy0, fx1, fy0), (fx1, fy0, fx1, fy1), (fx1, fy1, fx0, fy1), (fx0, fy1, fx0, fy0),
            (fx0, TITLE_TOP, fx1, TITLE_TOP),
        ] + [(x, fy0, x, TITLE_TOP) for x in TITLE_COLS[1:-1]]
        done = sum(1 for ln in lines if self._line(drawing, *ln))
        labels = ("DWG NO. / PART", "THICKNESS", "SCALE", "SHEET")
        values = (part_name, thickness_text, scale_text, "1 OF 1")
        notes = 0
        for i, (label, value) in enumerate(zip(labels, values)):
            x = TITLE_COLS[i] + 0.0025
            notes += self._insert_note(drawing, label, x, TITLE_TOP - 0.0015, height=0.0022)
            size = 0.0045 if (i == 0 and len(value) > 24) else 0.0055
            notes += self._insert_note(drawing, value, x, TITLE_TOP - 0.0075, height=size, bold=True)
        steps.append(f"frame + title block: {done}/{len(lines)} lines, {notes}/8 texts")

    def _set_properties(self, drawing, props, steps):
        try:
            manager = drawing.Extension.CustomPropertyManager("")
            for key, value in props.items():
                manager.Add3(key, 30, str(value), 2)  # 30 = text, 2 = replace existing
            steps.append("custom properties: " + ", ".join(props))
        except Exception as exc:
            steps.append(f"custom properties: not written ({str(exc)[:50]})")

    # =====================================================
    # ONE PART
    # =====================================================

    def _new_drawing(self, template, steps):
        if self.blank_sheet:
            try:
                drawing = self.sw_app.NewDocument(template, SW_DWG_PAPERS_USER_DEFINED, SHEET_W, SHEET_H)
                if drawing is not None:
                    steps.append("sheet: blank custom sheet (no sheet format)")
                    return drawing
            except Exception as exc:
                steps.append(f"blank sheet failed ({str(exc)[:50]}), using A4")
        return self.sw_app.NewDocument(template, self.paper_size, 0.0, 0.0)

    def _build_drawing(self, item, slddrw_path, pdf_path):
        steps = []
        path = Path(item["path"]).resolve()
        configuration = item.get("configuration")
        model = None
        opened_by_us = False
        previous_config = None
        drawing = None
        drawing_title = None

        try:
            template = self.find_template()
            if not template:
                return False, "no drawing template found (set SW_DRAWING_TEMPLATE in main.py)", steps
            steps.append(f"template: {Path(template).name}")

            steps.append(f"part file: {path} (exists={path.exists()})")
            model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                time.sleep(2)
                model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                return False, f"cannot open part: {open_status}", steps

            _, previous_config = self._activate_configuration(model, configuration)

            drawing = self._new_drawing(template, steps)
            if drawing is None:
                return False, "NewDocument returned nothing", steps
            drawing_title = self._get(drawing, "GetTitle")

            sheet, sheet_w, sheet_h = self._sheet_size(drawing)
            steps.append(f"sheet: {sheet_w * 1000:.0f} x {sheet_h * 1000:.0f} mm")
            self._call(sheet, "SetScale", 1.0, 1.0, True, True)

            # ---- 1) views at 1:1 to learn their sizes ----
            kinds = ["flat"] + (list(self.view_names) if self.folded_views else [])
            views, raw = {}, {}
            for kind in kinds:
                view = self._create_view(drawing, kind, path, configuration, 0.1, 0.15)
                outline = self._outline(view) if view is not None else None
                if outline is None:
                    steps.append(f"view {kind}: FAILED")
                    continue
                views[kind] = view
                raw[kind] = (outline[2] - outline[0], outline[3] - outline[1])
            steps.append("views: " + ", ".join(views) if views else "views: none")
            if "flat" not in views:
                return False, "flat pattern view could not be created", steps
            folded = [k for k in views if k != "flat"]

            # ---- 2) scales ----
            bend_count = len(self._get_bend_lines(views["flat"]))
            table_estimate = (bend_count + 3) * ROW_HEIGHT_GUESS if self.bend_table else 0.0
            flat_bottom = TABLE_BOTTOM + table_estimate + FLAT_BOTTOM_ROOM
            flat_den = self._fit_denominator(
                raw["flat"][0], raw["flat"][1], FLAT_X[1] - FLAT_X[0], FLAT_TOP - flat_bottom
            )

            cell_w = (VIEW_ROW_X[1] - VIEW_ROW_X[0] - VIEW_CELL_GAP * 2) / 3
            view_den = 1.0
            for kind in folded:
                view_den = max(view_den, self._fit_denominator(
                    raw[kind][0], raw[kind][1], cell_w - VIEW_DIM_ROOM, VIEW_ROW_Y[1] - VIEW_ROW_Y[0]
                ))
            steps.append(f"scales: flat {self._scale_text(flat_den)}, folded views {self._scale_text(view_den)}")

            for kind in folded:  # own scale first, so they do not follow the sheet scale
                if not self._set_view_scale(views[kind], view_den):
                    steps.append(f"view {kind}: custom scale could not be set")
            self._call(sheet, "SetScale", 1.0, float(flat_den), True, True)

            # ---- 3) positions ----
            fx = (FLAT_X[0] + FLAT_X[1]) / 2
            fy = (FLAT_TOP + flat_bottom) / 2
            if not self._move_view_center(views["flat"], fx, fy):
                return False, "could not move the flat view", steps
            for i, kind in enumerate(folded):
                cx = VIEW_ROW_X[0] + cell_w * (i + 0.5) + VIEW_CELL_GAP * i - VIEW_DIM_ROOM / 4
                if not self._move_view_center(views[kind], cx, sum(VIEW_ROW_Y) / 2):
                    steps.append(f"view {kind}: could not be moved")
            self._call(drawing, "ForceRebuild3", False)

            flat_bbox, flat_size = self._geometry_bbox(
                drawing, views["flat"], raw["flat"], flat_den,
                lambda d: self._call(sheet, "SetScale", 1.0, float(d), True, True), steps, "flat geometry",
            )
            if flat_bbox is None:
                return False, "flat view outline unavailable", steps
            folded_bbox = {}
            for kind in folded:
                bbox, _ = self._geometry_bbox(
                    drawing, views[kind], raw[kind], view_den,
                    lambda d, v=views[kind]: self._set_view_scale(v, d), steps, f"{kind} geometry",
                )
                if bbox is not None:
                    folded_bbox[kind] = bbox

            # ---- 4) bend table first: it creates the A/B/C tags that tell where the bend lines are ----
            if self.bend_table:
                table = self._insert_bend_table(views["flat"], sheet_w, steps)
                if table is not None:
                    self._set_bend_table_headers(table, steps)
                    self._drop_blank_row(table, steps)
                    table_height = self._place_bend_table(table, sheet_w, steps)
                    if table_height > table_estimate + 0.003:
                        steps.append(
                            f"WARNING: table is {table_height * 1000:.0f} mm tall, only "
                            f"{table_estimate * 1000:.0f} mm were reserved - it may touch the flat pattern"
                        )
                    self._describe_table(table, steps)
                else:
                    steps.append("BEND TABLE MISSING in this drawing")

            # ---- 5) native dimensions (the window must show the whole sheet or picks miss their target) ----
            self._call(drawing, "ViewZoomtofit2")
            self._call(drawing, "GraphicsRedraw2")
            if self.dimensions:
                self._insert_flat_pattern_dimensions(drawing, views["flat"], flat_bbox, flat_den, steps)
                for kind, bbox in folded_bbox.items():
                    self._select_error = ""
                    n = self._dimension_overall(drawing, views[kind], bbox, view_den, steps, f"{kind} dimensions")
                    steps.append(f"{kind} dimensions: {n}/2 created")
                try:  # leave any activated view
                    drawing.ActivateSheet(str(self._call(sheet, "GetName") or ""))
                except Exception:
                    pass
                self._clear(drawing)

            # ---- 6) notes: view labels, Fixed Face, QTY, frame + title block ----
            self._select_error = ""
            label_names = {"Front": "FRONT", "Top": "TOP", "Right": "RIGHT"}
            for kind in folded:
                bbox = folded_bbox.get(kind)
                if bbox is None:
                    continue
                text = f"{label_names.get(kind, kind.upper())} VIEW   SCALE {self._scale_text(view_den)}"
                self._insert_note(drawing, text, bbox[0], VIEW_LABEL_TOP, height=0.0025)

            # "Fixed Face" is written by SolidWorks itself on the flat pattern view.

            quantity = int(item.get("quantity", 1))
            mirror = int(item.get("mirror_quantity", 0) or 0)
            qty_text = f"QTY={quantity}" + (f" ({quantity - mirror} + {mirror} MIRROR)" if mirror else "")
            if self.qty_note:
                self._insert_note(drawing, qty_text, QTY_X, QTY_TOP, height=0.0065)

            part_name = str(item.get("name") or path.stem)
            thickness = item.get("thickness")
            thickness_text = f"{float(thickness):g} mm" if thickness is not None else "N/A"
            scale_text = self._scale_text(flat_den)
            if self.frame:
                self._draw_frame_and_title(drawing, part_name, thickness_text, scale_text, steps)
            self._set_properties(drawing, {
                "DWG NO": part_name, "THICKNESS": thickness_text, "SCALE": scale_text,
                "QTY": quantity, "SHEET": "1 OF 1",
            }, steps)
            if self._select_error:
                steps.append(f"notes/lines: last problem: {self._select_error}")

            # ---- 7) save the drawing, then the PDF ----
            self._call(drawing, "ViewZoomtofit2")
            slddrw_path.parent.mkdir(parents=True, exist_ok=True)
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            for old in (slddrw_path, pdf_path):
                if old.exists():
                    try:
                        old.unlink()
                    except Exception:
                        pass

            try:
                drawing.SaveAs3(str(slddrw_path), 0, 1)
            except Exception as exc:
                steps.append(f"save .slddrw: FAILED ({str(exc)[:60]})")
            else:
                steps.append("saved .slddrw" if self._written(slddrw_path) else "save .slddrw: no file")

            saved_pdf = False
            for attempt in (
                lambda: drawing.SaveAs3(str(pdf_path), 0, 1 | 2),
                lambda: drawing.Extension.SaveAs3(
                    str(pdf_path), 0, 1 | 2,
                    VARIANT(pythoncom.VT_DISPATCH, None), VARIANT(pythoncom.VT_DISPATCH, None),
                    self._byref_i4(), self._byref_i4(),
                ),
            ):
                try:
                    attempt()
                except Exception as exc:
                    steps.append(f"pdf save attempt: {str(exc)[:60]}")
                if self._written(pdf_path):
                    saved_pdf = True
                    break

            if not saved_pdf:
                return False, "PDF was not written", steps
            return True, "SolidWorks drawing", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
            try:
                if drawing_title:
                    self.sw_app.CloseDoc(str(drawing_title))
            except Exception:
                pass
            if model is not None:
                self._restore_configuration(model, previous_config)
            self._close_part(model, opened_by_us)

    def export_part(self, item, slddrw_path, pdf_path):
        ok, message, steps = self._build_drawing(item, Path(slddrw_path), Path(pdf_path))

        if not ok and self._is_connection_error(RuntimeError(message + " " + " ".join(steps))):
            print("  SolidWorks connection problem. Reconnecting and retrying once...")
            time.sleep(3)
            if self._reconnect():
                ok, message, steps = self._build_drawing(item, Path(slddrw_path), Path(pdf_path))

        return ok, message, steps

    # =====================================================
    # ALL PARTS
    # =====================================================

    @staticmethod
    def _safe_name(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    def export_all(self, exported_parts, output_folder, fallback=None, limit=None):
        """
        exported_parts: [{"item": {...}, "dxf": Path}] as returned by ThicknessDwgBuilder.
        PDFs go to <output_folder>/<thickness>mm/<part>.pdf, editable drawings to
        <output_folder>/_drawings/<thickness>mm/<part>.slddrw.
        fallback: a PdfDrawingExporter used when the SolidWorks drawing fails.
        """
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        entries = sorted(
            exported_parts,
            key=lambda e: (e["item"].get("thickness") or 0, e["item"]["name"]),
        )
        if limit:
            entries = entries[: int(limit)]
        total = len(entries)

        print()
        print("=" * 60)
        print("PDF DRAWINGS (SolidWorks drawings, one folder per thickness)")
        print("=" * 60)
        template = self.find_template()
        print(f"Drawing template: {template or 'NOT FOUND'}")
        if limit:
            print(f"TEST MODE: only the first {total} parts")

        rows, used = [], set()
        sw_count = fallback_count = 0
        failed = []

        for index, entry in enumerate(entries, start=1):
            item, dxf_path = entry["item"], entry["dxf"]
            name = item["name"]
            thickness = item.get("thickness")
            folder_name = "unknown_thickness" if thickness is None else f"{thickness:g}mm"

            base = self._safe_name(name)
            unique, counter = base, 2
            while f"{folder_name}/{unique}".lower() in used:
                unique, counter = f"{base}_{counter}", counter + 1
            used.add(f"{folder_name}/{unique}".lower())

            pdf_path = output_folder / folder_name / f"{unique}.pdf"
            slddrw_path = output_folder / "_drawings" / folder_name / f"{unique}.slddrw"
            pdf_path.parent.mkdir(parents=True, exist_ok=True)

            label = f"[{index}/{total}] {name}"
            started = time.time()
            ok, message, steps = self.export_part(item, slddrw_path, pdf_path)

            if ok:
                sw_count += 1
                print(f"{label} -> {folder_name}/{pdf_path.name} ({time.time() - started:.0f}s)")
                for line in steps:
                    print(f"      {line}")
                rows.append([name, thickness, item.get("quantity"),
                             f"{folder_name}/{pdf_path.name}", "SolidWorks", " | ".join(steps)])
                continue

            print(f"{label} -> SolidWorks drawing FAILED: {message}")
            for line in steps:
                print(f"      {line}")

            if fallback is not None:
                ok2, message2 = fallback.export_part(item, dxf_path, pdf_path)
                if ok2:
                    fallback_count += 1
                    print(f"      fallback PDF created ({message2})")
                    rows.append([name, thickness, item.get("quantity"),
                                 f"{folder_name}/{pdf_path.name}", "fallback",
                                 f"{message} | " + " | ".join(steps)])
                    continue
                message = f"{message}; fallback: {message2}"

            failed.append((name, message))
            rows.append([name, thickness, item.get("quantity"), "", "FAILED", message])

        report = output_folder / "pdf_report.csv"
        try:
            with open(report, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["Part", "Thickness (mm)", "Quantity", "PDF", "Method", "Details"])
                writer.writerows(rows)
        except Exception as exc:
            print(f"Could not save PDF report: {exc}")
            report = None

        print()
        print(f"PDFs: {sw_count} SolidWorks drawings, {fallback_count} fallback, {len(failed)} failed (of {total})")

        return {
            "created": sw_count + fallback_count,
            "solidworks": sw_count,
            "fallback": fallback_count,
            "failed": len(failed),
            "failed_parts": failed,
            "report_path": report,
        }
