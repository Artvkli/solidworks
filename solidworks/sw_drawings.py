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
SW_DOC_DRAWING = 3
SW_OPEN_SILENT = 1
SW_OPEN_READONLY = 2
SW_NOTE_TYPE = 6  # swNote
ANCHOR_BOTTOM_RIGHT = 4  # bend table anchor corner (bottom right)
FIRST_BEND_TAG = "A"  # first tag shown on the flat-pattern bend lines and in the table
NO_OPTION_SPECIFIED = 0

# ---------------------------------------------------------
# Layout (sheet metres). The sheet itself (size, frame, title block)
# comes from the DEFAULT SolidWorks drawing template, so everything below
# is relative to the real sheet size read from the template.
# ---------------------------------------------------------
SIDE_MARGIN = 0.025  # free space left/right of the frame
TOP_MARGIN = 0.025  # free space at the top
BOTTOM_MARGIN = 0.068  # room reserved for the template's title block (bend table sits above it)
FLAT_DIM_ROOM = 0.025  # room left/right/top of the flat pattern for its dimensions
FLAT_BOTTOM_ROOM = 0.020  # between the bend table and the flat pattern (chain dimensions)
ROW_HEIGHT_GUESS = 0.0065
TABLE_WIDTH_GUESS = 0.093
MARGIN_GUESS = 0.0084  # a view outline is wider than its geometry by about this much on the sheet
PICK_FRACTIONS = (0.5, 0.35, 0.65, 0.2, 0.8, 0.1, 0.9)  # where along a line a pick is tried

# ---------------------------------------------------------
# Dimension layout (sheet metres)
# ---------------------------------------------------------
DIM_OFFSET = 0.008  # first dimension line distance from the geometry
DIM_TIER_STEP = 0.005  # second row for neighbouring small chain dimensions
EDGE_TOL = 1.5e-4  # two coordinates closer than this are "the same line"
MIN_EDGE_LEN = 5e-4  # shorter edges are not used for dimensions


class SwDrawingPdfExporter(DwgExporter):
    """
    Creates a real SolidWorks drawing for every sheet metal part:
    flat pattern + bend table + native (editable) flat pattern dimensions,
    from the DEFAULT drawing template of SolidWorks.
    Saves the .slddrw and exports the PDF from the saved drawing.
    Every optional step is best effort and logged in `steps`.
    """

    
    def __init__(
        self,
        sw_app,
        assembly_path=None,
        reconnect=None,
        template=None,
        paper_size=None,
        bend_table=True,
        dimensions=True,
        bend_template=None,
        pdf_from_saved_file=True,
        edit_bend_table_text=True,
        dim_text_height=0.0025,
        bottom_margin=BOTTOM_MARGIN,
        sheet_format=None,        # Custom .slddrt sheet format (optional)
        sheet_width=None,
        sheet_height=None,
        table_corner="bottom_right",   # bend table sits at the bottom right, against the frame
        frame_right=0.011,             # distance of the template frame from the right sheet edge (m)
        **legacy_options,
    ):
        super().__init__(
            sw_app,
            assembly_path=assembly_path,
            reconnect=reconnect,
        )

        self.template = template
        self.paper_size = paper_size
        self.bend_table = bend_table
        self.dimensions = dimensions
        self.bend_template = bend_template
        self.pdf_from_saved_file = pdf_from_saved_file
        self.edit_bend_table_text = edit_bend_table_text
        self.dim_text_height = dim_text_height
        self.bottom_margin = bottom_margin
        self.sheet_width = sheet_width
        self.sheet_height = sheet_height
        self.table_corner = table_corner
        self.frame_right = frame_right
        self._current_drawing = None
        # Custom SolidWorks sheet format (.slddrt)
        self.sheet_format = (
            Path(sheet_format).resolve()
            if sheet_format is not None
            else None
        )

        self._template_path = None
        self._constants_cache = None
        self._select_error = ""
    # =====================================================
    # TEMPLATE (default SolidWorks drawing template)
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
            value = self.sw_app.GetDocumentTemplate(3, "", 0, 0.0, 0.0)
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
        if self.bend_template and Path(str(self.bend_template)).exists():
            return str(self.bend_template)

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

    def _new_drawing(self, template, steps):
        """New drawing from the default template, with the sheet of the template."""
        size = self.paper_size if self.paper_size is not None else 0
        try:
            drawing = self.sw_app.NewDocument(template, size, 0.0, 0.0)
            if drawing is not None:
                return drawing
        except Exception as exc:
            steps.append(f"NewDocument failed ({str(exc)[:60]})")
        return None

    # =====================================================
    # SOLIDWORKS CONSTANTS / UNITS
    # =====================================================

    def _constants(self):
        """Enumeration values from the swconst.tlb installed with YOUR SolidWorks."""
        if self._constants_cache is not None:
            return self._constants_cache or None
        self._constants_cache = False

        try:
            import win32com.client as client
            from win32com.client import gencache

            exe = self._get(self.sw_app, "GetExecutablePath")
            if not exe:
                return None
            install = Path(str(exe))
            if install.suffix:
                install = install.parent

            for tlb_path in sorted(install.glob("*.tlb")):
                if "swconst" not in tlb_path.name.lower():
                    continue
                tlb = pythoncom.LoadTypeLib(str(tlb_path))
                guid, lcid, _syskind, major, minor, _flags = tlb.GetLibAttr()
                gencache.EnsureModule(str(guid), lcid, major, minor)
                try:
                    getattr(client.constants, "swUnitSystem_MMGS")
                except AttributeError:
                    continue
                self._constants_cache = client.constants
                return self._constants_cache
        except Exception:
            pass
        return None

    def _ensure_millimeters(self, drawing, steps):
        """A template in INCHES shows a 2 mm radius as 0.08: switch the drawing to MMGS."""
        constants = self._constants()
        if constants is None:
            steps.append(
                "units: swconst.tlb could not be loaded - units NOT changed "
                "(use a metric .drwdot as DRAWING_TEMPLATE_PATH if values look like inches)"
            )
            return False
        try:
            extension = self._call(drawing, "Extension")
            extension.SetUserPreferenceInteger(
                getattr(constants, "swUnitSystem"), NO_OPTION_SPECIFIED, getattr(constants, "swUnitSystem_MMGS")
            )
            extension.SetUserPreferenceInteger(
                getattr(constants, "swUnitsLinear"), NO_OPTION_SPECIFIED, getattr(constants, "swMM")
            )
            current = extension.GetUserPreferenceInteger(getattr(constants, "swUnitsLinear"), NO_OPTION_SPECIFIED)
            ok = current == getattr(constants, "swMM")
            steps.append("units: " + ("millimetres" if ok else f"NOT millimetres (code {current})"))
            return ok
        except Exception as exc:
            steps.append(f"units: could not set millimetres: {str(exc)[:60]}")
            return False

    def _edge_code(self):
        constants = self._constants()
        try:
            return int(getattr(constants, "swViewEntityType_Edge"))
        except Exception:
            return 1

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
        return sheet, 0.297, 0.210

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

    def _create_flat_view(self, drawing, model_path, configuration, x, y):
        try:
            view = drawing.CreateFlatPatternViewFromModelView3(
                str(model_path), str(configuration or ""), x, y, 0.0, False, False
            )
            if view is not None:
                return view
        except Exception:
            pass
        try:
            return drawing.CreateDrawViewFromModelView3(
                str(model_path), "*Flat pattern", x, y, 0.0
            )
        except Exception:
            return None

    def _close_doc(self, doc):
        """Close a document by its CURRENT title and path (the title changes after SaveAs)."""
        if doc is None:
            return
        names = []
        for getter in ("GetTitle", "GetPathName"):
            value = self._get(doc, getter)
            if value:
                names.append(str(value))
        for name in names:
            try:
                self.sw_app.CloseDoc(name)
            except Exception:
                pass

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

    def _insert_bend_table(self, flat_view, x, y, steps):
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
            current = self._call(table, "Text", 0, column)
            if isinstance(current, str) and current.strip() == header:
                changed += 1
                continue
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
        steps.append(f"bend table headers: {changed}/4 ok (Tag / Direction / Angle / Inner Radius)")
        return changed == 4

    def _drop_blank_row(self, table, steps):
        """The default bend table has an empty row under the header; remove it."""
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

    def _place_bend_table(self, table, x_right, y_bottom, steps):
        """
        InsertBendTable already places the table with its bottom-right corner at
        (x_right, y_bottom) because it is called with anchor type BOTTOM_RIGHT,
        so no move is needed. This only reports the size.
        """
        width, height = self._table_size(table)
        steps.append(
            f"bend table: bottom-right corner at ({x_right * 1000:.0f}, {y_bottom * 1000:.0f}) mm, "
            f"size {width * 1000:.0f} x {height * 1000:.0f} mm"
        )
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
    # SELECTION HELPERS
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
        """Select an entity at sheet coordinates (metres). Only used as a fallback for bend lines."""
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
        """Select a COM entity (edge, bend line segment ...) inside a drawing view."""
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

    def _style_dimension(self, dim):
        """No parentheses and a text height that fits small views."""
        try:
            dim.ShowParenthesis = False
        except Exception:
            pass
        if self.dim_text_height:
            try:
                annotation = dim.GetAnnotation()
                fmt = annotation.GetTextFormat(0)
                fmt.CharHeight = float(self.dim_text_height)
                annotation.SetTextFormat(0, False, fmt)
            except Exception:
                pass

    def _add_checked_dimension(self, drawing, label, expected, tol=None):
        """AddDimension2 for the current selection. Kept only if its value matches `expected` (metres)."""
        try:
            dim = drawing.AddDimension2(float(label[0]), float(label[1]), 0.0)
        except Exception as exc:
            self._select_error = f"AddDimension2: {str(exc)[:60]}"
            self._clear(drawing)
            return False
        if dim is None:
            self._select_error = "AddDimension2 returned nothing"
            self._clear(drawing)
            return False
        value = self._dim_value(dim)
        self._clear(drawing)
        if value is not None and expected is not None:
            limit = tol if tol is not None else max(0.0005, 0.005 * expected)
            if abs(value - expected) > limit:
                self._select_error = (
                    f"wrong value {value * 1000:.2f} mm (expected {expected * 1000:.2f} mm) - removed"
                )
                self._delete_dimension(drawing, dim)
                return False
        self._style_dimension(dim)
        return True

    def _dim_between(self, drawing, view, first, second, label, expected, tol=None):
        self._clear(drawing)
        if not self._select_object(drawing, view, first, False):
            self._clear(drawing)
            return False
        if not self._select_object(drawing, view, second, True):
            self._clear(drawing)
            return False
        return self._add_checked_dimension(drawing, label, expected, tol)

    # =====================================================
    # VIEW GEOMETRY: the real edges of a view, in sheet coordinates
    # =====================================================

    def _visible_edges(self, view):
        """[{edge, points (model xyz of the end vertices), line (bool)}] of the edges SolidWorks shows in the view."""
        null = VARIANT(pythoncom.VT_DISPATCH, None)
        try:
            result = view.GetVisibleEntities2(null, self._edge_code())
        except Exception as exc:
            self._select_error = f"GetVisibleEntities2: {str(exc)[:60]}"
            return []

        edges = []
        for edge in list(result or []):
            curve = self._call(edge, "GetCurve")
            if curve is None:
                continue
            points = []
            for getter in ("GetStartVertex", "GetEndVertex"):
                vertex = self._call(edge, getter)
                point = self._xyz(self._call(vertex, "GetPoint")) if vertex is not None else None
                if point:
                    points.append(point)
            if not points:
                continue
            is_line = bool(self._call(curve, "IsLine", default=False)) and len(points) == 2
            edges.append({"edge": edge, "points": points, "line": is_line})
        return edges

    def _projector(self, view):
        """
        Returns (project, how, errors): project(model_xyz) -> (x, y) in view space.
        1) SolidWorks' own math objects (several ways to pass the point),
        2) the 4x4 transform data (ArrayData of ModelToViewTransform, or IView.GetXform) applied in Python.
        """
        errors = []
        transform = self._call(view, "ModelToViewTransform")
        math_util = self._call(self.sw_app, "GetMathUtility")

        if transform is not None and math_util is not None:
            makers = (
                ("variant array", lambda p: self._r8_array(*p)),
                ("tuple", lambda p: tuple(float(v) for v in p)),
                ("list", lambda p: [float(v) for v in p]),
            )
            for label, make in makers:
                def com(p, make=make):
                    math_point = math_util.CreatePoint(make(p))
                    result = math_point.MultiplyTransform(transform)
                    data = result.ArrayData
                    return float(data[0]), float(data[1])
                try:
                    com((0.0, 0.0, 0.0))
                    return com, f"SolidWorks math ({label})", errors
                except Exception as exc:
                    errors.append(f"{label}: {str(exc)[:70]}")
        else:
            errors.append(f"transform={'ok' if transform is not None else 'None'}, math utility={'ok' if math_util is not None else 'None'}")

        sources = (
            ("ModelToViewTransform.ArrayData", lambda: transform.ArrayData),
            ("IView.GetXform", lambda: view.GetXform()),
        )
        for label, getter in sources:
            try:
                raw = getter()
                data = [float(v) for v in raw]
                if len(data) < 13:
                    errors.append(f"{label}: only {len(data)} values")
                    continue
                r, t, k = data[:9], data[9:12], data[12] or 1.0

                def matrix(p, r=r, t=t, k=k):
                    return (
                        (r[0] * p[0] + r[1] * p[1] + r[2] * p[2]) * k + t[0],
                        (r[3] * p[0] + r[4] * p[1] + r[5] * p[2]) * k + t[1],
                    )
                return matrix, f"matrix from {label}", errors
            except Exception as exc:
                errors.append(f"{label}: {str(exc)[:70]}")
        return None, "", errors

    def _scale_of(self, view, fallback):
        try:
            value = float(self._get(view, "ScaleDecimal"))
            if value > 0:
                return value
        except Exception:
            pass
        return fallback

    def _view_geometry(self, view, scale_hint, steps, tag):
        """
        Reads the edges of a view and puts them on the sheet.
        The transform SolidWorks gives (model -> view) is checked against the view outline
        (scale included or not, origin at the view or at the sheet), so the result is
        self-calibrating. Returns None - and no dimension is made - when nothing fits.
        """
        edges = self._visible_edges(view)
        if not edges:
            steps.append(f"{tag}: no edges returned by SolidWorks ({self._select_error or 'empty list'})")
            return None

        project, how, errors = self._projector(view)
        outline = self._outline(view)
        if project is None or outline is None:
            steps.append(
                f"{tag}: no way to project the edges onto the sheet - dimensions skipped "
                f"({'; '.join(errors) or 'outline missing'})"
            )
            return None

        base = []
        for entry in edges:
            try:
                base.append([project(p) for p in entry["points"]])
            except Exception as exc:
                steps.append(f"{tag}: projecting an edge failed ({str(exc)[:60]}) - dimensions skipped")
                return None

        scale = self._scale_of(view, scale_hint)
        all_points = [p for pts in base for p in pts]
        ow, oh = outline[2] - outline[0], outline[3] - outline[1]
        ocx, ocy = (outline[0] + outline[2]) / 2, (outline[1] + outline[3]) / 2

        chosen = None
        for k, mode in ((1.0, "as is"), (scale, "scaled"), (1.0, "centred"), (scale, "scaled+centred")):
            xs = [p[0] * k for p in all_points]
            ys = [p[1] * k for p in all_points]
            gx0, gx1, gy0, gy1 = min(xs), max(xs), min(ys), max(ys)
            excess_x, excess_y = ow - (gx1 - gx0), oh - (gy1 - gy0)
            if not (-0.0015 <= excess_x <= 0.04 and -0.0015 <= excess_y <= 0.04):
                continue
            dx, dy = ocx - (gx0 + gx1) / 2, ocy - (gy0 + gy1) / 2
            if "centred" in mode:
                chosen = (k, dx, dy, mode)
                break
            if math.hypot(dx, dy) <= 0.006:
                chosen = (k, 0.0, 0.0, mode)
                break
        if chosen is None:
            steps.append(
                f"{tag}: the edge geometry does not fit the view outline ({ow * 1000:.0f} x {oh * 1000:.0f} mm) "
                f"- dimensions skipped instead of guessing"
            )
            return None

        k, dx, dy, mode = chosen

        def mapper(p3):
            q = project(p3)
            return (q[0] * k + dx, q[1] * k + dy)

        lines, xs, ys = [], [], []
        for entry, pts in zip(edges, base):
            sheet_pts = [(p[0] * k + dx, p[1] * k + dy) for p in pts]
            xs += [p[0] for p in sheet_pts]
            ys += [p[1] for p in sheet_pts]
            if entry["line"]:
                a, b = sheet_pts
                length = math.hypot(a[0] - b[0], a[1] - b[1])
                if length >= MIN_EDGE_LEN:
                    lines.append({"edge": entry["edge"], "a": a, "b": b, "len": length})

        bbox = (min(xs), min(ys), max(xs), max(ys))
        steps.append(
            f"{tag}: {len(edges)} edges ({len(lines)} straight), geometry "
            f"{(bbox[2] - bbox[0]) / scale * 1000:.1f} x {(bbox[3] - bbox[1]) / scale * 1000:.1f} mm "
            f"at 1:{1 / scale:g} (projection: {how}; {mode})"
        )
        return {"bbox": bbox, "lines": lines, "scale": scale, "mapper": mapper}

    @staticmethod
    def _vertical(lines):
        return [l for l in lines if abs(l["a"][0] - l["b"][0]) <= EDGE_TOL]

    @staticmethod
    def _horizontal(lines):
        return [l for l in lines if abs(l["a"][1] - l["b"][1]) <= EDGE_TOL]

    @staticmethod
    def _extreme(lines, axis, target):
        candidates = [l for l in lines if abs(l["a"][axis] - target) <= EDGE_TOL]
        return max(candidates, key=lambda l: l["len"]) if candidates else None

    # =====================================================
    # OVERALL DIMENSIONS
    # =====================================================

    def _dimension_overall(self, drawing, view, geo, steps, tag):
        """Overall length (above the view) and height (right of the view), between the outermost straight edges."""
        x0, y0, x1, y1 = geo["bbox"]
        scale = geo["scale"]
        vertical, horizontal = self._vertical(geo["lines"]), self._horizontal(geo["lines"])
        made = 0

        left, right = self._extreme(vertical, 0, x0), self._extreme(vertical, 0, x1)
        if left and right and left is not right:
            expected = (right["a"][0] - left["a"][0]) / scale
            self._select_error = ""
            if self._dim_between(drawing, view, left["edge"], right["edge"], ((x0 + x1) / 2, y1 + DIM_OFFSET), expected):
                made += 1
            else:
                steps.append(f"{tag}: overall length not created ({self._select_error or 'selection failed'})")
        else:
            steps.append(f"{tag}: overall length skipped (the left/right outline is not a straight vertical edge)")

        bottom, top = self._extreme(horizontal, 1, y0), self._extreme(horizontal, 1, y1)
        if bottom and top and bottom is not top:
            expected = (top["a"][1] - bottom["a"][1]) / scale
            self._select_error = ""
            if self._dim_between(drawing, view, bottom["edge"], top["edge"], (x1 + DIM_OFFSET, (y0 + y1) / 2), expected):
                made += 1
            else:
                steps.append(f"{tag}: overall height not created ({self._select_error or 'selection failed'})")
        else:
            steps.append(f"{tag}: overall height skipped (the top/bottom outline is not a straight horizontal edge)")
        return made

    # =====================================================
    # BEND LINES: positions from the geometry, checked with the A/B/C tag notes
    # =====================================================

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
                    if int(ann.GetType()) == SW_NOTE_TYPE:
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
        Fallback: sketch coordinates (real metres) -> sheet coordinates:  sheet = T + (1/den) * M * p.
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
            f"flat dimensions: bend line mapping by tags = {name}, residual {residual * 1000:.1f} mm, fill {fill:.2f}"
            + (" (single bend: orientation guessed from its length)" if n == 1 else "")
        )

        def mapper(p):
            q = apply(M, p)
            return (tx + s * q[0], ty + s * q[1])

        return mapper

    def _bend_mapper(self, pairs, tags, geo, steps):
        """
        Returns (mapper, exact). First try: the bend lines go through the same model->view transform as the
        edges (exact). It is accepted only if every bend line lies inside the geometry and, when tags exist,
        each tag is next to its bend line. Otherwise the tag based fit is used.
        """
        x0, y0, x1, y1 = geo["bbox"]
        pad = 0.001
        n = len(pairs)
        try:
            mapped = [(geo["mapper"](a), geo["mapper"](b)) for a, b in pairs]
            inside = all(
                x0 - pad <= p[0] <= x1 + pad and y0 - pad <= p[1] <= y1 + pad
                for a, b in mapped for p in (a, b)
            )
            good = inside
            residual = None
            if inside and len(tags) >= n:
                residual = sum(
                    math.hypot(tags[i][1] - (a[0] + b[0]) / 2, tags[i][2] - (a[1] + b[1]) / 2)
                    for i, (a, b) in enumerate(mapped)
                ) / n
                good = residual <= 0.010
            if good:
                steps.append(
                    "flat dimensions: bend lines located with the view transform"
                    + (f" (tag distance {residual * 1000:.1f} mm)" if residual is not None else " (no tags to compare)")
                )
                return geo["mapper"], True
        except Exception as exc:
            steps.append(f"flat dimensions: direct bend line mapping failed ({str(exc)[:50]})")

        mapper = self._fit_bend_mapping(pairs, tags, geo["bbox"], 1.0 / geo["scale"], steps)
        return mapper, False

    def _select_entity(self, drawing, view, ent, append):
        if ent["kind"] == "bend":
            if self._select_object(drawing, view, ent["seg"], append):
                return True
            for p in ent["along"]:
                if self._select_at(drawing, p[0], p[1], append, kinds=("SKETCHSEGMENT", "EDGE")):
                    return True
            return False
        return self._select_object(drawing, view, ent["edge"], append)

    def _dimension_chain(self, drawing, view, entities, geo, bottom_row, exact):
        """edge -> bend -> bend -> ... -> edge, one dimension between every neighbouring pair."""
        x0, y0, x1, y1 = geo["bbox"]
        scale = geo["scale"]
        made = 0
        entities = sorted(entities, key=lambda e: e["coord"])
        for index, (e1, e2) in enumerate(zip(entities, entities[1:])):
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
            tier = DIM_TIER_STEP if (gap < 0.012 and index % 2) else 0.0  # small neighbours: texts on two rows
            label = (mid, y0 - DIM_OFFSET - tier) if bottom_row else (x0 - DIM_OFFSET - tier, mid)
            both_bends = e1["kind"] == "bend" and e2["kind"] == "bend"
            tol = None if (exact or both_bends) else 0.003 / scale
            self._select_error = ""
            if self._add_checked_dimension(drawing, label, gap / scale, tol):
                made += 1
        self._clear(drawing)
        return made

    def _insert_flat_pattern_dimensions(self, drawing, view, den, steps):
        """
        Overall length (top) and height (right) + chain dimensions
        edge | bend | bend | edge along the bottom (vertical bend lines) and the left (horizontal bend lines).
        Needs the bend table to exist already (its A/B/C tags are used to check where the bend lines are).
        """
        tag = "flat dimensions"
        self._select_error = ""
        geo = self._view_geometry(view, 1.0 / den, steps, tag)
        if geo is None:
            return 0

        created = self._dimension_overall(drawing, view, geo, steps, tag)
        x0, y0, x1, y1 = geo["bbox"]
        vertical, horizontal = self._vertical(geo["lines"]), self._horizontal(geo["lines"])

        segments = []
        for seg in self._get_bend_lines(view):
            pts = self._line_segment_points(seg)
            if pts:
                segments.append((seg, pts))
        tags = self._view_tags(view)
        steps.append(
            f"{tag}: bend lines readable {len(segments)}, tags found {''.join(t[0] for t in tags) or 'none'}"
        )

        if segments:
            pairs = [p for _, p in segments]
            mapper, exact = self._bend_mapper(pairs, tags, geo, steps)
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
                    for line, coord in ((self._extreme(vertical, 0, x0), x0), (self._extreme(vertical, 0, x1), x1)):
                        if line:
                            chain.append({"kind": "edge", "coord": coord, "edge": line["edge"]})
                    n_v = self._dimension_chain(drawing, view, chain, geo, True, exact)
                if h_bends:
                    chain = list(h_bends)
                    for line, coord in ((self._extreme(horizontal, 1, y0), y0), (self._extreme(horizontal, 1, y1), y1)):
                        if line:
                            chain.append({"kind": "edge", "coord": coord, "edge": line["edge"]})
                    n_h = self._dimension_chain(drawing, view, chain, geo, False, exact)
                steps.append(
                    f"{tag}: chain along the bottom {n_v}/{len(v_bends) + 1 if v_bends else 0}, "
                    f"along the left {n_h}/{len(h_bends) + 1 if h_bends else 0}"
                )
                created += n_v + n_h
        if self._select_error:
            steps.append(f"{tag}: last problem: {self._select_error}")
        steps.append(f"{tag}: {created} native dimensions in total")
        return created

    # =====================================================
    # VIEW SCALE
    # =====================================================

    @staticmethod
    def _fit_denominator(raw_w, raw_h, avail_w, avail_h):
        aw, ah = max(avail_w - MARGIN_GUESS, 0.005), max(avail_h - MARGIN_GUESS, 0.005)
        den = max((raw_w - MARGIN_GUESS) / aw, (raw_h - MARGIN_GUESS) / ah, 1.0)
        return math.ceil(den * 10) / 10

    @staticmethod
    def _scale_text(den):
        return "1:" + f"{den:.2f}".rstrip("0").rstrip(".")

    # =====================================================
    # SAVE / PDF
    # =====================================================

    def _export_pdf(self, doc, pdf_path, steps):
        attempts = (
            lambda: doc.SaveAs3(str(pdf_path), 0, 1 | 2),
            lambda: doc.Extension.SaveAs3(
                str(pdf_path), 0, 1 | 2,
                VARIANT(pythoncom.VT_DISPATCH, None), VARIANT(pythoncom.VT_DISPATCH, None),
                self._byref_i4(), self._byref_i4(),
            ),
        )
        for attempt in attempts:
            try:
                attempt()
            except Exception as exc:
                steps.append(f"pdf save attempt: {str(exc)[:60]}")
            if self._written(pdf_path):
                return True
        return False

    
    def _open_drawing(self, path, steps):
        try:
            return self.sw_app.OpenDoc6(
                str(path),
                SW_DOC_DRAWING,
                SW_OPEN_SILENT,
                "",
                self._byref_i4(),
                self._byref_i4(),
            )
        except Exception as exc:
            steps.append(
                f"reopening the .slddrw failed: {str(exc)[:100]}"
            )
            return None

    
    def _apply_sheet_format(self, drawing, steps):
        if not self.sheet_format:
            steps.append("sheet format: none configured, using the template sheet")
            return True

        format_path = Path(self.sheet_format).resolve()
        if not format_path.is_file():
            steps.append(f"sheet format: file not found: {format_path}")
            return False

        constants = self._constants()
        custom = getattr(constants, "swDwgTemplateCustom", 12) if constants else 12
        user_defined = getattr(constants, "swDwgPapersUserDefined", 12) if constants else 12

        try:
            sheet = drawing.GetCurrentSheet()
            name = sheet.GetName()
            _, w, h = self._sheet_size(drawing)
            # size of YOUR format (A3 = 0.420 x 0.297)
            w = self.sheet_width or w
            h = self.sheet_height or h

            ok = drawing.SetupSheet5(
                name, user_defined, custom, 1.0, 1.0, True,
                str(format_path), float(w), float(h), "", True,
            )
            if ok:
                drawing.ForceRebuild3(False)
                steps.append(f"sheet format applied (SetupSheet5): {format_path.name}")
                return True
            steps.append("SetupSheet5 returned False")
        except Exception as exc:
            steps.append(f"SetupSheet5 failed: {type(exc).__name__}: {str(exc)[:100]}")

        # second attempt
        try:
            sheet = drawing.GetCurrentSheet()
            sheet.SetTemplateName(str(format_path))
            code = sheet.ReloadTemplate(True)      # 0 = success
            if code in (0, None, True):
                drawing.ForceRebuild3(False)
                steps.append(f"sheet format applied (ReloadTemplate): {format_path.name}")
                return True
            steps.append(f"ReloadTemplate error code: {code}")
        except Exception as exc:
            steps.append(f"ReloadTemplate failed: {type(exc).__name__}: {str(exc)[:100]}")
        return False

    # =====================================================
    # ONE PART
    # =====================================================

    def _build_drawing(self, item, slddrw_path, pdf_path):
        steps = []
        path = Path(item["path"]).resolve()
        configuration = item.get("configuration")
        model = None
        opened_by_us = False
        previous_config = None
        drawing = None
        reopened = None

        try:
            template = self.find_template()
            if not template:
                return False, "no drawing template found (set DRAWING_TEMPLATE_PATH in main.py)", steps
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
            self._current_drawing = drawing
            if drawing is None:
                return False, "NewDocument returned nothing", steps

            self._ensure_millimeters(drawing, steps)

            if not self._apply_sheet_format(drawing, steps):
                steps.append("WARNING: custom sheet format NOT applied, continuing with the default template")

            sheet, sheet_w, sheet_h = self._sheet_size(drawing)
            steps.append(f"sheet: {sheet_w * 1000:.0f} x {sheet_h * 1000:.0f} mm (from the default template)")
            self._call(sheet, "SetScale", 1.0, 1.0, True, True)
            self._select_error = ""

            # ---- 1) flat pattern at 1:1 to learn its size ----
            flat = self._create_flat_view(drawing, path, configuration, sheet_w / 2, sheet_h / 2)
            outline = self._outline(flat) if flat is not None else None
            if outline is None:
                return False, "flat pattern view could not be created", steps
            raw_w, raw_h = outline[2] - outline[0], outline[3] - outline[1]
            steps.append(f"flat pattern view: {raw_w * 1000:.0f} x {raw_h * 1000:.0f} mm at 1:1")

            # ---- 2) layout and scale ----
            bend_count = len(self._get_bend_lines(flat))
            table_estimate = (bend_count + 3) * ROW_HEIGHT_GUESS if self.bend_table else 0.0

            # bottom-right corner of the bend table, against the template frame
            table_right = sheet_w - self.frame_right - 0.001

            area_x0 = SIDE_MARGIN + FLAT_DIM_ROOM
            area_x1 = sheet_w - SIDE_MARGIN - FLAT_DIM_ROOM

            if self.table_corner == "top_right":
                table_bottom = sheet_h - TOP_MARGIN - table_estimate
                area_top = table_bottom - FLAT_BOTTOM_ROOM
                area_bottom = self.bottom_margin + FLAT_DIM_ROOM
            else:  # "bottom_right"
                table_bottom = self.bottom_margin + 0.001
                area_top = sheet_h - TOP_MARGIN - FLAT_DIM_ROOM
                area_bottom = table_bottom + table_estimate + FLAT_BOTTOM_ROOM

            flat_den = self._fit_denominator(
                raw_w, raw_h, area_x1 - area_x0, area_top - area_bottom
            )
            steps.append(f"scale: flat pattern {self._scale_text(flat_den)}")
            self._call(sheet, "SetScale", 1.0, float(flat_den), True, True)

            # ---- 3) position ----
            fx = (area_x0 + area_x1) / 2
            fy = (area_top + area_bottom) / 2
            if not self._move_view_center(flat, fx, fy):
                return False, "could not move the flat view", steps
            self._call(drawing, "ForceRebuild3", False)

            # ---- 4) bend table first: it creates the A/B/C tags used to check the bend line positions ----
            if self.bend_table:
                table = self._insert_bend_table(flat, table_right, table_bottom, steps)
                if table is not None:
                    if self.edit_bend_table_text:
                        self._set_bend_table_headers(table, steps)
                        self._drop_blank_row(table, steps)
                    table_height = self._place_bend_table(table, table_right, table_bottom, steps)
                    if table_height > table_estimate + 0.003:
                        steps.append(
                            f"WARNING: table is {table_height * 1000:.0f} mm tall, only "
                            f"{table_estimate * 1000:.0f} mm were reserved - it may touch the flat pattern"
                        )
                    self._describe_table(table, steps)
                else:
                    steps.append("BEND TABLE MISSING in this drawing")

            # ---- 5) native dimensions of the flat pattern ----
            self._call(drawing, "ForceRebuild3", False)
            self._call(drawing, "ViewZoomtofit2")
            if self.dimensions:
                self._insert_flat_pattern_dimensions(drawing, flat, flat_den, steps)
                self._clear(drawing)
            self._call(drawing, "ForceRebuild3", False)

            # ---- 6) save the drawing, then the PDF ----
            self._call(drawing, "ViewZoomtofit2")
            slddrw_path.parent.mkdir(parents=True, exist_ok=True)
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            for old in (slddrw_path, pdf_path):
                if old.exists():
                    try:
                        old.unlink()
                    except Exception:
                        pass

            saved_drawing = False
            try:
                drawing.SaveAs3(str(slddrw_path), 0, 1)
            except Exception as exc:
                steps.append(f"save .slddrw: FAILED ({str(exc)[:60]})")
            else:
                saved_drawing = self._written(slddrw_path)
                steps.append("saved .slddrw" if saved_drawing else "save .slddrw: no file")

            # first PDF straight from the open drawing (safety net) ...
            saved_pdf = self._export_pdf(drawing, pdf_path, steps)

            # ... then the final PDF from the SAVED .slddrw (close it, open the file again)
            if self.pdf_from_saved_file and saved_drawing:
                self._close_doc(drawing)
                drawing = None
                reopened = self._open_drawing(slddrw_path, steps)
                if reopened is not None:
                    temp_pdf = pdf_path.with_name(pdf_path.stem + "__from_slddrw.pdf")
                    if temp_pdf.exists():
                        try:
                            temp_pdf.unlink()
                        except Exception:
                            pass
                    if self._export_pdf(reopened, temp_pdf, steps):
                        try:
                            os.replace(str(temp_pdf), str(pdf_path))
                            saved_pdf = True
                            steps.append("PDF exported from the saved .slddrw")
                        except Exception as exc:
                            steps.append(f"could not replace the PDF ({str(exc)[:50]}), kept the first one")
                    else:
                        steps.append("PDF from the saved .slddrw failed, kept the PDF made from the open drawing")

            if not saved_pdf:
                return False, "PDF was not written", steps
            return True, "SolidWorks drawing", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
            self._close_doc(reopened)
            self._close_doc(drawing)
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
        print("PDF DRAWINGS (flat pattern + bend table, one folder per thickness)")
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