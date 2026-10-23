import math
import os
import re
import shutil
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
SW_NOTE_TYPE = 6  # swNote
NO_OPTION_SPECIFIED = 0

# ---------------------------------------------------------
# Layout / dimension constants (sheet metres)
# ---------------------------------------------------------
MARGIN_GUESS = 0.0084  # a view outline is wider than its geometry by about this much
PICK_FRACTIONS = (0.5, 0.35, 0.65, 0.2, 0.8, 0.1, 0.9)
DIM_OFFSET = 0.008
DIM_TIER_STEP = 0.005
EDGE_TOL = 1.5e-4
MIN_EDGE_LEN = 5e-4

STANDARD_DENOMINATORS = (1, 1.5, 2, 2.5, 3, 4, 5, 7.5, 10, 15, 20, 25, 30, 40, 50, 60, 75,
                         100, 150, 200, 250, 300, 500)


def _standard_denominator(den):
    """Smallest standard scale denominator that is not smaller than `den`."""
    for value in STANDARD_DENOMINATORS:
        if value >= den - 1e-9:
            return float(value)
    return float(math.ceil(den))


class SwDrawingPdfExporter(DwgExporter):
    """
    For every sheet metal part:
      1) copies master.slddrw (a complete sample drawing: flat pattern, bend table, bent views, title block),
      2) replaces the model referenced by the copy with the part,
      3) fits every view into the space it had in the master (same centre),
      4) adds native dimensions on the flat pattern,
      5) saves the .slddrw and exports the PDF from the SAVED .slddrw.
    Every step is logged in `steps`.
    """

    def __init__(
        self,
        sw_app,
        assembly_path=None,
        reconnect=None,
        master_drawing=None,
        dimensions=True,
        dim_text_height=0.0025,
        pdf_from_saved_file=True,
        remove_master_dimensions=True,
        **legacy_options,
    ):
        super().__init__(
            sw_app,
            assembly_path=assembly_path,
            reconnect=reconnect,
        )
        self.master_drawing = Path(master_drawing).resolve() if master_drawing else None
        self.dimensions = dimensions
        self.dim_text_height = dim_text_height
        self.pdf_from_saved_file = pdf_from_saved_file
        self.remove_master_dimensions = remove_master_dimensions
        self._constants_cache = None
        self._typelib_cache = None
        self._select_error = ""

    # =====================================================
    # SOLIDWORKS CONSTANTS / UNITS
    # =====================================================

    def _constants(self):
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

    def _typelib_module(self):
        """sldworks.tlb of YOUR SolidWorks as a generated (early-bound) module."""
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

    def _cast_strict(self, obj, interface):
        """Typed wrapper of a COM object for one interface, or None when that is not possible."""
        if obj is None:
            return None
        module = self._typelib_module()
        if module is None:
            return None
        try:
            iid = module.NamesToIIDMap[interface]
            cls = getattr(module, interface)
            return cls(obj._oleobj_.QueryInterface(iid, pythoncom.IID_IDispatch))
        except Exception:
            return None

    def _cast(self, obj, interface):
        typed = self._cast_strict(obj, interface)
        return typed if typed is not None else obj

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

    def _ensure_millimeters(self, drawing, steps):
        constants = self._constants()
        if constants is None:
            steps.append("units: swconst.tlb could not be loaded - units NOT changed")
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

    # =====================================================
    # TABLE REPORT
    # =====================================================

    def _describe_table(self, table, steps):
        try:
            rows = int(self._get(table, "RowCount"))
            columns = int(self._get(table, "ColumnCount"))
        except Exception:
            steps.append("table content: could not be read")
            return
        steps.append(f"table content ({max(rows - 1, 0)} rows after the header):")
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
        for target in (self._cast(view, "IView"), view):
            try:
                return list(target.GetBendLines() or [])
            except Exception:
                continue
        return []

    def _select_at(self, drawing, x, y, append, kinds=("EDGE", "SKETCHSEGMENT")):
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
        stage = "SelectionManager"
        try:
            sel_mgr = drawing.SelectionManager
            stage = "CreateSelectData"
            try:
                data = sel_mgr.CreateSelectData()
            except Exception:
                data = self._cast(sel_mgr, "ISelectionMgr").CreateSelectData()
            stage = "SelectData.View"
            data.View = view
            stage = "Select4"
            candidates = [obj]
            for name in ("IEdge", "IEntity", "ISketchSegment"):
                typed = self._cast_strict(obj, name)
                if typed is not None:
                    candidates.append(typed)
            last = None
            for candidate in candidates:
                try:
                    return bool(candidate.Select4(bool(append), data))
                except Exception as exc:
                    last = exc
            raise last
        except Exception as exc:
            self._select_error = f"{stage}: {str(exc)[:70]}"
            return False

    def _select_edge(self, drawing, view, edge, append, line=None):
        """Select an edge by its COM object; if that fails, by its middle point on the sheet."""
        if self._select_object(drawing, view, edge, append):
            return True
        if line is not None:
            mx = (line["a"][0] + line["b"][0]) / 2
            my = (line["a"][1] + line["b"][1]) / 2
            if self._select_at(drawing, mx, my, append, kinds=("EDGE",)):
                return True
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
        if not self._select_edge(drawing, view, first["edge"], False, first):
            self._clear(drawing)
            return False
        if not self._select_edge(drawing, view, second["edge"], True, second):
            self._clear(drawing)
            return False
        return self._add_checked_dimension(drawing, label, expected, tol)

    # =====================================================
    # VIEW GEOMETRY
    # =====================================================

    def _visible_edges(self, view):
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
            errors.append(
                f"transform={'ok' if transform is not None else 'None'}, "
                f"math utility={'ok' if math_util is not None else 'None'}"
            )

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
        x0, y0, x1, y1 = geo["bbox"]
        scale = geo["scale"]
        vertical, horizontal = self._vertical(geo["lines"]), self._horizontal(geo["lines"])
        made = 0

        left, right = self._extreme(vertical, 0, x0), self._extreme(vertical, 0, x1)
        if left and right and left is not right:
            expected = (right["a"][0] - left["a"][0]) / scale
            self._select_error = ""
            if self._dim_between(drawing, view, left, right, ((x0 + x1) / 2, y1 + DIM_OFFSET), expected):
                made += 1
            else:
                steps.append(f"{tag}: overall length not created ({self._select_error or 'selection failed'})")
        else:
            steps.append(f"{tag}: overall length skipped (the left/right outline is not a straight vertical edge)")

        bottom, top = self._extreme(horizontal, 1, y0), self._extreme(horizontal, 1, y1)
        if bottom and top and bottom is not top:
            expected = (top["a"][1] - bottom["a"][1]) / scale
            self._select_error = ""
            if self._dim_between(drawing, view, bottom, top, (x1 + DIM_OFFSET, (y0 + y1) / 2), expected):
                made += 1
            else:
                steps.append(f"{tag}: overall height not created ({self._select_error or 'selection failed'})")
        else:
            steps.append(f"{tag}: overall height skipped (the top/bottom outline is not a straight horizontal edge)")
        return made

    # =====================================================
    # BEND LINES
    # =====================================================

    def _view_tags(self, view):
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
        return self._select_edge(drawing, view, ent["edge"], append, ent.get("line"))

    def _dimension_chain(self, drawing, view, entities, geo, bottom_row, exact):
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
            tier = DIM_TIER_STEP if (gap < 0.012 and index % 2) else 0.0
            label = (mid, y0 - DIM_OFFSET - tier) if bottom_row else (x0 - DIM_OFFSET - tier, mid)
            both_bends = e1["kind"] == "bend" and e2["kind"] == "bend"
            tol = None if (exact or both_bends) else 0.003 / scale
            self._select_error = ""
            if self._add_checked_dimension(drawing, label, gap / scale, tol):
                made += 1
        self._clear(drawing)
        return made

    def _insert_flat_pattern_dimensions(self, drawing, view, den, steps):
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
                            chain.append({"kind": "edge", "coord": coord, "edge": line["edge"], "line": line})
                    n_v = self._dimension_chain(drawing, view, chain, geo, True, exact)
                if h_bends:
                    chain = list(h_bends)
                    for line, coord in ((self._extreme(horizontal, 1, y0), y0), (self._extreme(horizontal, 1, y1), y1)):
                        if line:
                            chain.append({"kind": "edge", "coord": coord, "edge": line["edge"], "line": line})
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
        return _standard_denominator(den)

    @staticmethod
    def _scale_text(den):
        return "1:" + f"{den:.2f}".rstrip("0").rstrip(".")

    # =====================================================
    # OPEN / SAVE / PDF
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

    # =====================================================
    # MASTER DRAWING
    # =====================================================

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

    def _set_view_scale(self, view, scale):
        try:
            view.UseSheetScale = False
        except Exception:
            pass
        try:
            view.ScaleDecimal = float(scale)
            return True
        except Exception:
            pass
        try:
            view.ScaleRatio = self._r8_array(1.0, 1.0 / float(scale))
            return True
        except Exception:
            return False

    def _fit_view_to_slot(self, view, slot, steps):
        """Scale a view so that it fits the space it had in the master and keep the same centre."""
        name = slot["name"]
        cx, cy, slot_w, slot_h = slot["box"]
        if not self._set_view_scale(view, 1.0):
            steps.append(f"view {name}: scale could not be set (child view follows its parent?)")
            return 1.0 / self._scale_of(view, 1.0)
        outline = self._outline(view)
        if outline is None:
            steps.append(f"view {name}: outline not readable")
            return 1.0
        raw_w, raw_h = outline[2] - outline[0], outline[3] - outline[1]
        den = self._fit_denominator(raw_w, raw_h, slot_w, slot_h)
        self._set_view_scale(view, 1.0 / den)
        moved = self._move_view_center(view, cx, cy)
        steps.append(
            f"view {name}: {raw_w * 1000:.0f} x {raw_h * 1000:.0f} mm at 1:1 -> scale {self._scale_text(den)}"
            + ("" if moved else " (centre NOT confirmed)")
        )
        return den

    def _replace_model(self, drawing_path, old_ref, new_ref, steps):
        """The drawing must be CLOSED. Replaces the model referenced by the drawing file."""
        try:
            result = self.sw_app.ReplaceReferencedDocument(str(drawing_path), str(old_ref), str(new_ref))
            steps.append(f"model replaced (ReplaceReferencedDocument): result={result}")
            if result in (None, True, 1):
                return True
        except Exception as exc:
            steps.append(f"ReplaceReferencedDocument failed ({type(exc).__name__}: {str(exc)[:80]})")
        return False

    def _replace_model_in_open_drawing(self, drawing, views, new_ref, steps):
        """Fallback: replace the model of the views while the drawing is open (IDrawingDoc.ReplaceViewModel)."""
        names = [str(self._call(v, "GetName2") or "") for v in views]
        attempts = (
            ("names", lambda: drawing.ReplaceViewModel(str(new_ref), tuple(names), tuple([""] * len(names)))),
            ("variant", lambda: drawing.ReplaceViewModel(
                str(new_ref),
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, tuple(names)),
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, tuple([""] * len(names))))),
        )
        for label, attempt in attempts:
            try:
                result = attempt()
                steps.append(f"model replaced (ReplaceViewModel {label}): result={result}")
                if result in (None, True, 1):
                    return True
            except Exception as exc:
                steps.append(f"ReplaceViewModel {label} failed ({type(exc).__name__}: {str(exc)[:70]})")
        return False

    @staticmethod
    def _val(obj, name):
        attr = getattr(obj, name)
        return attr() if callable(attr) else attr

    def _walk_features(self, feat, depth=0, guard=None):
        guard = guard if guard is not None else [0]
        while feat is not None and guard[0] < 3000:
            guard[0] += 1
            feat = self._cast(feat, "IFeature")
            yield feat
            try:
                sub = feat.GetFirstSubFeature()
            except Exception:
                sub = None
            if sub is not None and depth < 6:
                yield from self._walk_features(sub, depth + 1, guard)
            try:
                feat = feat.GetNextFeature() if depth == 0 else feat.GetNextSubFeature()
            except Exception:
                feat = None

    def _find_bend_tables(self, drawing, flat_view, steps, tag):
        tables = []
        # route 1: the feature tree (Drawing View > Bend Table1)
        try:
            first = self._val(self._cast(drawing, "IModelDoc2"), "FirstFeature")
            seen = []
            for feat in self._walk_features(first):
                try:
                    type_name = str(feat.GetTypeName2())
                    name = str(self._val(feat, "Name"))
                except Exception:
                    continue
                low = (type_name + " " + name).lower()
                if "table" in low:
                    seen.append(f"{name} [{type_name}]")
                if "bend" in low and "table" in low and "anchor" not in low:
                    try:
                        spec = feat.GetSpecificFeature2()
                    except Exception:
                        spec = None
                    if spec is not None:
                        typed = (self._cast_strict(spec, "IBendTableAnnotation")
                                 or self._cast_strict(spec, "ITableAnnotation"))
                        tables.append(typed if typed is not None else spec)
            steps.append(f"{tag}: table features seen: {', '.join(seen) or 'none'}")
        except Exception as exc:
            steps.append(f"{tag}: feature walk failed ({type(exc).__name__}: {str(exc)[:60]})")
        # route 2: methods of the view
        if not tables:
            for method in ("GetBendTableAnnotations", "GetTableAnnotations"):
                try:
                    found = list(self._call(self._cast(flat_view, "IView"), method) or [])
                except Exception:
                    found = []
                steps.append(f"{tag}: view.{method}() -> {len(found)}")
                if found:
                    tables = found
                    break
        return tables

    def _refresh_bend_tables(self, drawing, flat_view, steps):
        tag = "bend table"
        steps.append(f"{tag}: bend lines readable by the API: {len(self._get_bend_lines(flat_view))}")
        tables = self._find_bend_tables(drawing, flat_view, steps, tag)
        steps.append(f"{tag}: tables found: {len(tables)}")
        for index, table in enumerate(tables, 1):
            rows_before = self._get(table, "RowCount")
            try:
                table.Update()
            except Exception as exc:
                steps.append(f"{tag} {index}: Update failed ({type(exc).__name__}: {str(exc)[:60]})")
            self._call(drawing, "ForceRebuild3", False)
            rows_after = self._get(table, "RowCount")
            steps.append(f"{tag} {index}: rows {rows_before} -> {rows_after}")
            try:
                ann = self._val(table, "GetAnnotation")
                pos = self._val(ann, "GetPosition")
                steps.append(
                    f"{tag} {index}: position ({float(pos[0]) * 1000:.0f}, {float(pos[1]) * 1000:.0f}) mm, "
                    f"visible={self._val(ann, 'Visible')}"
                )
            except Exception as exc:
                steps.append(f"{tag} {index}: position not readable ({type(exc).__name__}: {str(exc)[:50]})")
            self._describe_table(table, steps)

    def _remove_view_dimensions(self, drawing, views, steps):
        """Deletes the dimensions that were drawn on the master (they do not belong to the new part)."""
        dims = []
        for view in views:
            typed_view = self._cast(view, "IView")
            try:
                dim = typed_view.GetFirstDisplayDimension5()
            except Exception as exc:
                steps.append(f"master dimensions: cannot be read ({type(exc).__name__}: {str(exc)[:60]})")
                return 0
            guard = 0
            while dim is not None and guard < 3000:
                guard += 1
                dim = self._cast(dim, "IDisplayDimension")
                dims.append(dim)
                try:
                    dim = dim.GetNext5()
                except Exception:
                    break
        removed = 0
        for dim in dims:
            self._clear(drawing)
            try:
                ann = self._cast(dim.GetAnnotation(), "IAnnotation")
                try:
                    ann.Select3(False, None)
                except Exception:
                    ann.Select3(False, VARIANT(pythoncom.VT_DISPATCH, None))
                if drawing.Extension.DeleteSelection2(0):
                    removed += 1
            except Exception:
                pass
        self._clear(drawing)
        steps.append(f"master dimensions: {removed}/{len(dims)} removed")
        return removed

    def _save_and_pdf(self, drawing, slddrw_path, pdf_path, steps):
        """Saves the drawing, closes it, reopens the SAVED file and exports the PDF from it."""
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

        if pdf_path.exists():
            try:
                pdf_path.unlink()
            except Exception:
                pass

        if self.pdf_from_saved_file and saved:
            self._close_doc(drawing)
            reopened = self._open_drawing(slddrw_path, steps)
            if reopened is not None:
                ok = self._export_pdf(reopened, pdf_path, steps)
                if ok:
                    steps.append("PDF exported from the saved .slddrw")
                self._close_doc(reopened)
                return ok, True
            steps.append("could not reopen the saved .slddrw")
            return False, True
        return self._export_pdf(drawing, pdf_path, steps), False

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

        try:
            master = self.master_drawing
            if master is None or not master.is_file():
                return False, f"master drawing not found: {master}", steps
            steps.append(f"master: {master.name}")
            steps.append(f"part file: {path} (exists={path.exists()})")

            model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                time.sleep(2)
                model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                return False, f"cannot open part: {open_status}", steps
            _, previous_config = self._activate_configuration(model, configuration)

            # ---- 1) copy the master ----
            slddrw_path.parent.mkdir(parents=True, exist_ok=True)
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            if slddrw_path.exists():
                slddrw_path.unlink()
            shutil.copy2(str(master), str(slddrw_path))
            try:
                os.chmod(str(slddrw_path), 0o666)
            except Exception:
                pass

            # ---- 2) read the model reference and the space of every view ----
            doc = self._open_drawing(slddrw_path, steps)
            if doc is None:
                return False, "cannot open the copy of the master", steps
            old_ref = None
            slots = {}
            for view in self._model_views(doc):
                name = str(self._call(view, "GetName2") or "")
                ref = self._call(view, "GetReferencedModelName")
                if ref and not old_ref:
                    old_ref = str(ref)
                outline = self._outline(view)
                if outline is None or not name:
                    continue
                slots[name] = {
                    "name": name,
                    "box": ((outline[0] + outline[2]) / 2, (outline[1] + outline[3]) / 2,
                            outline[2] - outline[0], outline[3] - outline[1]),
                }
            self._close_doc(doc)
            doc = None
            steps.append(f"master views: {', '.join(slots) or 'NONE'}; model: {Path(old_ref).name if old_ref else '?'}")
            if not old_ref or not slots:
                return False, "no model views found in the master drawing", steps

            # ---- 3) replace the model (drawing closed) ----
            replaced_in_closed = True
            same_part = False
            try:
                same_part = Path(old_ref).resolve() == path
            except Exception:
                pass
            if not same_part:
                try:
                    self.sw_app.CloseDoc(old_ref)
                except Exception:
                    pass
                replaced_in_closed = self._replace_model(slddrw_path, old_ref, path, steps)
            else:
                steps.append("the part is the master's own part - no replacement needed")

            # ---- 4) reopen, replace in the open drawing if needed, rebuild, fit views ----
            drawing = self._open_drawing(slddrw_path, steps)
            if drawing is None:
                return False, "cannot reopen the drawing", steps
            if not replaced_in_closed:
                if not self._replace_model_in_open_drawing(drawing, self._model_views(drawing), path, steps):
                    return False, "could not replace the model in the drawing", steps
            self._ensure_millimeters(drawing, steps)
            steps.append("typed SolidWorks API (sldworks.tlb): " + ("loaded" if self._typelib_module() else "NOT available"))
            self._call(drawing, "ForceRebuild3", False)

            if self.remove_master_dimensions:
                self._remove_view_dimensions(drawing, self._model_views(drawing), steps)

            flat_view, flat_den = None, 1.0
            for view in self._model_views(drawing):
                name = str(self._call(view, "GetName2") or "")
                if configuration:
                    try:
                        view.ReferencedConfiguration = str(configuration)
                    except Exception:
                        pass
                slot = slots.get(name)
                if slot is None:
                    continue
                den = self._fit_view_to_slot(view, slot, steps)
                is_flat = bool(self._call(view, "IsFlatPatternView", default=False))
                if is_flat:
                    flat_view, flat_den = view, den
            self._call(drawing, "ForceRebuild3", False)

            # ---- 5) bend table update + dimensions ----
            if flat_view is None:
                steps.append("WARNING: no flat pattern view found in the drawing")
            else:
                self._refresh_bend_tables(drawing, flat_view, steps)
                if self.dimensions:
                    flat_den = 1.0 / self._scale_of(flat_view, 1.0 / flat_den)
                    self._insert_flat_pattern_dimensions(drawing, flat_view, flat_den, steps)
                    self._clear(drawing)
            self._call(drawing, "ForceRebuild3", False)
            self._call(drawing, "ViewZoomtofit2")

            # ---- 6) save + PDF from the saved file ----
            ok, closed = self._save_and_pdf(drawing, slddrw_path, pdf_path, steps)
            if closed:
                drawing = None
            if not ok:
                return False, "PDF was not written", steps
            return True, "SolidWorks drawing (from master)", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
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

    @staticmethod
    def _safe_name(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"