import json
import math
from pathlib import Path

PT_PER_MM = 72.0 / 25.4          # 1:1 on paper
BEND_LAYER_WORDS = ("BEND",)
DASHED_WORDS = ("DASH", "CENTER", "PHANTOM", "HIDDEN", "DOT")
AXIS_TOLERANCE = 0.01            # a bend line is "vertical/horizontal" if it leans less than 1 %
POSITION_TOLERANCE = 0.05        # mm: bend lines closer than this share one dimension
MATCH_TOLERANCE = 0.6            # mm: a SolidWorks bend line and a DXF line are "the same" within this
TAG_MAX_DISTANCE = 15.0          # sheet mm: a tag belongs to a bend line only if it sits this close to it
GEOMETRY_TYPES = ("LINE", "ARC", "CIRCLE", "LWPOLYLINE", "POLYLINE", "SPLINE", "ELLIPSE")
# the 8 ways a view can sit relative to its sketch: 4 rotations, each with or without a mirror
ORIENTATIONS = [
    (1, 0, 0, 1), (0, -1, 1, 0), (-1, 0, 0, -1), (0, 1, -1, 0),
    (-1, 0, 0, 1), (1, 0, 0, -1), (0, 1, 1, 0), (0, -1, -1, 0),
]


def _dist(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def _point_segment_distance(p, a, b):
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    length_sq = dx * dx + dy * dy
    if length_sq < 1e-12:
        return _dist(p, a)
    t = max(0.0, min(1.0, ((p[0] - ax) * dx + (p[1] - ay) * dy) / length_sq))
    return _dist(p, (ax + t * dx, ay + t * dy))


def match_bend_lines(probe_lines, dxf_lines, tolerance=MATCH_TOLERANCE):
    """
    SolidWorks tells us the bend lines in the flat pattern's own sketch coordinates; the drawing DXF has the
    same lines in sheet coordinates, but turned and moved. Find the rotation / mirror / shift that puts ALL
    probe lines exactly on DXF lines, and return the indexes of those DXF lines.

    probe_lines: [((x1, y1), (x2, y2)), ...] in mm.   dxf_lines: [(key, (x1, y1), (x2, y2)), ...] in mm.
    Returns (set_of_keys or None, message).
    """
    if not probe_lines or not dxf_lines:
        return None, "no bend lines to match"

    def close(a, b):
        return _dist(a, b) <= tolerance

    first_a, first_b = probe_lines[0]
    first_len = _dist(first_a, first_b)
    solutions = set()

    for (m00, m01, m10, m11) in ORIENTATIONS:
        def move(p, m=(m00, m01, m10, m11)):
            return (m[0] * p[0] + m[1] * p[1], m[2] * p[0] + m[3] * p[1])

        for key, s, e in dxf_lines:
            if abs(_dist(s, e) - first_len) > tolerance:
                continue
            for u, v in ((s, e), (e, s)):
                a0, b0 = move(first_a), move(first_b)
                shift = (u[0] - a0[0], u[1] - a0[1])
                if not close((b0[0] + shift[0], b0[1] + shift[1]), v):
                    continue

                matched = set()
                for pa, pb in probe_lines:
                    A = move(pa)
                    B = move(pb)
                    A = (A[0] + shift[0], A[1] + shift[1])
                    B = (B[0] + shift[0], B[1] + shift[1])
                    hit = next(
                        (k for k, s2, e2 in dxf_lines if (close(A, s2) and close(B, e2)) or (close(A, e2) and close(B, s2))),
                        None,
                    )
                    if hit is None:
                        matched = None
                        break
                    matched.add(hit)
                if matched:
                    solutions.add(frozenset(matched))

    if not solutions:
        return None, "the SolidWorks bend lines were not found in the DXF"
    if len(solutions) > 1:
        return None, f"{len(solutions)} different placements fit - ambiguous"
    return set(next(iter(solutions))), "OK"


def assign_tags(bend_lines, tags, max_distance):
    """
    Each tag sits right next to its bend line. Pair them by distance, nearest pairs first, one line per tag.
    bend_lines: [((x1, y1), (x2, y2)), ...]   tags: [(text, (x, y)), ...]
    Returns [{"tag": text, "pos": (x, y), "line": index}, ...] (only pairs within max_distance).
    """
    pairs = []
    for ti, (_text, pos) in enumerate(tags):
        for li, (a, b) in enumerate(bend_lines):
            pairs.append((_point_segment_distance(pos, a, b), ti, li))
    pairs.sort()

    used_tags, used_lines, result = set(), set(), []
    for distance, ti, li in pairs:
        if distance > max_distance:
            break
        if ti in used_tags or li in used_lines:
            continue
        used_tags.add(ti)
        used_lines.add(li)
        result.append({"tag": tags[ti][0], "pos": tags[ti][1], "line": li})
    return result


def _round_pos(value):
    return round(value / POSITION_TOLERANCE) * POSITION_TOLERANCE


class FlatPatternView:
    """
    The flat pattern of one part (read from the DXF that SolidWorks exported), ready to be drawn on a
    PDF page with its overall length / width and the positions of the bend lines.
    """

    def __init__(self, paths, info=None, tags=None, labels=None, folded=None):
        # paths: list of {"points": [(x, y), ...], "bend": bool}
        self.paths = paths
        self.info = info or []
        self.tags = tags or []      # [{"tag": "A", "pos": (x, y), "line": index}, ...]
        self.labels = labels or []  # [("Fixed Face", (x, y)), ...]
        self.folded = folded or []  # [{"name": "Front", "paths": [...], "bbox": (x0, y0, x1, y1)}, ...] real mm
        self.scale_label = ""

    # =====================================================
    # LOADING
    # =====================================================

    @staticmethod
    def _is_bend(entity, doc):
        layer = str(entity.dxf.layer or "").upper()
        if any(word in layer for word in BEND_LAYER_WORDS):
            return True

        linetype = str(entity.dxf.get("linetype", "BYLAYER") or "BYLAYER").upper()
        if linetype in ("BYLAYER", "BYBLOCK", ""):
            try:
                linetype = str(doc.layers.get(entity.dxf.layer).dxf.linetype).upper()
            except Exception:
                linetype = "CONTINUOUS"
        return any(word in linetype for word in DASHED_WORDS)

    @classmethod
    def load(cls, dxf_path):
        """Returns (view_or_None, message)."""
        try:
            import ezdxf
            from ezdxf import path as ezpath
        except ImportError:
            return None, "ezdxf is not installed (pip install ezdxf)"

        try:
            doc = ezdxf.readfile(str(dxf_path))
        except Exception as exc:
            return None, f"cannot read {Path(dxf_path).name}: {exc}"

        paths = []
        layers = {}
        for entity in doc.modelspace():
            kind = entity.dxftype()
            if kind in ("TEXT", "MTEXT", "DIMENSION", "POINT", "INSERT", "HATCH"):
                continue
            try:
                geometry = ezpath.make_path(entity)
            except Exception:
                continue

            is_bend = cls._is_bend(entity, doc)
            layers[str(entity.dxf.layer)] = layers.get(str(entity.dxf.layer), 0) + 1

            for sub in geometry.sub_paths():
                points = [(v.x, v.y) for v in sub.flattening(0.05)]
                if len(points) >= 2:
                    paths.append({"points": points, "bend": is_bend})

        if not paths:
            return None, "the DXF has no drawable geometry"

        info = [f"layers: {layers}"]
        view = cls(paths, info)
        if not view.bend_lines():
            info.append("bend lines could not be identified (no layer containing 'BEND' and no dashed linetype)")
        return view, "OK"

    @classmethod
    def from_drawing_dxf(cls, dxf_path, json_path, tag_names=()):
        """
        The flat pattern view exactly as SolidWorks drew it in the temporary drawing: geometry from the
        exported drawing DXF, the bend lines, tag notes and view outlines from the JSON saved next to it.
        The same drawing may also hold the folded views (Front / Top / Right) far away from the flat one;
        the outlines in the JSON tell which geometry belongs to which view.
        Returns (view_or_None, message).
        """
        try:
            import ezdxf
            from ezdxf import path as ezpath
        except ImportError:
            return None, "ezdxf is not installed (pip install ezdxf)"

        try:
            data = json.loads(Path(json_path).read_text(encoding="utf-8"))
            doc = ezdxf.readfile(str(dxf_path))
        except Exception as exc:
            return None, f"cannot read the view data: {exc}"

        def positive(value):
            try:
                value = float(value)
                return value if value > 0 else 1.0
            except Exception:
                return 1.0

        flat_scale = positive(data.get("scale"))   # sheet size / real size

        def box_mm(outline, margin=8.0):
            try:
                x0, y0, x1, y1 = [float(v) * 1000.0 for v in outline][:4]
                return (x0 - margin, y0 - margin, x1 + margin, y1 + margin)
            except Exception:
                return None

        flat_box = box_mm(data.get("outline"))
        folded_defs = []
        for entry in data.get("folded_views", []) or []:
            box = box_mm(entry.get("outline"))
            if box is not None:
                folded_defs.append((str(entry.get("name") or "View"), box, positive(entry.get("scale"))))

        # 1. every drawable entity, in sheet millimetres
        flatten = 0.05 * min([flat_scale] + [sc for _n, _b, sc in folded_defs])
        raw = []
        for entity in doc.modelspace():
            kind = entity.dxftype()
            if kind not in GEOMETRY_TYPES:
                continue   # blocks (table, notes, centre marks) and text are not part of any shape
            try:
                geometry = ezpath.make_path(entity)
            except Exception:
                continue
            for sub in geometry.sub_paths():
                points = [(v.x, v.y) for v in sub.flattening(flatten)]
                if len(points) >= 2:
                    raw.append({"points": points, "kind": kind})

        # 2. which view does each piece of geometry belong to?
        def inside(point, box):
            return box[0] <= point[0] <= box[2] and box[1] <= point[1] <= box[3]

        flat_items = []
        folded_items = [[] for _ in folded_defs]
        for item in raw:
            xs = [p[0] for p in item["points"]]
            ys = [p[1] for p in item["points"]]
            centre = ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)
            if flat_box is None or inside(centre, flat_box):
                flat_items.append(item)
                continue
            for index, (_name, box, _sc) in enumerate(folded_defs):
                if inside(centre, box):
                    folded_items[index].append(item)
                    break

        paths, dxf_lines = [], []
        for item in flat_items:
            points = [(x / flat_scale, y / flat_scale) for x, y in item["points"]]
            paths.append({"points": points, "bend": False})
            if item["kind"] == "LINE" and len(points) == 2:
                dxf_lines.append((len(paths) - 1, points[0], points[1]))

        if not paths:
            return None, "the drawing DXF has no geometry for the flat pattern view"

        info = []
        probe = []
        for line in data.get("bend_lines", []):
            try:
                probe.append((
                    (line["start"][0] * 1000.0, line["start"][1] * 1000.0),
                    (line["end"][0] * 1000.0, line["end"][1] * 1000.0),
                ))
            except Exception:
                continue

        matched, message = match_bend_lines(probe, dxf_lines)
        bend_lines = []
        if matched is None:
            info.append(f"bend lines could not be identified - {message}")
        else:
            for index in sorted(matched):
                paths[index]["bend"] = True
            bend_lines = [(paths[i]["points"][0], paths[i]["points"][-1]) for i in sorted(matched)]
            info.append(f"{len(matched)} bend line(s) identified from SolidWorks")

        # the notes in the view: bend tags (A, B, ...) and other labels such as "Fixed Face"
        wanted = {str(t).strip() for t in tag_names}
        tag_notes, labels = [], []
        for note in data.get("notes", []):
            text = str(note.get("text") or "").strip()
            position = note.get("position")
            if not text or not position or len(position) < 2:
                continue
            at = (position[0] * 1000.0 / flat_scale, position[1] * 1000.0 / flat_scale)
            (tag_notes if text in wanted else labels).append((text, at))

        tags = []
        if bend_lines and tag_notes:
            tags = assign_tags(bend_lines, tag_notes, TAG_MAX_DISTANCE / flat_scale)
            info.append(f"tags placed: {', '.join(t['tag'] for t in tags)} of {len(tag_notes)}")
            if len(tags) != len(tag_notes) or len(tags) != len(bend_lines):
                info.append(
                    f"WARNING: {len(tag_notes)} tag(s), {len(bend_lines)} bend line(s), {len(tags)} paired - check this part"
                )
        elif tag_notes:
            info.append("tags not placed: bend lines were not identified")

        # 3. the folded views, in real millimetres
        folded = []
        for (name, _box, scale), items in zip(folded_defs, folded_items):
            if not items:
                continue
            fpaths = [
                {"points": [(x / scale, y / scale) for x, y in it["points"]], "kind": it["kind"]}
                for it in items
            ]
            xs = [p[0] for fp in fpaths for p in fp["points"]]
            ys = [p[1] for fp in fpaths for p in fp["points"]]
            folded.append({"name": name, "paths": fpaths, "bbox": (min(xs), min(ys), max(xs), max(ys))})
        if folded_defs:
            info.append(f"folded views found: {', '.join(f['name'] for f in folded) or 'none'}")

        return cls(paths, info, tags, labels, folded), "OK"

    # =====================================================
    # GEOMETRY
    # =====================================================

    def _outline_points(self):
        pts = [p for path in self.paths if not path["bend"] for p in path["points"]]
        return pts or [p for path in self.paths for p in path["points"]]

    def bbox(self):
        pts = self._outline_points()
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        return min(xs), min(ys), max(xs), max(ys)

    def bend_lines(self):
        """Straight bend lines as ((x1, y1), (x2, y2))."""
        lines = []
        for path in self.paths:
            if path["bend"]:
                lines.append((path["points"][0], path["points"][-1]))
        return lines

    def rotated(self):
        """The same view turned 90 degrees (x, y) -> (-y, x)."""
        def turn(p):
            return (-p[1], p[0])

        paths = [
            {"points": [turn(q) for q in p["points"]], "bend": p["bend"]}
            for p in self.paths
        ]
        tags = [dict(t, pos=turn(t["pos"])) for t in self.tags]
        labels = [(text, turn(pos)) for text, pos in self.labels]
        return FlatPatternView(paths, list(self.info), tags, labels)

    def bend_positions(self):
        """(x positions of vertical bend lines, y positions of horizontal bend lines, number of slanted ones)."""
        xs, ys, slanted = set(), set(), 0
        for (x1, y1), (x2, y2) in self.bend_lines():
            dx, dy = abs(x2 - x1), abs(y2 - y1)
            length = math.hypot(dx, dy)
            if length < 1e-6:
                continue
            if dx <= AXIS_TOLERANCE * length:
                xs.add(_round_pos((x1 + x2) / 2.0))
            elif dy <= AXIS_TOLERANCE * length:
                ys.add(_round_pos((y1 + y2) / 2.0))
            else:
                slanted += 1
        return sorted(xs), sorted(ys), slanted

    # =====================================================
    # DRAWING
    # =====================================================

    @staticmethod
    def _arrow(c, x, y, ux, uy, size=4.0):
        """Filled arrow head with its tip at (x, y), pointing along (ux, uy)."""
        px, py = -uy, ux
        p = c.beginPath()
        p.moveTo(x, y)
        p.lineTo(x - ux * size * 2.2 + px * size * 0.55, y - uy * size * 2.2 + py * size * 0.55)
        p.lineTo(x - ux * size * 2.2 - px * size * 0.55, y - uy * size * 2.2 - py * size * 0.55)
        p.close()
        c.drawPath(p, stroke=0, fill=1)

    def _hdim(self, c, xa, xb, y_dim, y_from, text, font, size):
        """Horizontal dimension between page x positions xa < xb, dimension line at y_dim."""
        sign = 1.0 if y_dim >= y_from else -1.0
        for x in (xa, xb):
            c.line(x, y_from + sign * 3.0, x, y_dim + sign * 3.0)
        c.line(xa, y_dim, xb, y_dim)
        if xb - xa > 9.0:
            self._arrow(c, xa, y_dim, 1, 0)
            self._arrow(c, xb, y_dim, -1, 0)
        c.setFont(font, size)
        c.drawCentredString((xa + xb) / 2.0, y_dim + (2.5 if sign > 0 else -size - 1.5), text)

    def _vdim(self, c, ya, yb, x_dim, x_from, text, font, size):
        """Vertical dimension between page y positions ya < yb, dimension line at x_dim."""
        sign = 1.0 if x_dim >= x_from else -1.0
        for y in (ya, yb):
            c.line(x_from + sign * 3.0, y, x_dim + sign * 3.0, y)
        c.line(x_dim, ya, x_dim, yb)
        if yb - ya > 9.0:
            self._arrow(c, x_dim, ya, 0, 1)
            self._arrow(c, x_dim, yb, 0, -1)
        c.saveState()
        c.setFont(font, size)
        c.translate(x_dim + (-2.5 if sign > 0 else size + 1.5), (ya + yb) / 2.0)
        c.rotate(90)
        c.drawCentredString(0, 0, text)
        c.restoreState()

    # =====================================================
    # FOLDED VIEWS (the part after bending)
    # =====================================================

    def folded_selection(self, thickness=None):
        """
        The folded views worth showing, biggest first (at most 3).
        A view that is only a thin strip (the part seen edge-on) is dropped: it tells nothing the others do not.
        """
        minimum = max(4.0, 3.0 * float(thickness)) if thickness else 4.0
        chosen = []
        for view in self.folded:
            x0, y0, x1, y1 = view["bbox"]
            if min(x1 - x0, y1 - y0) >= minimum:
                chosen.append(view)
        chosen.sort(key=lambda v: -((v["bbox"][2] - v["bbox"][0]) * (v["bbox"][3] - v["bbox"][1])))
        return chosen[:3]

    @staticmethod
    def _step_marks(view):
        """
        Positions of the corners of a simple profile (straight edges only), for step dimensions such as
        21.5 / 24.5 in the sample. None when the outline is too detailed to dimension this way.
        """
        xs, ys = set(), set()
        for path in view["paths"]:
            if path["kind"] == "LINE":
                corners = (path["points"][0], path["points"][-1])
            elif path["kind"] in ("LWPOLYLINE", "POLYLINE"):
                corners = path["points"]   # a polyline with arcs has many points -> too detailed -> no step dimensions
            else:
                continue                   # circles, arcs and splines are details, not profile corners
            for x, y in corners:
                xs.add(round(x, 1))
                ys.add(round(y, 1))
        x0, y0, x1, y1 = view["bbox"]
        xs |= {round(x0, 1), round(x1, 1)}
        ys |= {round(y0, 1), round(y1, 1)}
        xm, ym = sorted(xs), sorted(ys)
        return (xm if 3 <= len(xm) <= 7 else None), (ym if 3 <= len(ym) <= 7 else None)

    def draw_folded(self, c, area, thickness=None, font="Helvetica", font_size=8.5):
        """
        Draws the chosen folded views side by side in area = (x0, y0, x1, y1), all at the same scale, each with
        its overall size after bending (and the step dimensions when the profile is simple).
        Returns a description, or None when nothing was drawn.
        """
        views = self.folded_selection(thickness)
        if not views:
            return None

        ax0, ay0, ax1, ay1 = area
        aw, ah = ax1 - ax0, ay1 - ay0
        marks = [self._step_marks(v) for v in views]
        any_x = any(m[0] for m in marks)
        any_y = any(m[1] for m in marks)

        top = 30.0
        bottom = 70.0 if any_x else 40.0
        gap = 78.0 if any_y else 52.0
        side = 44.0 if any_y else 20.0
        cap = 5.0 * PT_PER_MM

        widths = [v["bbox"][2] - v["bbox"][0] for v in views]
        heights = [v["bbox"][3] - v["bbox"][1] for v in views]
        usable_h = ah - top - bottom
        n = len(views)

        # one scale for all views ...
        k_common = max(min((aw - 2 * side - gap * (n - 1)) / sum(widths), usable_h / max(heights), cap), 1e-6)
        ks = [k_common] * n

        # ... except views that would come out tiny (a profile next to a 2 m long face view): they get their own,
        # larger scale, like a detail view, and their scale is written under them
        extent = [max(w, h) for w, h in zip(widths, heights)]
        big = [i for i in range(n) if extent[i] * k_common >= 80.0]
        if big and len(big) < n:
            small = [i for i in range(n) if i not in big]
            for i in small:
                ks[i] = max(min(110.0 / extent[i], usable_h / heights[i], cap), 1e-6)
            used = sum(widths[i] * ks[i] for i in small) + gap * (n - 1)
            k_big = min(
                (aw - 2 * side - used) / sum(widths[i] for i in big),
                usable_h / max(heights[i] for i in big),
                cap,
            )
            if k_big > 0:
                for i in big:
                    ks[i] = k_big
            else:
                ks = [k_common] * n

        def scale_text(k):
            ratio = k / PT_PER_MM
            return f"{ratio:.3g}:1" if ratio >= 1 else f"1:{1.0 / ratio:.3g}"

        total = sum(w * kk for w, kk in zip(widths, ks)) + gap * (n - 1)
        cursor = ax0 + side + (aw - 2 * side - total) / 2.0
        base_y = ay0 + bottom + (usable_h - max(h * kk for h, kk in zip(heights, ks))) / 2.0

        c.saveState()
        for view, (xmarks, ymarks), k in zip(views, marks, ks):
            x0, y0, x1, y1 = view["bbox"]
            ox = cursor
            cursor += (x1 - x0) * k + gap

            def px(x, ox=ox, x0=x0):
                return ox + (x - x0) * k

            def py(y, y0=y0):
                return base_y + (y - y0) * k

            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(0.7)
            c.setDash()
            for path in view["paths"]:
                pts = path["points"]
                p = c.beginPath()
                p.moveTo(px(pts[0][0]), py(pts[0][1]))
                for x, y in pts[1:]:
                    p.lineTo(px(x), py(y))
                c.drawPath(p, stroke=1, fill=0)

            gray = 0.25
            c.setStrokeColorRGB(gray, gray, gray)
            c.setFillColorRGB(gray, gray, gray)
            c.setLineWidth(0.4)
            left_x, right_x = px(x0), px(x1)
            bottom_y, top_y = py(y0), py(y1)

            self._hdim(c, left_x, right_x, top_y + 20.0, top_y, f"{x1 - x0:.1f}", font, font_size)
            self._vdim(c, bottom_y, top_y, right_x + 20.0, right_x, f"{y1 - y0:.1f}", font, font_size)

            levels_x = []
            if xmarks:
                levels_x = self._levels(xmarks, k, font, font_size)
                for i in range(len(xmarks) - 1):
                    self._hdim(c, px(xmarks[i]), px(xmarks[i + 1]), bottom_y - (14.0 + 13.0 * levels_x[i]), bottom_y,
                               f"{xmarks[i + 1] - xmarks[i]:.1f}", font, font_size)
            if ymarks:
                levels = self._levels(ymarks, k, font, font_size)
                for i in range(len(ymarks) - 1):
                    self._vdim(c, py(ymarks[i]), py(ymarks[i + 1]), left_x - (14.0 + 13.0 * levels[i]), left_x,
                               f"{ymarks[i + 1] - ymarks[i]:.1f}", font, font_size)

            # the name and scale of the view, right under its dimensions
            rows = (max(levels_x) + 1) if xmarks else 0
            c.setFillColorRGB(0, 0, 0)
            c.setFont(font, 7)
            c.drawString(left_x, bottom_y - (30.0 + 13.0 * rows), f"{view['name'].upper()} VIEW   SCALE {scale_text(k)}")
        c.restoreState()

        return ", ".join(
            f"{v['name']} {w:.1f} x {h:.1f} ({scale_text(kk)})" for v, w, h, kk in zip(views, widths, heights, ks)
        )

    @staticmethod
    def _chain_marks(low, high, positions):
        inside = [p for p in positions if low + 1e-3 < p < high - 1e-3]
        return [low] + inside + [high] if positions else None

    @staticmethod
    def _levels(marks, k, font, size):
        """
        Row (0, 1, 2, ...) for each dimension of a chain, so that neighbouring numbers never overlap.
        A number goes to the first row where it does not touch the previous number of that row.
        """
        from reportlab.pdfbase.pdfmetrics import stringWidth

        levels, last_end = [], []
        for a, b in zip(marks, marks[1:]):
            text_width = stringWidth(f"{b - a:.1f}", font, size) + 3.0
            centre = (a + b) / 2.0 * k
            start, end = centre - text_width / 2.0, centre + text_width / 2.0
            level = 0
            while True:
                if level >= len(last_end):
                    last_end.append(-1e12)
                if start >= last_end[level]:
                    last_end[level] = end
                    break
                level += 1
            levels.append(level)
        return levels

    def _plan(self, area_size, font, size):
        """Scale and margins for this orientation. The chain margins grow with the rows they need."""
        aw, ah = area_size
        xs, ys, _ = self.bend_positions()
        x0, y0, x1, y1 = self.bbox()
        width, height = max(x1 - x0, 1e-6), max(y1 - y0, 1e-6)
        xmarks = self._chain_marks(x0, x1, xs)
        ymarks = self._chain_marks(y0, y1, ys)
        top, right = 30.0, 30.0
        left = 44.0 if ymarks else 14.0
        bottom = 44.0 if xmarks else 14.0
        k = 1.0

        for _ in range(4):
            k = min((aw - left - right) / width, (ah - top - bottom) / height, 5.0 * PT_PER_MM)
            k = max(k, 1e-6)
            new_bottom = 22.0 + 13.0 * (max(self._levels(xmarks, k, font, size)) + 1) if xmarks else 14.0
            new_left = 22.0 + 13.0 * (max(self._levels(ymarks, k, font, size)) + 1) if ymarks else 14.0
            if abs(new_bottom - bottom) < 0.5 and abs(new_left - left) < 0.5:
                break
            bottom, left = new_bottom, new_left
        return k, (left, bottom, top, right)

    def draw(self, c, area, font="Helvetica", font_size=8.5):
        """
        Draws the view (with dimensions) centred in area = (x0, y0, x1, y1) in PDF points.
        Returns a short description of what was drawn.
        """
        ax0, ay0, ax1, ay1 = area
        aw, ah = ax1 - ax0, ay1 - ay0

        view = self
        k, margins = self._plan((aw, ah), font, font_size)
        turned = self.rotated()
        k_turned, margins_turned = turned._plan((aw, ah), font, font_size)
        if k_turned > k * 1.2:
            view, k, margins = turned, k_turned, margins_turned

        left, bottom, top, right = margins
        x0, y0, x1, y1 = view.bbox()
        width, height = x1 - x0, y1 - y0
        drawn_w, drawn_h = width * k, height * k

        # centre the drawing (without its dimension margins) inside the free area
        ox = ax0 + left + (aw - left - right - drawn_w) / 2.0
        oy = ay0 + bottom + (ah - top - bottom - drawn_h) / 2.0

        def px(x):
            return ox + (x - x0) * k

        def py(y):
            return oy + (y - y0) * k

        # ---- geometry ----
        c.saveState()
        c.setStrokeColorRGB(0, 0, 0)
        for path in view.paths:
            pts = path["points"]
            c.setLineWidth(0.35 if path["bend"] else 0.7)
            c.setDash(3, 2) if path["bend"] else c.setDash()
            p = c.beginPath()
            p.moveTo(px(pts[0][0]), py(pts[0][1]))
            for x, y in pts[1:]:
                p.lineTo(px(x), py(y))
            c.drawPath(p, stroke=1, fill=0)
        c.setDash()

        # ---- bend tags and labels, where SolidWorks put them ----
        c.setFillColorRGB(0, 0, 0)
        for tag in view.tags:
            c.setFont(font + "-Bold", 10)
            c.drawCentredString(px(tag["pos"][0]), py(tag["pos"][1]) - 8.0, tag["tag"])
        for text, pos in view.labels:
            c.setFont(font, 9)
            c.drawCentredString(px(pos[0]), py(pos[1]) - 3.0, text)

        # ---- dimensions ----
        gray = 0.25
        c.setStrokeColorRGB(gray, gray, gray)
        c.setFillColorRGB(gray, gray, gray)
        c.setLineWidth(0.4)

        left_x, right_x = px(x0), px(x1)
        bottom_y, top_y = py(y0), py(y1)

        self._hdim(c, left_x, right_x, top_y + 22.0, top_y, f"{width:.1f}", font, font_size)
        self._vdim(c, bottom_y, top_y, right_x + 22.0, right_x, f"{height:.1f}", font, font_size)

        xs, ys, slanted = view.bend_positions()
        xmarks = self._chain_marks(x0, x1, xs)
        ymarks = self._chain_marks(y0, y1, ys)

        if xmarks:
            levels = self._levels(xmarks, k, font, font_size)
            for i in range(len(xmarks) - 1):
                self._hdim(c, px(xmarks[i]), px(xmarks[i + 1]), bottom_y - (16.0 + 13.0 * levels[i]), bottom_y,
                           f"{xmarks[i + 1] - xmarks[i]:.1f}", font, font_size)
        if ymarks:
            levels = self._levels(ymarks, k, font, font_size)
            for i in range(len(ymarks) - 1):
                self._vdim(c, py(ymarks[i]), py(ymarks[i + 1]), left_x - (16.0 + 13.0 * levels[i]), left_x,
                           f"{ymarks[i + 1] - ymarks[i]:.1f}", font, font_size)
        c.restoreState()

        ratio = k / PT_PER_MM
        self.scale_label = f"{ratio:.3g}:1" if ratio >= 1 else f"1:{1.0 / ratio:.3g}"
        return (
            f"{width:.1f} x {height:.1f} mm, scale {self.scale_label}, "
            f"{len(xs)} vertical / {len(ys)} horizontal / {slanted} slanted bend line(s)"
            + (f", {len(view.tags)} tag(s)" if view.tags else "")
            + (", turned 90 degrees" if view is not self else "")
        )