"""COM-free geometry for flat-pattern bend dimensions.

Everything in this module is plain Python so it can be unit-tested without SolidWorks.
sw_drawings.py feeds it numbers read through the SolidWorks API.

Coordinate systems used by sw_drawings.py
-----------------------------------------
* model  : part (model) space, metres. Visible edges / vertices of a drawing view report their points here.
* sketch : local 2D frame of the flat pattern's bend-line sketch, metres. IView.GetBendLines() returns sketch
           segments whose GetStartPoint2/GetEndPoint2 X/Y/Z are in THIS frame (ISketch.ModelToSketchTransform
           maps model -> sketch; its inverse maps sketch -> model).
* frame  : the measuring frame = the sketch frame of the first bend line. It is true size (no drawing scale) and
           lies in the plane of the flat sheet, so bend-to-edge distances are measured here.
* sheet  : drawing sheet, metres on paper (includes view scale, rotation and position). Used only to place text,
           pair bend tags and avoid collisions - never to compute a dimension value.

MathTransform convention
------------------------
ArrayData = [R00..R22 (9), Tx, Ty, Tz, scale, 0, 0, 0]. The documented SolidWorks form treats points as row
vectors: p' = scale * (p . R) + T   (rows of R are the images of the x/y/z axes). The convention actually used
by the running SolidWorks is verified at runtime with IMathPoint.MultiplyTransform (see detect_convention).
"""

import math

EPS = 1e-12
DEFAULT_TOL = 1e-5                     # 0.01 mm in metres
DEFAULT_ANGLE_TOL = math.radians(0.5)  # parallelism tolerance for bend line / edge


# =====================================================================
# 2D vector helpers
# =====================================================================

def v_sub(a, b):
    return (a[0] - b[0], a[1] - b[1])


def v_add(a, b):
    return (a[0] + b[0], a[1] + b[1])


def v_mul(a, k):
    return (a[0] * k, a[1] * k)


def v_dot(a, b):
    return a[0] * b[0] + a[1] * b[1]


def v_cross(a, b):
    return a[0] * b[1] - a[1] * b[0]


def v_len(a):
    return math.hypot(a[0], a[1])


def v_unit(a):
    length = v_len(a)
    if length < EPS:
        raise ValueError("zero-length vector")
    return (a[0] / length, a[1] / length)


def v_perp(a):
    return (-a[1], a[0])


def v_mid(a, b):
    return ((a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0)


def xy(p):
    return (float(p[0]), float(p[1]))


# =====================================================================
# MathTransform
# =====================================================================

def _mat_mul(a, b):
    return [[sum(a[i][k] * b[k][j] for k in range(3)) for j in range(3)] for i in range(3)]


def _mat_inv(m):
    (a, b, c), (d, e, f), (g, h, i) = m
    det = a * (e * i - f * h) - b * (d * i - f * g) + c * (d * h - e * g)
    if abs(det) < EPS:
        raise ValueError("singular rotation matrix")
    return [
        [(e * i - f * h) / det, (c * h - b * i) / det, (b * f - c * e) / det],
        [(f * g - d * i) / det, (a * i - c * g) / det, (c * d - a * f) / det],
        [(d * h - e * g) / det, (b * g - a * h) / det, (a * e - b * d) / det],
    ]


def _transpose(m):
    return [[m[j][i] for j in range(3)] for i in range(3)]


class Xform:
    """SolidWorks MathTransform in row-vector form: p' = s * (p . R) + T."""

    __slots__ = ("R", "T", "s")

    def __init__(self, R=None, T=(0.0, 0.0, 0.0), s=1.0):
        R = R if R is not None else [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]]
        self.R = [[float(v) for v in row] for row in R]
        self.T = tuple(float(v) for v in T)
        self.s = float(s)

    @classmethod
    def identity(cls):
        return cls()

    @classmethod
    def from_array(cls, data, convention="row"):
        d = [float(v) for v in data]
        if len(d) < 13:
            raise ValueError(f"MathTransform ArrayData needs 13+ values, got {len(d)}")
        R = [d[0:3], d[3:6], d[6:9]]
        if convention == "column":
            R = _transpose(R)
        elif convention != "row":
            raise ValueError(f"unknown convention {convention!r}")
        return cls(R, d[9:12], d[12] if d[12] else 1.0)

    def to_array(self):
        r = self.R
        return tuple(r[0] + r[1] + r[2]) + self.T + (self.s, 0.0, 0.0, 0.0)

    def apply(self, p):
        x, y = float(p[0]), float(p[1])
        z = float(p[2]) if len(p) > 2 else 0.0
        r, s, t = self.R, self.s, self.T
        return tuple(s * (x * r[0][i] + y * r[1][i] + z * r[2][i]) + t[i] for i in range(3))

    def inverse(self):
        inv = _mat_inv(self.R)
        s = self.s
        t = self.T
        t_new = tuple(-(t[0] * inv[0][i] + t[1] * inv[1][i] + t[2] * inv[2][i]) / s for i in range(3))
        return Xform(inv, t_new, 1.0 / s)

    def then(self, other):
        """Transform that applies self first, then other."""
        r = _mat_mul(self.R, other.R)
        t1 = self.T
        t = tuple(other.s * (t1[0] * other.R[0][i] + t1[1] * other.R[1][i] + t1[2] * other.R[2][i]) + other.T[i]
                  for i in range(3))
        return Xform(r, t, self.s * other.s)

    def normal_z(self, normal):
        """Rotate a direction (no scale / translation) and return its unit z component in the target frame."""
        x, y, z = (float(v) for v in normal)
        r = self.R
        out = [x * r[0][i] + y * r[1][i] + z * r[2][i] for i in range(3)]
        length = math.sqrt(sum(v * v for v in out))
        return out[2] / length if length > EPS else 0.0


def detect_convention(data, probe_in, probe_out, tol=1e-7):
    """Return "row" or "column" - whichever reproduces what SolidWorks computed for probe_in - or None."""
    for convention in ("row", "column"):
        got = Xform.from_array(data, convention).apply(probe_in)
        scale = 1.0 + max(abs(float(v)) for v in probe_out[:3])
        if all(abs(got[i] - float(probe_out[i])) <= tol * scale for i in range(3)):
            return convention
    return None


# =====================================================================
# Segment / polygon helpers
# =====================================================================

def point_segment_distance(p, a, b):
    d = v_sub(b, a)
    length_sq = v_dot(d, d)
    if length_sq < EPS:
        return v_len(v_sub(p, a))
    t = max(0.0, min(1.0, v_dot(v_sub(p, a), d) / length_sq))
    return v_len(v_sub(p, v_add(a, v_mul(d, t))))


def segments_intersect(p1, p2, q1, q2, tol=1e-12):
    """True if the closed segments p1-p2 and q1-q2 touch or cross."""
    def orient(a, b, c):
        value = v_cross(v_sub(b, a), v_sub(c, a))
        return 0 if abs(value) <= tol else (1 if value > 0 else -1)

    def on_segment(a, b, c):
        return (min(a[0], b[0]) - tol <= c[0] <= max(a[0], b[0]) + tol and
                min(a[1], b[1]) - tol <= c[1] <= max(a[1], b[1]) + tol)

    o1, o2 = orient(p1, p2, q1), orient(p1, p2, q2)
    o3, o4 = orient(q1, q2, p1), orient(q1, q2, p2)
    if o1 != o2 and o3 != o4:
        return True
    if o1 == 0 and on_segment(p1, p2, q1):
        return True
    if o2 == 0 and on_segment(p1, p2, q2):
        return True
    if o3 == 0 and on_segment(q1, q2, p1):
        return True
    if o4 == 0 and on_segment(q1, q2, p2):
        return True
    return False


def polygon_area(points):
    area = 0.0
    for i, a in enumerate(points):
        b = points[(i + 1) % len(points)]
        area += a[0] * b[1] - b[0] * a[1]
    return area / 2.0


def point_in_polygon(p, polygon):
    inside = False
    n = len(polygon)
    for i in range(n):
        a, b = polygon[i], polygon[(i + 1) % n]
        if (a[1] > p[1]) != (b[1] > p[1]):
            x = a[0] + (p[1] - a[1]) * (b[0] - a[0]) / (b[1] - a[1])
            if x > p[0]:
                inside = not inside
    return inside


def same_segment(a0, a1, b0, b1, tol):
    d = lambda p, q: v_len(v_sub(p, q))  # noqa: E731
    return (d(a0, b0) <= tol and d(a1, b1) <= tol) or (d(a0, b1) <= tol and d(a1, b0) <= tol)


class PointIndex:
    """Clusters points closer than tol into one node id."""

    def __init__(self, tol):
        self.tol = tol
        self.points = []
        self.grid = {}

    def add(self, p):
        cell = (int(math.floor(p[0] / self.tol)), int(math.floor(p[1] / self.tol)))
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for idx in self.grid.get((cell[0] + dx, cell[1] + dy), ()):
                    if v_len(v_sub(self.points[idx], p)) <= self.tol:
                        return idx
        self.points.append(p)
        self.grid.setdefault(cell, []).append(len(self.points) - 1)
        return len(self.points) - 1


def build_loops(segments, tol=DEFAULT_TOL * 5):
    """Group 2D segments [(p0, p1), ...] into connected chains.

    Returns [{"segments": [i...], "closed": bool, "polygon": [pts] | None, "area": float, "bbox": (...)}].
    A chain is closed when every node has exactly two segments.
    """
    index = PointIndex(tol)
    ends = []
    for p0, p1 in segments:
        ends.append((index.add(xy(p0)), index.add(xy(p1))))

    parent = list(range(len(index.points)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for a, b in ends:
        parent[find(a)] = find(b)

    groups = {}
    for seg_idx, (a, b) in enumerate(ends):
        if a == b:
            continue  # degenerate after clustering
        groups.setdefault(find(a), []).append(seg_idx)

    loops = []
    for seg_ids in groups.values():
        degree = {}
        adjacency = {}
        for s in seg_ids:
            a, b = ends[s]
            degree[a] = degree.get(a, 0) + 1
            degree[b] = degree.get(b, 0) + 1
            adjacency.setdefault(a, []).append((s, b))
            adjacency.setdefault(b, []).append((s, a))
        closed = all(v == 2 for v in degree.values()) and len(seg_ids) >= 2
        pts = [index.points[n] for n in degree]
        bbox = (min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts))
        polygon = None
        area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
        if closed:
            start = next(iter(degree))
            polygon, used, node = [], set(), start
            while True:
                polygon.append(index.points[node])
                nxt = next(((s, other) for s, other in adjacency[node] if s not in used), None)
                if nxt is None:
                    break
                used.add(nxt[0])
                node = nxt[1]
                if node == start:
                    break
            if len(used) == len(seg_ids) and len(polygon) >= 3:
                area = abs(polygon_area(polygon))
            else:
                closed, polygon = False, None
        loops.append({"segments": sorted(seg_ids), "closed": closed, "polygon": polygon,
                      "area": area, "bbox": bbox})
    return loops


def classify_outer(segments, outer_hint=None, tol=DEFAULT_TOL * 5):
    """Decide which segments form the outer boundary of the flat sheet.

    segments   : [(p0, p1), ...] in the measuring frame (lines and arc/spline chords).
    outer_hint : optional [(p0, p1), ...] from the SolidWorks face loop flagged ILoop2.IsOuter() - preferred.
    Returns (set_of_indices, info) with info = {"method", "closed", "polygon", "warning"}.
    """
    info = {"method": None, "closed": False, "polygon": None, "warning": None}
    if outer_hint:
        chosen = {i for i, (p0, p1) in enumerate(segments)
                  if any(same_segment(xy(p0), xy(p1), xy(h0), xy(h1), tol) for h0, h1 in outer_hint)}
        hint_loops = [lp for lp in build_loops(outer_hint, tol) if lp["closed"]]
        if chosen:
            info["method"] = "face outer loop (ILoop2.IsOuter)"
            if hint_loops:
                best = max(hint_loops, key=lambda lp: lp["area"])
                info["closed"], info["polygon"] = True, best["polygon"]
            else:
                info["warning"] = "outer loop from SolidWorks is not closed in 2D"
            return chosen, info

    loops = build_loops(segments, tol)
    if not loops:
        info["warning"] = "no edges to build a boundary from"
        return set(), info
    closed = [lp for lp in loops if lp["closed"]]
    if closed:
        best = max(closed, key=lambda lp: lp["area"])
        info.update(method="largest closed edge loop", closed=True, polygon=best["polygon"])
    else:
        best = max(loops, key=lambda lp: lp["area"])
        info.update(method="largest edge chain (NOT closed)",
                    warning="outer boundary could not be closed - results need manual review")
    return set(best["segments"]), info


# =====================================================================
# Bend -> outer edge measurement
# =====================================================================

def _canonical_dir(u):
    """Direction with a fixed sign so that parallel lines get the same canonical normal."""
    if u[0] < -EPS or (abs(u[0]) <= EPS and u[1] < 0):
        return (-u[0], -u[1])
    return u


def measure_bends(bends, edges, outer, polygon=None, tol=DEFAULT_TOL, angle_tol=DEFAULT_ANGLE_TOL,
                  merge_collinear=True):
    """Distance from every bend line to the nearest suitable OUTER edge, measured in the true-size frame.

    bends : [{"id", "p0", "p1"}]                       bend line end points (frame, metres)
    edges : [{"id", "p0", "p1", "is_line"}]           edge end points (frame, metres)
    outer : set of indexes into edges that form the outer boundary
    A candidate edge must be a straight outer edge, parallel to the bend, overlap the bend line along its
    direction, and be reachable without crossing another outer edge (so edges beyond a notch are rejected).
    Returns one result dict per bend (same order); "status" is one of
    ok | no_reference_edge | invalid | outside_boundary | duplicate | merged.
    """
    sin_tol = math.sin(angle_tol)
    outer_ids = sorted(outer)
    results = []

    for bend in bends:
        res = {"id": bend["id"], "p0": xy(bend["p0"]), "p1": xy(bend["p1"]), "status": None,
               "value": None, "edge": None, "signed": None, "overlap": None, "tie": False,
               "rejected": [], "reason": ""}
        results.append(res)
        p0, p1 = res["p0"], res["p1"]
        length = v_len(v_sub(p1, p0))
        if length < tol:
            res.update(status="invalid", reason="bend line has zero length")
            continue
        u = v_unit(v_sub(p1, p0))
        n = v_perp(u)
        res["length"] = length

        if polygon and not point_in_polygon(v_mid(p0, p1), polygon):
            if min(point_segment_distance(v_mid(p0, p1), polygon[i], polygon[(i + 1) % len(polygon)])
                   for i in range(len(polygon))) > tol:
                res.update(status="outside_boundary",
                           reason="bend line midpoint is outside the outer boundary (coordinate mismatch?)")
                continue

        candidates = []
        for i in outer_ids:
            e = edges[i]
            if not e.get("is_line", True):
                continue
            e0, e1 = xy(e["p0"]), xy(e["p1"])
            le = v_len(v_sub(e1, e0))
            if le < tol:
                continue
            if abs(v_cross(u, v_mul(v_sub(e1, e0), 1.0 / le))) > sin_tol:
                continue
            s = (v_dot(v_sub(e0, p0), n) + v_dot(v_sub(e1, p0), n)) / 2.0
            if abs(s) <= tol:
                continue  # collinear with the bend line
            t0, t1 = v_dot(v_sub(e0, p0), u), v_dot(v_sub(e1, p0), u)
            lo, hi = max(0.0, min(t0, t1)), min(length, max(t0, t1))
            if hi - lo <= tol:
                res["rejected"].append({"edge": i, "distance": abs(s), "why": "does not overlap the bend line"})
                continue
            tm = (lo + hi) / 2.0
            sign = 1.0 if s > 0 else -1.0
            start = v_add(v_add(p0, v_mul(u, tm)), v_mul(n, sign * 2 * tol))
            end = v_add(v_add(p0, v_mul(u, tm)), v_mul(n, s - sign * 2 * tol))
            blocker = next((j for j in outer_ids if j != i and
                            segments_intersect(start, end, xy(edges[j]["p0"]), xy(edges[j]["p1"]))), None)
            if blocker is not None:
                res["rejected"].append({"edge": i, "distance": abs(s),
                                        "why": f"another outer edge ({blocker}) lies in between"})
                continue
            candidates.append({"edge": i, "signed": s, "distance": abs(s), "overlap": (lo, hi)})

        if not candidates:
            res.update(status="no_reference_edge",
                       reason="no straight outer edge parallel to and facing this bend line")
            continue
        candidates.sort(key=lambda c: (c["distance"], -c["signed"]))
        best = candidates[0]
        other_side = [c for c in candidates if (c["signed"] > 0) != (best["signed"] > 0)]
        res["tie"] = bool(other_side and abs(other_side[0]["distance"] - best["distance"]) <= tol)
        res.update(status="ok", value=best["distance"], edge=best["edge"], signed=best["signed"],
                   overlap=best["overlap"])

    # ---- identity: exact duplicates and collinear bends sharing the same reference edge ----
    seen = []
    for res in results:
        if res["status"] != "ok":
            continue
        u = _canonical_dir(v_unit(v_sub(res["p1"], res["p0"])))
        n = v_perp(u)
        offset = v_dot(res["p0"], n)
        for prior in seen:
            if not same_segment(res["p0"], res["p1"], prior["p0"], prior["p1"], tol):
                if not merge_collinear:
                    continue
                pu = _canonical_dir(v_unit(v_sub(prior["p1"], prior["p0"])))
                if (abs(v_cross(u, pu)) > sin_tol or abs(offset - v_dot(prior["p0"], n)) > tol
                        or prior["edge"] != res["edge"]):
                    continue
                res.update(status="merged", merged_into=prior["id"],
                           reason=f"collinear with bend {prior['id']} and measured to the same edge")
            else:
                res.update(status="duplicate", merged_into=prior["id"],
                           reason=f"same geometry as bend {prior['id']}")
            prior.setdefault("shared_with", []).append(res["id"])
            break
        else:
            seen.append(res)
    return results


# =====================================================================
# Bend tags
# =====================================================================

def assign_tags(bend_lines, notes, tag_names, max_distance):
    """Pair bend tag notes with bend lines, nearest pairs first, one tag per line.

    bend_lines : {bend_id: (a, b)} in sheet coordinates
    notes      : [(text, (x, y))] in sheet coordinates
    tag_names  : texts that are bend tags (from the bend table); empty = accept every note
    Returns ({bend_id: tag}, [unassigned tag texts], {bend_id: tag} of AMBIGUOUS pairs).
    A pair is ambiguous when another bend line is within (own distance + D + 0.5 mm) of the tag, where D is the
    largest tag-to-nearest-line distance seen in this view (how far SolidWorks puts tags from their lines):
    under that same offset the tag could equally belong to the other line. Ambiguous pairs are left out of
    the first dict and returned in the third, so a tag is never silently given to the wrong bend.
    """
    wanted = {str(t).strip() for t in tag_names if str(t).strip()}
    tags = [(str(text).strip(), xy(pos)) for text, pos in notes
            if str(text).strip() and (not wanted or str(text).strip() in wanted)]
    pairs = sorted((point_segment_distance(pos, a, b), ti, bid)
                   for ti, (_t, pos) in enumerate(tags) for bid, (a, b) in bend_lines.items())
    used_tags, used_lines, result = set(), set(), {}
    for distance, ti, bid in pairs:
        if distance > max_distance:
            break
        if ti in used_tags or bid in used_lines:
            continue
        used_tags.add(ti)
        used_lines.add(bid)
        result[bid] = (tags[ti][0], distance, ti)
    ambiguous = {}
    offset = max((min(point_segment_distance(pos, a, b) for a, b in bend_lines.values())
                  for _t, pos in tags), default=0.0) if bend_lines else 0.0
    for bid, (text, own, ti) in list(result.items()):
        pos = tags[ti][1]
        others = [point_segment_distance(pos, a, b) for other, (a, b) in bend_lines.items() if other != bid]
        if others and min(others) <= own + offset + 0.0005:
            ambiguous[bid] = text
            del result[bid]
    unassigned = [tags[i][0] for i in range(len(tags)) if i not in used_tags]
    return {bid: v[0] for bid, v in result.items()}, unassigned, ambiguous


# =====================================================================
# Annotation placement
# =====================================================================

def _rect_overlap(a, b, gap=0.0):
    return a[0] < b[2] + gap and a[2] > b[0] - gap and a[1] < b[3] + gap and a[3] > b[1] - gap


def _segment_hits_rect(p, q, r):
    if (min(p[0], q[0]) > r[2] or max(p[0], q[0]) < r[0] or
            min(p[1], q[1]) > r[3] or max(p[1], q[1]) < r[1]):
        return False
    if r[0] <= p[0] <= r[2] and r[1] <= p[1] <= r[3]:
        return True
    corners = [(r[0], r[1]), (r[2], r[1]), (r[2], r[3]), (r[0], r[3])]
    return any(segments_intersect(p, q, corners[i], corners[(i + 1) % 4]) for i in range(4))


def _oriented_box(center, axis_a, half_a, axis_b, half_b):
    pts = [v_add(center, v_add(v_mul(axis_a, sa * half_a), v_mul(axis_b, sb * half_b)))
           for sa in (-1, 1) for sb in (-1, 1)]
    return (min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts))


def placement_direction(u):
    """Which end of the bend line the dimension goes to: prefer below the part, then to its right."""
    best = None
    for sgn in (1.0, -1.0):
        w = (u[0] * sgn, u[1] * sgn)
        score = max(-w[1], w[0])
        key = (round(score, 9), w[0])  # tie -> prefer the right-hand side
        if best is None or key > best[0]:
            best = (key, w)
    return best[1]


def side_name(w):
    if abs(w[0]) >= abs(w[1]):
        return "right" if w[0] > 0 else "left"
    return "top" if w[1] > 0 else "bottom"


def plan_placements(dims, part_points, offset, step, text_w, text_h, max_levels=12, gap=0.0005,
                    bounds=None, obstacles=()):
    """Deterministic, collision-aware placement of bend dimensions in sheet coordinates.

    dims        : ordered [{"key", "b0", "b1", "e0", "e1"}] - bend line and reference edge (sheet coords)
    part_points : sheet points of the flat pattern outline (dimension lines go outside it)
    bounds      : (xmin, ymin, xmax, ymax) drawable area, or None for unlimited (used to measure space needs)
    obstacles   : [(xmin, ymin, xmax, ymax)] areas text must not cover (bend tags, other dimensions, table...)
    The dimension line runs perpendicular to the bend through a point beyond the part along the bend direction;
    text sits on it between the bend line and the edge. Each dimension takes the lowest level (distance from the
    part) where nothing collides:
      * its text box vs. placed text boxes, obstacles, and placed dimension / extension lines,
      * its dimension and extension lines vs. placed text boxes,
      * its dimension line vs. placed dimension lines (two dimensions stacked on the same line).
    Extension lines crossing other dimension lines are normal in stacked dimensioning and are allowed.
    Returns (placements, needs) where needs = {"left"|"right"|"top"|"bottom": metres beyond the part bbox}.
    """
    bbox = (min(p[0] for p in part_points), min(p[1] for p in part_points),
            max(p[0] for p in part_points), max(p[1] for p in part_points))
    placed_rects, placed_lines, placed_dim_lines, placements = [], [], [], []
    obstacles = [tuple(o) for o in obstacles]

    for dim in dims:
        b0, b1, e0, e1 = (xy(dim[k]) for k in ("b0", "b1", "e0", "e1"))
        u = v_unit(v_sub(b1, b0))
        w = placement_direction(u)
        n0 = v_perp(u)
        s = v_dot(v_sub(e0, b0), n0)
        n = n0 if s >= 0 else (-n0[0], -n0[1])
        dist = abs(s)
        t_part = max(v_dot(v_sub(p, b0), w) for p in part_points)
        bend_end = max((b0, b1), key=lambda p: v_dot(v_sub(p, b0), w))
        edge_end = max((e0, e1), key=lambda p: v_dot(v_sub(p, b0), w))

        best = None
        for level in range(max_levels):
            t_line = t_part + offset + level * step
            a = v_add(b0, v_mul(w, t_line))
            c = v_add(a, v_mul(n, dist))
            lines = [(a, c), (bend_end, a), (edge_end, c)]
            # text centred; if the dimension is narrower than its text, also try just outside the
            # extension lines (edge side first - beyond the outer edge is usually free - then bend side)
            centers = [v_mid(a, c)]
            if dist < text_w + 2 * gap:
                centers += [v_add(c, v_mul(n, text_w / 2.0 + gap)), v_sub(a, v_mul(n, text_w / 2.0 + gap))]
            for center in centers:
                text = _oriented_box(center, n, text_w / 2.0, w, text_h / 2.0)
                hits = 0
                hits += sum(_rect_overlap(text, r, gap) for r in placed_rects)
                hits += sum(_rect_overlap(text, o) for o in obstacles)
                hits += sum(_segment_hits_rect(p, q, r) for p, q in lines for r in placed_rects)
                hits += sum(_segment_hits_rect(p, q, text) for p, q in placed_lines)
                hits += sum(_segment_hits_rect(p, q, text) for p, q in lines[1:])  # own extension lines
                hits += sum(segments_intersect(a, c, p, q) for p, q in placed_dim_lines)
                if bounds is not None and not (bounds[0] <= text[0] and text[2] <= bounds[2] and
                                               bounds[1] <= text[1] and text[3] <= bounds[3]):
                    hits += 1
                if best is None or hits < best[0]:
                    best = (hits, level, t_line, a, c, center, text, lines)
                if hits == 0:
                    break
            if best[0] == 0:
                break

        hits, level, t_line, a, c, center, text, lines = best
        placed_rects.append(text)
        placed_lines.extend(lines)
        placed_dim_lines.append((a, c))
        placements.append({"key": dim["key"], "text_point": center, "level": level, "dim_line": (a, c),
                           "text_box": text, "side": side_name(w), "collisions": hits, "sheet_distance": dist})

    needs = {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}
    for p in placements:
        r = p["text_box"]
        needs["left"] = max(needs["left"], bbox[0] - r[0])
        needs["right"] = max(needs["right"], r[2] - bbox[2])
        needs["bottom"] = max(needs["bottom"], bbox[1] - r[1])
        needs["top"] = max(needs["top"], r[3] - bbox[3])
    return placements, needs


ORDER_ALIASES = {
    "largest_outside": "largest_outside",
    "smallest_outside": "smallest_outside",
    # names used by earlier versions (same behaviour as before):
    "largest_first": "largest_outside",
    "smallest_first": "smallest_outside",
}


def order_key(order):
    """Sort key for placement order. Dimensions placed first get the levels nearest the part, so
    'largest_outside' places the smallest values first (largest end up farthest out)."""
    mode = ORDER_ALIASES.get(str(order), "largest_outside")
    if mode == "largest_outside":
        return lambda item: (item["value"], str(item["id"]))
    return lambda item: (-item["value"], str(item["id"]))
