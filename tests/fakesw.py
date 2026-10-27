"""Fake SolidWorks objects for offline tests.

The fakes model the coordinate chain that matters for bend dimensions, with their OWN math
(nothing is imported from the code under test):

  local (u, v, w)  flat sheet coordinates, metres; w = 0 top face, w = -t bottom face
  model            = origin + u*ax + v*ay + w*az        (part embedded at an arbitrary orientation)
  sketch           bend-line sketch: own origin, own rotation, optionally mirrored (normal = -az)
  sheet            = position + scale * Rot(view angle) * (u, v)

MathTransforms are exposed as ArrayData in the documented SolidWorks row-vector form
p' = s * (p . R) + T, and IMathPoint.MultiplyTransform applies exactly that.
"""
import math


def _val(v):
    return getattr(v, "value", v)


def dot3(a, b):
    return sum(x * y for x, y in zip(a, b))


def cross3(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])


def unit3(a):
    n = math.sqrt(dot3(a, a))
    return tuple(x / n for x in a)


class RowXform:
    """p' = s * (p . R) + T  with R[i][j] = (image of axis i)[j]."""

    def __init__(self, columns, T, s=1.0):
        # columns: three 3-vectors; output_j = s * dot(p, columns[j]) + T[j]
        self.R = [[columns[j][i] for j in range(3)] for i in range(3)]
        self.T = tuple(T)
        self.s = s

    def apply(self, p):
        p = tuple(p) + (0.0,) * (3 - len(p))
        return tuple(self.s * sum(p[i] * self.R[i][j] for i in range(3)) + self.T[j] for j in range(3))

    @property
    def ArrayData(self):
        r = self.R
        return tuple(r[0] + r[1] + r[2]) + self.T + (self.s, 0.0, 0.0, 0.0)


class FakeMathPoint:
    def __init__(self, values):
        self._v = tuple(float(x) for x in values)

    @property
    def ArrayData(self):
        return self._v

    def MultiplyTransform(self, xform):
        d = xform.ArrayData
        R = [d[0:3], d[3:6], d[6:9]]
        out = tuple(d[12] * sum(self._v[i] * R[i][j] for i in range(3)) + d[9 + j] for j in range(3))
        return FakeMathPoint(out)


class FakeMathUtility:
    def CreatePoint(self, data):
        return FakeMathPoint(_val(data))


# ---------------------------------------------------------------- part geometry
class FlatPart:
    """Flat sheet in local (u, v): outer polygon, holes, bend lines (in metres)."""

    def __init__(self, outer, bends, holes=(), thickness=0.002, arc_segments=(), tags=None):
        self.outer = [tuple(p) for p in outer]
        self.holes = [[tuple(p) for p in h] for h in holes]
        self.bends = [(tuple(a), tuple(b)) for a, b in bends]
        self.thickness = thickness
        self.arc_segments = set(arc_segments)     # indexes of outer segments that are arcs (shown as chords)
        self.tags = list(tags) if tags else [chr(65 + i) for i in range(len(bends))]


class Embedding:
    """Where the flat sheet sits in model space and how its bend-line sketch is oriented."""

    def __init__(self, origin=(0.31, -0.12, 0.45), ax=None, ay=None,
                 sketch_origin=(0.013, -0.027), sketch_angle_deg=33.0, sketch_mirrored=False):
        c, s = math.cos(math.radians(25)), math.sin(math.radians(25))
        self.origin = origin
        self.ax = unit3(ax or (c, 0.0, -s))          # sheet lies in a tilted plane, not XY
        ay = ay or (0.2, 0.9, 0.3)
        ay = tuple(a - dot3(ay, self.ax) * b for a, b in zip(ay, self.ax))
        self.ay = unit3(ay)
        self.az = cross3(self.ax, self.ay)
        self.sketch_origin = sketch_origin
        self.sketch_angle = math.radians(sketch_angle_deg)
        self.sketch_mirrored = sketch_mirrored

    def model(self, u, v, w=0.0):
        return tuple(o + u * a + v * b + w * c for o, a, b, c in zip(self.origin, self.ax, self.ay, self.az))

    def _sketch_axes(self):
        ca, sa = math.cos(self.sketch_angle), math.sin(self.sketch_angle)
        sign = -1.0 if self.sketch_mirrored else 1.0
        x_local, y_local = (ca, sa), (-sa * sign, ca * sign)
        X = tuple(x_local[0] * a + x_local[1] * b for a, b in zip(self.ax, self.ay))
        Y = tuple(y_local[0] * a + y_local[1] * b for a, b in zip(self.ax, self.ay))
        return X, Y, cross3(X, Y)

    def model_to_sketch(self):
        X, Y, Z = self._sketch_axes()
        O = self.model(*self.sketch_origin)
        return RowXform((X, Y, Z), tuple(-dot3(O, axis) for axis in (X, Y, Z)))

    def sketch_point(self, u, v):
        return self.model_to_sketch().apply(self.model(u, v))


# ---------------------------------------------------------------- topology objects
class World:
    def __init__(self):
        self.selection = []
        self.dims = []


class FakeVertex:
    def __init__(self, world, view, local, model):
        self.world, self.view, self.local, self._model = world, view, local, model

    def GetPoint(self):
        return self._model

    def Select4(self, append, sel):
        if not append:
            self.world.selection.clear()
        self.world.selection.append(self)
        return True


class FakeCurve:
    def __init__(self, is_line):
        self._line = is_line

    def IsLine(self):
        return self._line


class FakeLoop:
    def __init__(self, edges, outer):
        self._edges, self._outer = edges, outer

    def IsOuter(self):
        return self._outer

    def GetEdges(self):
        return tuple(self._edges)


class FakeFace:
    def __init__(self, area, normal, loops=()):
        self._area, self.Normal, self._loops = area, normal, list(loops)

    def GetArea(self):
        return self._area

    def GetLoops(self):
        return tuple(self._loops)


class FakeEdge:
    def __init__(self, world, v0, v1, is_line=True, faces=()):
        self.world, self.v0, self.v1 = world, v0, v1
        self._curve = FakeCurve(is_line)
        self.faces = list(faces)

    def GetStartVertex(self):
        return self.v0

    def GetEndVertex(self):
        return self.v1

    def GetCurve(self):
        return self._curve

    def GetTwoAdjacentFaces2(self):
        return tuple(self.faces) + (None,) * (2 - len(self.faces))

    def Select4(self, append, sel):
        if not append:
            self.world.selection.clear()
        self.world.selection.append(self)
        return True


class FakeSketchPoint:
    def __init__(self, xyz):
        self.X, self.Y, self.Z = xyz


class FakeSketch:
    def __init__(self, xform):
        self.ModelToSketchTransform = xform


class FakeSketchSegment:
    def __init__(self, world, sketch, p0, p1, local, unreadable=False):
        self.world, self._sketch, self.local, self.unreadable = world, sketch, local, unreadable
        self._p0, self._p1 = FakeSketchPoint(p0), FakeSketchPoint(p1)

    def GetSketch(self):
        return self._sketch

    def GetStartPoint2(self):
        return None if self.unreadable else self._p0

    def GetEndPoint2(self):
        return self._p1

    def Select4(self, append, sel):
        if not append:
            self.world.selection.clear()
        self.world.selection.append(self)
        return True


class FakeAnnotation:
    def __init__(self, pos=(0.0, 0.0, 0.0)):
        self._pos = tuple(pos)

    def GetPosition(self):
        return self._pos

    def SetPosition2(self, x, y, z):
        self._pos = (x, y, z)
        return True


class FakeNote:
    def __init__(self, text, annotation):
        self._text, self._ann = text, annotation

    def GetText(self):
        return self._text

    def GetAnnotation(self):
        return self._ann


class FakeTable:
    def __init__(self, tags):
        self.header = ["Tag", "Direction", "Angle", "Inner Radius", "K-Factor"]
        self.rows = [[t, "UP", "90°", "1"] + ["0.5"] for t in tags]
        self._ann = FakeAnnotation()

    @property
    def ColumnCount(self):
        return len(self.header)

    @property
    def RowCount(self):
        return 1 + len(self.rows)

    def Text(self, r, c):
        return self.header[c] if r == 0 else self.rows[r - 1][c]

    def DeleteColumn(self, c):
        del self.header[c]
        for row in self.rows:
            del row[c]
        return True

    def GetColumnWidth(self, c):
        return 0.018

    def GetRowHeight(self, r):
        return 0.006

    def GetAnnotation(self):
        return self._ann


# ---------------------------------------------------------------- drawing view
class FakeView:
    """A drawing view of a FlatPart (flat pattern) or of a plain rectangle (folded model view)."""

    def __init__(self, world, part, emb=None, scale=1.0, angle_deg=0.0, position=(0.2, 0.15),
                 includes_position=True, name="Flat", flat=True, unreadable_bends=()):
        self.world, self.part, self.emb = world, part, emb or Embedding()
        self.ScaleDecimal, self.Angle = scale, math.radians(angle_deg)
        self._pos = tuple(position)
        self.includes_position = includes_position
        self.name, self.flat = name, flat
        self.ShowSheetMetalBendNotes = False
        self.table = None
        self._build(unreadable_bends)

    # geometry -> topology
    def _build(self, unreadable_bends):
        p, e, w = self.part, self.emb, self.world
        t = p.thickness
        self.top_face = FakeFace(abs(_area(p.outer)) - sum(abs(_area(h)) for h in p.holes), e.az)
        self.bottom_face = FakeFace(self.top_face.GetArea(), tuple(-x for x in e.az))
        self.edges, self.vertices, self.hidden_edges = [], [], []
        loops_top, loops_bottom = [], []
        for loop_idx, poly in enumerate([p.outer] + p.holes):
            top_v = [FakeVertex(w, self, q, e.model(q[0], q[1], 0.0)) for q in poly]
            bot_v = [FakeVertex(w, self, q, e.model(q[0], q[1], -t)) for q in poly]
            self.vertices.extend(top_v)
            top_edges, bot_edges = [], []
            for i in range(len(poly)):
                j = (i + 1) % len(poly)
                is_line = not (loop_idx == 0 and i in p.arc_segments)
                side = FakeFace(0.0, (0.0, 0.0, 0.0))  # side faces: not planar-parallel to the sheet
                te = FakeEdge(w, top_v[i], top_v[j], is_line, [self.top_face, side])
                be = FakeEdge(w, bot_v[i], bot_v[j], is_line, [self.bottom_face, side])
                top_edges.append(te)
                bot_edges.append(be)
                self.edges.append(te)
                self.hidden_edges.append(be)
                self.edges.append(FakeEdge(w, top_v[i], bot_v[i], True, [side]))  # thickness edge
            loops_top.append(FakeLoop(top_edges, loop_idx == 0))
            loops_bottom.append(FakeLoop(bot_edges, loop_idx == 0))
        self.top_face._loops, self.bottom_face._loops = loops_top, loops_bottom
        sketch = FakeSketch(e.model_to_sketch())
        self.bend_segments = [
            FakeSketchSegment(w, sketch, e.sketch_point(*a), e.sketch_point(*b), (a, b), i in unreadable_bends)
            for i, (a, b) in enumerate(p.bends)] if self.flat else []

    # local <-> sheet
    def _local_of_model(self, m):
        d = tuple(a - b for a, b in zip(m, self.emb.origin))
        return dot3(d, self.emb.ax), dot3(d, self.emb.ay)

    def sheet_of_local(self, u, v):
        c, s = math.cos(self.Angle), math.sin(self.Angle)
        k = self.ScaleDecimal
        return (self._pos[0] + k * (c * u - s * v), self._pos[1] + k * (s * u + c * v))

    def sheet_of_model(self, m):
        return self.sheet_of_local(*self._local_of_model(m))

    # SolidWorks IView API
    @property
    def ModelToViewTransform(self):
        c, s = math.cos(self.Angle), math.sin(self.Angle)
        e = self.emb
        colx = tuple(c * a - s * b for a, b in zip(e.ax, e.ay))
        coly = tuple(s * a + c * b for a, b in zip(e.ax, e.ay))
        cols = (colx, coly, e.az)
        k = self.ScaleDecimal
        pos = self._pos if self.includes_position else (0.0, 0.0)
        T = (pos[0] - k * dot3(e.origin, colx), pos[1] - k * dot3(e.origin, coly), -k * dot3(e.origin, e.az))
        return RowXform(cols, T, k)

    def GetOutline(self):
        pts = [self.sheet_of_local(*q) for q in self.part.outer]
        return (min(x for x, _ in pts), min(y for _, y in pts), max(x for x, _ in pts), max(y for _, y in pts))

    @property
    def Position(self):
        return self._pos

    @Position.setter
    def Position(self, value):
        value = _val(value)
        self._pos = (float(value[0]), float(value[1]))

    def GetVisibleEntities2(self, component, entity_type):
        if entity_type == 1:
            return tuple(self.edges)
        if entity_type == 3:
            return tuple(self.vertices)
        return None

    def GetBendLines(self):
        return tuple(self.bend_segments)

    def GetNotes(self):
        """Bend tags as SolidWorks shows them: next to the middle of their bend line (3 mm on paper)."""
        if not (self.flat and self.ShowSheetMetalBendNotes and self.table):
            return ()
        notes = []
        for tag, seg in zip(self.part.tags, self.bend_segments):
            (a, b) = seg.local
            pa, pb = self.sheet_of_local(*a), self.sheet_of_local(*b)
            mx, my = (pa[0] + pb[0]) / 2, (pa[1] + pb[1]) / 2
            dx, dy = pb[0] - pa[0], pb[1] - pa[1]
            n = math.hypot(dx, dy)
            notes.append(FakeNote(tag, FakeAnnotation((mx - dy / n * 0.003, my + dx / n * 0.003, 0.0))))
        return tuple(notes)

    def InsertBendTable(self, use_anchor, x, y, anchor, start, template):
        self.table = FakeTable(self.part.tags)
        self.table._ann.SetPosition2(x, y, 0.0)
        return self.table

    def SelectEntity(self, entity, append):
        return entity.Select4(append, None)

    def GetTableAnnotations(self):
        return (self.table,) if self.table else ()

    def GetName2(self):
        return self.name


def _area(poly):
    return sum(poly[i][0] * poly[(i + 1) % len(poly)][1] - poly[(i + 1) % len(poly)][0] * poly[i][1]
               for i in range(len(poly))) / 2.0


def rect(w, h, x0=0.0, y0=0.0):
    return [(x0, y0), (x0 + w, y0), (x0 + w, y0 + h), (x0, y0 + h)]


# ---------------------------------------------------------------- drawing + application
class FakeDisplayDimension:
    def __init__(self, value, kind, pos, entities):
        self.value, self.kind, self.pos, self.entities = value, kind, pos, entities
        self._ann = FakeAnnotation(pos)

    def GetDimension2(self, index):
        class _D:
            pass
        d = _D()
        d.SystemValue = self.value
        return d

    def GetAnnotation(self):
        return self._ann


class _SelMgr:
    def CreateSelectData(self):
        class _S:
            View = None
        return _S()


class _PropMgr:
    def __init__(self, store):
        self.store = store

    def Add3(self, key, kind, value, option):
        self.store[key] = value
        return 0


class _Extension:
    def __init__(self, drawing):
        self.drawing = drawing

    def CustomPropertyManager(self, config):
        return _PropMgr(self.drawing.properties)

    def DeleteSelection2(self, options):
        return True

    def SelectByID2(self, *args):
        return False


class _Sheet:
    def __init__(self, size):
        self.size, self.scale = size, None

    def GetSize(self, w, h):
        w.value, h.value = self.size
        return True

    def SetScale(self, num, den, a, b):
        self.scale = (num, den)
        return True


class _SheetView:
    def GetFirstNote(self):
        return None

    def GetNextView(self):
        return None


class FakeDrawing:
    def __init__(self, app, path, sheet=(0.420, 0.297)):
        self.app, self.path = app, path
        self.world = app.world
        self.sheet = _Sheet(sheet)
        self.properties = {}
        self.views = []
        self.notes = []
        self.SelectionManager = _SelMgr()
        self.Extension = _Extension(self)

    def GetTitle(self):
        return self.path.split("/")[-1]

    def GetCurrentSheet(self):
        return self.sheet

    def GetFirstView(self):
        return _SheetView()

    def EditTemplate(self):
        return None

    def EditSheet(self):
        return None

    def ClearSelection2(self, all_):
        self.world.selection.clear()
        return True

    def CreateFlatPatternViewFromModelView3(self, path, cfg, x, y, z, hide, flip):
        view = FakeView(self.world, self.app.part, self.app.emb, scale=1.0, position=(x, y), name="Flat")
        self.views.append(("flat", view))
        return view

    def CreateDrawViewFromModelView3(self, path, name, x, y, z):
        L, D, H = self.app.box
        sizes = {"*Top": (L, D), "*Front": (L, H), "*Right": (D, H)}
        if name not in sizes:
            return None
        part = FlatPart(rect(*sizes[name]), [])
        emb = Embedding(origin=(0, 0, 0), ax=(1, 0, 0), ay=(0, 1, 0))
        view = FakeView(self.world, part, emb, position=(x, y), name=name, flat=False)
        self.views.append((name, view))
        return view

    def _record(self, kind, value, x, y):
        dim = FakeDisplayDimension(value, kind, (x, y), list(self.world.selection))
        self.world.dims.append(dim)
        return dim

    def AddDimension2(self, x, y, z):
        sel = self.world.selection
        if len(sel) != 2 or not isinstance(sel[0], FakeSketchSegment) or not isinstance(sel[1], FakeEdge):
            return None
        (a, b), edge = sel[0].local, sel[1]
        e0, e1 = edge.v0.local, edge.v1.local
        dx, dy = b[0] - a[0], b[1] - a[1]
        n = math.hypot(dx, dy)
        mid = ((e0[0] + e1[0]) / 2, (e0[1] + e1[1]) / 2)
        value = abs((mid[0] - a[0]) * dy - (mid[1] - a[1]) * dx) / n   # true-size distance in the sheet plane
        return self._record("bend", value, x, y)

    def _overall(self, kind, axis, x, y):
        sel = self.world.selection
        if len(sel) != 2:
            return None
        pa, pb = (s.view.sheet_of_model(s.GetPoint()) for s in sel)
        return self._record(kind, abs(pa[axis] - pb[axis]) / sel[0].view.ScaleDecimal, x, y)

    def AddHorizontalDimension2(self, x, y, z):
        return self._overall("horizontal", 0, x, y)

    def AddVerticalDimension2(self, x, y, z):
        return self._overall("vertical", 1, x, y)

    def InsertNote(self, text):
        note = FakeNote(text, FakeAnnotation())
        self.notes.append(note)
        return note

    def ForceRebuild3(self, top_only):
        return True

    def SaveAs3(self, path, version, options):
        with open(path, "wb") as f:
            f.write(b"fake slddrw")
        return 0


class _Feature:
    Name = "Flat-Pattern1"

    def GetTypeName2(self):
        return "FlatPattern"

    def IsSuppressed(self):
        return False


class _Config:
    Name = "Default"


class _ConfigMgr:
    ActiveConfiguration = _Config()


class FakeModel:
    ConfigurationManager = _ConfigMgr()

    def __init__(self, path):
        self.path = path

    def GetTitle(self):
        return self.path.split("/")[-1]

    def GetPathName(self):
        return self.path

    def ShowConfiguration2(self, name):
        return True

    def ForceRebuild3(self, top_only):
        return True

    def FeatureByName(self, name):
        return _Feature() if name == "Flat-Pattern1" else None


class FakeApp:
    def __init__(self, part, emb=None, box=(0.12, 0.05, 0.04)):
        self.world = World()
        self.part, self.emb, self.box = part, emb or Embedding(), box
        self.drawings = []
        self.closed = []

    def GetOpenDocumentByName(self, path):
        return None

    def OpenDoc6(self, path, doc_type, options, config, errors, warnings):
        if doc_type == 1:
            return FakeModel(path)
        drawing = FakeDrawing(self, path)
        self.drawings.append(drawing)
        return drawing

    def CloseDoc(self, title):
        self.closed.append(title)

    def ActivateDoc3(self, *args):
        return None

    def GetMathUtility(self):
        return FakeMathUtility()

    def GetExecutablePath(self):
        return "/nonexistent/sldworks.exe"
