"""Bend dimensions through SwDrawingExporter with fake SolidWorks objects (coordinate chain, tags, verify)."""
import math

import pytest

from solidworks.bend_geometry import Xform
from solidworks.sw_drawings import SwDrawingExporter
from fakesw import Embedding, FakeApp, FakeDrawing, FakeView, FlatPart, RowXform, rect

MM = 0.001


def mm(*pts):
    return [(x * MM, y * MM) for x, y in pts]


def setup(part, emb=None, verify="flag", **view_kw):
    app = FakeApp(part, emb)
    drawing = FakeDrawing(app, "/tmp/test.SLDDRW")
    view = FakeView(app.world, part, app.emb, **view_kw)
    exporter = SwDrawingExporter(app, bend_dim_verify=verify)
    exporter._current_item = {"name": "part", "thickness": part.thickness * 1000}
    return exporter, drawing, view, app


def bend_dims(app):
    return [d for d in app.world.dims if d.kind == "bend"]


def results_by_id(analysis):
    return {r["id"]: r for r in analysis["results"]}


PLATE = FlatPart(mm(*rect(100, 60)), [mm((25, 0), (25, 60))[0:2], mm((0, 45), (100, 45))[0:2]])


# ------------------------------------------------------------------ root cause
def test_bend_line_points_are_sketch_coordinates_not_model_coordinates():
    """Documents the original defect: mapping sketch points with ModelToViewTransform (as the old code did)
    puts the bend lines in the wrong place, because they are in the sketch's local frame."""
    exporter, drawing, view, app = setup(PLATE, scale=0.5)
    seg = view.bend_segments[0]
    sketch_pt = (seg._p0.X, seg._p0.Y, seg._p0.Z)
    true_model = app.emb.model(*PLATE.bends[0][0])
    assert math.dist(sketch_pt, true_model) > 0.05                          # not the same frame at all
    mv = Xform.from_array(view.ModelToViewTransform.ArrayData)
    naive = mv.apply(sketch_pt)[:2]                                          # old approach
    correct = view.sheet_of_local(*PLATE.bends[0][0])
    assert math.dist(naive, correct) > 0.01
    # the fixed chain: sketch -> model (inverse ModelToSketchTransform) -> sheet
    m2s = Xform.from_array(seg.GetSketch().ModelToSketchTransform.ArrayData)
    fixed = mv.apply(m2s.inverse().apply(sketch_pt))[:2]
    assert fixed == pytest.approx(correct, abs=1e-12)


def test_plate_values_correct_through_full_chain():
    exporter, drawing, view, app = setup(PLATE, scale=0.5)
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert analysis["error"] is None, steps
    res = results_by_id(analysis)
    assert res["#1"]["value"] == pytest.approx(25 * MM)   # x=25 -> left edge
    assert res["#2"]["value"] == pytest.approx(15 * MM)   # y=45 -> top edge (nearest)
    created = bend_dims(app)
    assert sorted(round(d.value / MM, 6) for d in created) == [15.0, 25.0]   # what "SolidWorks" measured
    assert any("transform convention verified" in s for s in steps)
    assert not any("MISMATCH" in s for s in steps)


# ------------------------------------------------------------------ translation / rotation / mirror
@pytest.mark.parametrize("sketch_angle", [0.0, 33.0, 90.0, 211.0])
@pytest.mark.parametrize("mirrored", [False, True])
@pytest.mark.parametrize("view_angle", [0.0, 90.0, 180.0, 37.0])
def test_values_independent_of_sketch_and_view_orientation(sketch_angle, mirrored, view_angle):
    emb = Embedding(sketch_angle_deg=sketch_angle, sketch_mirrored=mirrored, sketch_origin=(-0.04, 0.071))
    exporter, drawing, view, app = setup(PLATE, emb, scale=0.4, angle_deg=view_angle, position=(0.21, 0.16))
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.4, 0.0, steps)
    assert analysis["error"] is None, steps
    assert sorted(round(r["value"] / MM, 6) for r in analysis["results"]) == [15.0, 25.0]
    assert sorted(round(d.value / MM, 6) for d in bend_dims(app)) == [15.0, 25.0]
    # every text point lies outside the part outline, near it (placement uses the correct sheet frame)
    from solidworks.bend_geometry import point_in_polygon
    part_on_sheet = [view.sheet_of_local(*q) for q in PLATE.outer]
    o = view.GetOutline()
    for p in plan["placements"]:
        x, y = p["text_point"]
        assert not point_in_polygon((x, y), part_on_sheet)
        assert o[0] - 0.08 < x < o[2] + 0.08 and o[1] - 0.08 < y < o[3] + 0.08


def test_square_flat_pattern_in_rotated_view_puts_dimensions_on_the_right_sides():
    square = FlatPart(mm(*rect(80, 80)), [mm((20, 0), (20, 80)), mm((0, 30), (80, 30))])
    exporter, drawing, view, app = setup(square, scale=0.5, angle_deg=90.0)
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    res = results_by_id(analysis)
    assert res["#1"]["value"] == pytest.approx(20 * MM) and res["#2"]["value"] == pytest.approx(30 * MM)
    sides = {p["key"]: p["side"] for p in plan["placements"]}
    # rotated 90 deg: bend #1 (vertical in the part) is horizontal on the sheet -> its dimension goes right;
    # bend #2 becomes vertical on the sheet -> below
    assert sides == {"#1": "right", "#2": "bottom"}


def test_view_transform_without_position_is_corrected_from_the_outline():
    exporter, drawing, view, app = setup(PLATE, scale=0.5, includes_position=False, position=(0.3, 0.2))
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert "position corrected" in plan["transform_status"]
    o = view.GetOutline()
    for p in plan["placements"]:
        assert o[0] - 0.08 < p["text_point"][0] < o[2] + 0.08
        assert o[1] - 0.08 < p["text_point"][1] < o[3] + 0.08


def test_column_convention_solidworks_is_detected():
    """If SolidWorks stored matrices transposed, the MathUtility probe must detect it and values stay right."""
    class ColumnXform:
        def __init__(self, row):
            self.row = row

        @property
        def ArrayData(self):
            d = list(self.row.ArrayData)
            r = [d[0:3], d[3:6], d[6:9]]
            t = [[r[j][i] for j in range(3)] for i in range(3)]
            return tuple(t[0] + t[1] + t[2]) + tuple(d[9:])

    class ColumnPoint:
        def __init__(self, v):
            self.ArrayData = tuple(v)

        def MultiplyTransform(self, xf):
            return ColumnPoint(xf.row.apply(self.ArrayData))

    class ColumnMath:
        def CreatePoint(self, data):
            return ColumnPoint(getattr(data, "value", data))

    exporter, drawing, view, app = setup(PLATE, scale=0.5, angle_deg=30)
    app.GetMathUtility = lambda: ColumnMath()
    original = type(view).ModelToViewTransform.fget
    type(view).ModelToViewTransform = property(lambda self: ColumnXform(original(self)))
    try:
        for sketch in {id(s._sketch): s._sketch for s in view.bend_segments}.values():  # each sketch once
            sketch.ModelToSketchTransform = ColumnXform(sketch.ModelToSketchTransform)
        steps = []
        analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
        assert exporter._convention == "column" and exporter._convention_source == "IMathPoint.MultiplyTransform"
        assert sorted(round(r["value"] / MM, 6) for r in analysis["results"]) == [15.0, 25.0]
        assert not any("MISMATCH" in s for s in steps)
    finally:
        type(view).ModelToViewTransform = property(original)


# ------------------------------------------------------------------ outer edge selection
def test_internal_cutout_is_excluded():
    part = FlatPart(mm(*rect(100, 60)), [mm((30, 0), (30, 60))], holes=[mm(*rect(10, 20, 35, 20))])
    exporter, drawing, view, app = setup(part, scale=0.5)
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert analysis["results"][0]["value"] == pytest.approx(30 * MM)       # not 5 mm to the slot
    assert any("outer loop" in s for s in steps)                              # API outer loop was used
    assert [round(d.value / MM, 6) for d in bend_dims(app)] == [30.0]


def test_outer_loop_from_api_preferred_over_geometry():
    part = FlatPart(mm(*rect(100, 60)), [mm((30, 0), (30, 60))], holes=[mm(*rect(10, 20, 35, 20))])
    exporter, drawing, view, app = setup(part, scale=0.5)
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert analysis["outer_info"]["method"].startswith("face outer loop")
    assert len(analysis["outer"]) == 4


def test_rounded_corners_and_arcs_do_not_break_the_boundary():
    # outer polygon with two "arc" segments (shown as chords); bend measured to a straight edge
    outer = mm((0, 0), (90, 0), (100, 10), (100, 60), (0, 60))
    part = FlatPart(outer, [mm((70, 0), (70, 60))], arc_segments={1})
    exporter, drawing, view, app = setup(part, scale=0.5)
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, [])
    assert analysis["results"][0]["value"] == pytest.approx(30 * MM)


# ------------------------------------------------------------------ identity, tags, isolation
def test_two_bends_with_identical_values_both_dimensioned_with_their_own_tags():
    part = FlatPart(mm(*rect(100, 60)), [mm((20, 0), (20, 60)), mm((80, 0), (80, 60))], tags=["B", "A"])
    exporter, drawing, view, app = setup(part, scale=0.5)
    view.ShowSheetMetalBendNotes = True
    view.InsertBendTable(False, 0, 0, 1, "A", "")
    exporter._bend_tags = ["A", "B"]
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert [round(d.value / MM, 6) for d in bend_dims(app)] == [20.0, 20.0]
    assert plan["tags"] == {"#1": "B", "#2": "A"}                    # SolidWorks put B on bend 1, A on bend 2
    # each created dimension references its own bend line
    segs = {id(d.entities[0]) for d in bend_dims(app)}
    assert segs == {id(s) for s in view.bend_segments}
    assert any("[tag B] 20.00 mm" in s for s in steps) and any("[tag A] 20.00 mm" in s for s in steps)


def test_unreadable_bend_line_is_isolated_and_reported():
    part = FlatPart(mm(*rect(100, 60)), [mm((20, 0), (20, 60)), mm((60, 0), (60, 60)), mm((0, 45), (100, 45))])
    exporter, drawing, view, app = setup(part, scale=0.5, unreadable_bends={1})
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert analysis["error"] is None
    assert sorted(round(d.value / MM, 6) for d in bend_dims(app)) == [15.0, 20.0]
    assert any("bend #2" in s and "NOT dimensioned" in s for s in steps)


def test_no_reference_edge_is_reported_not_dimensioned():
    part = FlatPart(mm((0, 0), (100, 0), (80, 60), (20, 60)), [mm((50, 0), (50, 60)), mm((10, 20), (90, 20))])
    exporter, drawing, view, app = setup(part, scale=0.5)
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert [round(d.value / MM, 6) for d in bend_dims(app)] == [20.0]
    assert any("bend #1" in s and "no_reference_edge" in s for s in steps)


def test_frame_mismatch_aborts_instead_of_guessing():
    exporter, drawing, view, app = setup(PLATE, scale=0.5)
    # a sketch transform that is not the sheet plane (rotated 90 deg about the sheet x axis)
    e = app.emb
    wrong = RowXform((e.ax, e.az, tuple(-x for x in e.ay)), (0.0, 0.0, 0.0))
    for seg in view.bend_segments:
        seg._sketch.ModelToSketchTransform = wrong
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert analysis["error"] and "do not match" in analysis["error"]
    assert plan is None and bend_dims(app) == []


@pytest.mark.parametrize("mode,expect_kept", [("flag", 2), ("delete", 1)])
def test_dimension_read_back_mismatch(mode, expect_kept, monkeypatch):
    exporter, drawing, view, app = setup(PLATE, verify=mode, scale=0.5)
    real = drawing.AddDimension2
    calls = {"n": 0}

    def wrong_first(x, y, z):
        dim = real(x, y, z)
        calls["n"] += 1
        if calls["n"] == 1:
            dim.value += 0.001  # SolidWorks measured something else (e.g. the wrong entity)
        return dim
    monkeypatch.setattr(drawing, "AddDimension2", wrong_first)
    deleted = []
    monkeypatch.setattr(exporter, "_delete_annotation", lambda d, disp: deleted.append(disp) or True)
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, steps)
    assert sum("MISMATCH" in s for s in steps) == 2  # the per-bend line + the summary
    created = [r for r in analysis["results"] if r.get("created")]
    assert len(created) == expect_kept
    assert len(deleted) == (1 if mode == "delete" else 0)


# ------------------------------------------------------------------ collisions in a real chain
def test_dense_bends_placed_without_text_overlap():
    xs = [8, 11, 14, 17, 20, 70, 74]
    part = FlatPart(mm(*rect(100, 60)), [mm((x, 0), (x, 60)) for x in xs] + [mm((0, 50), (100, 50))])
    exporter, drawing, view, app = setup(part, scale=1.0, angle_deg=0.0)
    view.ShowSheetMetalBendNotes = True
    view.InsertBendTable(False, 0, 0, 1, "A", "")
    exporter._bend_tags = list(part.tags)
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 1.0, 0.0, steps)
    boxes = [p["text_box"] for p in plan["placements"]]
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            assert not (a[0] < b[2] and a[2] > b[0] and a[1] < b[3] and a[3] > b[1])
    assert all(p["collisions"] == 0 for p in plan["placements"])
    assert len(bend_dims(app)) == len(xs) + 1
    # bends 3 mm apart with tags 3 mm beside their lines are genuinely ambiguous: every tag is either
    # paired correctly or reported as ambiguous - never silently given to the wrong bend
    expected = {f"#{i + 1}": t for i, t in enumerate(part.tags)}
    assert all(expected[b] == t for b, t in plan["tags"].items())
    assert set(plan["tags"]) | set(plan["ambiguous_tags"]) | set() <= set(expected)
    assert plan["ambiguous_tags"] and any("AMBIGUOUS" in s for s in steps)


def test_well_spaced_tags_all_paired():
    xs = [10, 30, 50, 70]
    part = FlatPart(mm(*rect(100, 60)), [mm((x, 0), (x, 60)) for x in xs], tags=["C", "A", "D", "B"])
    exporter, drawing, view, app = setup(part, scale=0.5)
    view.ShowSheetMetalBendNotes = True
    view.InsertBendTable(False, 0, 0, 1, "A", "")
    exporter._bend_tags = ["A", "B", "C", "D"]
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 0.5, 0.0, [])
    assert plan["tags"] == {"#1": "C", "#2": "A", "#3": "D", "#4": "B"} and not plan["ambiguous_tags"]


# ------------------------------------------------------------------ randomized property check
@pytest.mark.parametrize("seed", range(40))
def test_randomized_plates_match_ground_truth(seed):
    import random
    rnd = random.Random(seed)
    W, H = rnd.uniform(40, 400), rnd.uniform(30, 300)
    xs = sorted(rnd.sample(range(5, int(W) - 5), rnd.randint(1, 3)))
    ys = sorted(rnd.sample(range(5, int(H) - 5), rnd.randint(0, 2)))
    bends = [mm((x, 0), (x, H)) for x in xs] + [mm((0, y), (W, y)) for y in ys]
    expected = sorted(round(min(x, W - x), 6) for x in xs) + sorted(round(min(y, H - y), 6) for y in ys)
    emb = Embedding(origin=(rnd.uniform(-1, 1), rnd.uniform(-1, 1), rnd.uniform(-1, 1)),
                    ax=(rnd.uniform(-1, 1), rnd.uniform(-1, 1), rnd.uniform(-1, 1)),
                    ay=(rnd.uniform(-1, 1), rnd.uniform(-1, 1), rnd.uniform(-1, 1)),
                    sketch_origin=(rnd.uniform(-0.2, 0.2), rnd.uniform(-0.2, 0.2)),
                    sketch_angle_deg=rnd.uniform(0, 360), sketch_mirrored=rnd.random() < 0.5)
    part = FlatPart(mm(*rect(W, H)), bends)
    exporter, drawing, view, app = setup(part, emb, scale=rnd.choice([1, 0.5, 0.2, 0.1]),
                                         angle_deg=rnd.choice([0, 90, 180, 270, rnd.uniform(0, 360)]))
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, view.ScaleDecimal, -1.0, steps)
    assert analysis["error"] is None, steps
    nv = len(xs)
    got = [r["value"] / MM for r in analysis["results"]]
    assert sorted(round(v, 6) for v in got[:nv]) == pytest.approx(expected[:nv], abs=1e-6)
    assert sorted(round(v, 6) for v in got[nv:]) == pytest.approx(expected[nv:], abs=1e-6)
    measured = sorted(d.value / MM for d in bend_dims(app))
    assert measured == pytest.approx(sorted(got), abs=1e-6)       # the created dimensions read back the same
