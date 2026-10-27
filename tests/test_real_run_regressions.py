"""Regressions from the first real SolidWorks 2025 run (part AHU22000-D03-01).

Observed there:
  * IView.GetOutline is the geometry plus a constant 4.2 mm border per side, on every view and scale;
    the phase-2 check rejected that and bend dimensions were not created.
  * IMathPoint.MultiplyTransform failed with "Member not found" -> convention could not be probed.
  * overall dimensions then fell back to model coordinates (wrong for a rotated view; flat width failed).
  * visible entity types: 1 = edges (incl. closed hole edges without vertices), 2 = vertices, 3 = face.
"""
from pathlib import Path

import pytest

from solidworks.sw_drawings import SwDrawingExporter
from fakesw import Embedding, FakeApp, FakeDrawing, FakeView, FlatPart, rect

MM = 0.001
REAL_LIKE = FlatPart(
    [(x * MM, y * MM) for x, y in rect(2149, 1159)],
    [((300 * MM, 0), (300 * MM, 1159 * MM)), ((1800 * MM, 0), (1800 * MM, 1159 * MM)),
     ((0, 900 * MM), (2149 * MM, 900 * MM)), ((0, 200 * MM), (2149 * MM, 200 * MM))],
    thickness=0.001, round_holes=28)


def make(part=REAL_LIKE, emb=None, math_broken=True, **view_kw):
    app = FakeApp(part, emb, math_broken=math_broken)
    drawing = FakeDrawing(app, "/tmp/t.SLDDRW", sheet=(0.210, 0.297))
    view = FakeView(app.world, part, app.emb, **view_kw)
    exporter = SwDrawingExporter(app)
    exporter._current_item = {"name": "AHU", "thickness": part.thickness * 1000}
    return exporter, drawing, view, app


def test_outline_border_is_accepted_not_treated_as_a_mismatch():
    exporter, drawing, view, app = make(scale=1 / 22, angle_deg=90)
    steps = []
    mapper = exporter._sheet_mapper(view, steps)
    assert "outline border 4.2 mm per side" in mapper.status and "corrected" not in mapper.status
    # mapped geometry is exactly the outline minus the border
    corners = [mapper(app.emb.model(*q)) for q in REAL_LIKE.outer]
    g = view.geometry_box()
    assert min(x for x, _ in corners) == pytest.approx(g[0], abs=1e-9)
    assert max(y for _, y in corners) == pytest.approx(g[3], abs=1e-9)


def test_a_real_mismatch_is_still_rejected():
    exporter, drawing, view, app = make(scale=1 / 22)
    original = type(view).ModelToViewTransform.fget
    try:  # stretch x only -> not a border, not a uniform scale -> must abort
        def stretched(self):
            xf = original(self)
            for row in xf.R:
                row[0] *= 1.3
            return xf
        type(view).ModelToViewTransform = property(stretched)
        with pytest.raises(RuntimeError, match="does not match the view outline"):
            exporter._sheet_mapper(view, [])
    finally:
        type(view).ModelToViewTransform = property(original)


def test_broken_math_utility_geometry_decides_row_and_logs_once():
    exporter, drawing, view, app = make(scale=1 / 22)
    steps = []
    analysis, plan = exporter._add_bend_dimensions(drawing, view, 1 / 22, 0.0, steps)
    assert analysis["error"] is None, steps
    assert exporter._convention == "row" and exporter._convention_source == "geometry (bend frame)"
    assert sum("convention check not possible" in s for s in steps) == 1      # not once per bend
    assert any("Member not found" in s for s in steps)                         # failing call is named
    assert sorted(round(r["value"] / MM, 3) for r in analysis["results"]) == [200.0, 259.0, 300.0, 349.0]
    assert plan is not None and len([d for d in app.world.dims if d.kind == "bend"]) == 4


def test_broken_math_utility_with_transposed_storage_geometry_decides_column():
    class ColumnXform:  # SolidWorks storing matrices the other way round
        def __init__(self, row):
            self.row = row

        @property
        def ArrayData(self):
            d = list(self.row.ArrayData)
            r = [d[0:3], d[3:6], d[6:9]]
            t = [[r[j][i] for j in range(3)] for i in range(3)]
            return tuple(t[0] + t[1] + t[2]) + tuple(d[9:])

    exporter, drawing, view, app = make(scale=1 / 22)
    original = type(view).ModelToViewTransform.fget
    type(view).ModelToViewTransform = property(lambda self: ColumnXform(original(self)))
    try:
        for sketch in {id(s._sketch): s._sketch for s in view.bend_segments}.values():
            sketch.ModelToSketchTransform = ColumnXform(sketch.ModelToSketchTransform)
        steps = []
        analysis, plan = exporter._add_bend_dimensions(drawing, view, 1 / 22, 0.0, steps)
        assert exporter._convention == "column"
        assert sorted(round(r["value"] / MM, 3) for r in analysis["results"]) == [200.0, 259.0, 300.0, 349.0]
        assert "outline border" in plan["transform_status"]
    finally:
        type(view).ModelToViewTransform = property(original)


def test_symmetric_frame_reports_identical_results_without_guessing():
    emb = Embedding(origin=(0, 0, 0), ax=(1, 0, 0), ay=(0, 1, 0), sketch_origin=(0, 0), sketch_angle_deg=0)
    exporter, drawing, view, app = make(emb=emb, scale=1 / 22)
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 1 / 22, 0.0, steps)
    assert analysis["error"] is None
    assert any("both forms give identical bend results" in s for s in steps)
    assert exporter._convention is None       # nothing learnt -> nothing claimed


def test_overall_dimensions_never_fall_back_to_model_coordinates():
    exporter, drawing, view, app = make(scale=1 / 22, angle_deg=90)
    original = type(view).ModelToViewTransform.fget
    try:
        def stretched(self):
            xf = original(self)
            for row in xf.R:
                row[0] *= 1.3
            return xf
        type(view).ModelToViewTransform = property(stretched)
        steps = []
        exporter._add_overall_dimensions(drawing, view, steps, "flat", True, True, "left")
        assert any("dimensions flat: NOT created - view transform unusable" in s for s in steps)
        assert [d for d in app.world.dims if d.kind in ("horizontal", "vertical")] == []
    finally:
        type(view).ModelToViewTransform = property(original)


def test_overall_dimensions_verified_on_a_rotated_view():
    exporter, drawing, view, app = make(scale=1 / 22, angle_deg=90)
    steps = []
    exporter._add_overall_dimensions(drawing, view, steps, "flat", True, True, "left")
    line = next(s for s in steps if s.startswith("dimensions flat ->"))
    # rotated 90 deg on the sheet: horizontal extent is the 1159 side, vertical the 2149 side
    assert "length: ok 1159.0 mm verified" in line and "width: ok 2149.0 mm verified" in line


def test_unit_size_separates_the_border():
    exporter, drawing, view, app = make(scale=1 / 22)
    w, h, bw, bh = exporter._unit_size_and_border(view)
    assert w == pytest.approx(2.149, abs=1e-9) and h == pytest.approx(1.159, abs=1e-9)
    assert bw == pytest.approx(0.0084, abs=1e-9) and bh == pytest.approx(0.0084, abs=1e-9)


def test_round_hole_edges_without_vertices_are_ignored():
    exporter, drawing, view, app = make(scale=1 / 22)
    steps = []
    analysis, _ = exporter._add_bend_dimensions(drawing, view, 1 / 22, 0.0, steps)
    assert len(analysis["outer"]) == 4 and analysis["outer_info"]["method"].startswith("face outer loop")


def test_real_like_export_end_to_end(tmp_path):
    master = tmp_path / "master.slddrw"
    master.write_bytes(b"m")
    app = FakeApp(REAL_LIKE, box=(2.149, 0.5, 0.04), math_broken=True)
    exporter = SwDrawingExporter(app, master_drawing=master, folded_view_names=("*Top", "*Front", "*Right"),
                                 bend_debug_json=True)
    ok, message, steps = exporter.export_part(
        {"name": "AHU22000-D03-01", "path": str(tmp_path / "p.SLDPRT"), "quantity": 4, "thickness": 1.0},
        tmp_path / "AHU22000-D03-01.SLDDRW")
    assert ok, steps
    assert sum("verified" in s and s.startswith("dimensions ") for s in steps) == 4   # all four views
    assert not any("MISMATCH" in s or "NOT created" in s or "model coordinates" in s for s in steps)
    assert any("bend dimensions: 4/4 created and verified" in s for s in steps)
    assert any("bend tags paired: 4 of 4" in s for s in steps)
    assert Path(tmp_path / "AHU22000-D03-01_bend_probe.json").exists()
