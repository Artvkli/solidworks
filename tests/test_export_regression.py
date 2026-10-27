"""End-to-end regression of SwDrawingExporter.export_part against a fake SolidWorks application.

Checks the drawing still gets: flat pattern + folded views (Top / Front / Right), overall length/width
dimensions per view, verified bend dimensions with tags, QTY note, bend table, saved .SLDDRW.
"""
import json

import pytest

from solidworks.sw_drawings import SCALES, SwDrawingExporter
from fakesw import FakeApp, FlatPart, rect

MM = 0.001


def mm(*pts):
    return [(x * MM, y * MM) for x, y in pts]


PART = FlatPart(mm(*rect(180, 120)),
                [mm((30, 0), (30, 120)), mm((150, 0), (150, 120)), mm((0, 100), (180, 100)), mm((0, 20), (180, 20))],
                holes=[mm(*rect(20, 30, 80, 45))])
BOX = (0.12, 0.08, 0.02)   # folded part L x D x H


def export(tmp_path, folded=("*Top", "*Front", "*Right"), **kw):
    master = tmp_path / "master.slddrw"
    master.write_bytes(b"master")
    app = FakeApp(PART, box=BOX)
    exporter = SwDrawingExporter(app, master_drawing=master, folded_view_names=folded,
                                 bend_debug_json=True, text_replacements={},
                                 bend_table_remove_words=("factor", "allowance", "deduction"), **kw)
    out = tmp_path / "out" / "PART-1.SLDDRW"
    item = {"name": "PART-1", "path": str(tmp_path / "PART-1.SLDPRT"), "quantity": 3, "thickness": 2.0}
    ok, message, steps = exporter.export_part(item, out)
    return ok, message, steps, app, out


def test_full_export_contains_all_required_content(tmp_path):
    ok, message, steps, app, out = export(tmp_path)
    assert ok, (message, steps)
    assert out.exists()
    drawing = app.drawings[0]
    names = [n for n, _ in drawing.views]
    assert names == ["flat", "*Top", "*Front", "*Right"]

    overall = [d for d in app.world.dims if d.kind in ("horizontal", "vertical")]
    assert len(overall) == 8                                    # length + width for each of the 4 views
    values = sorted(round(d.value / MM, 3) for d in overall)
    assert values == sorted([180.0, 120.0, 120.0, 80.0, 120.0, 20.0, 80.0, 20.0])

    bend = sorted(round(d.value / MM, 3) for d in app.world.dims if d.kind == "bend")
    assert bend == [20.0, 20.0, 30.0, 30.0]                    # equal values on different bends all kept
    assert any("bend dimensions: 4/4 created and verified" in s for s in steps)
    assert any("bend tags paired: 4 of 4" in s for s in steps)
    assert any(s.startswith("QTY note added: QTY=3") for s in steps)
    assert any("bend table tags: ['A', 'B', 'C', 'D']" in s for s in steps)
    assert any("K-Factor" in s and "removed" in s for s in steps)   # existing table cleanup still runs


def test_bend_dimensions_stay_on_the_sheet_and_clear_of_the_table(tmp_path):
    ok, _, steps, app, out = export(tmp_path)
    data = json.loads(out.with_name("PART-1_bend_probe.json").read_text(encoding="utf-8"))
    placements = data["plan"]["placements"]
    assert len(placements) == 4 and all(p["collisions"] == 0 for p in placements)
    table_top = (66 + 0) * MM + 5 * 0.006   # bend_table_bottom_mm + 1 header + 4 rows of the fake table
    for p in placements:
        x0, y0, x1, y1 = p["text_box"]
        assert 0.010 <= x0 and x1 <= 0.420 - 0.010
        assert y0 >= table_top - 1e-9 and y1 <= 0.297 - 0.010
    # folded views do not overlap each other or the flat view
    outlines = [v.GetOutline() for _, v in app.drawings[0].views]
    for i, a in enumerate(outlines):
        for b in outlines[i + 1:]:
            assert not (a[0] < b[2] and a[2] > b[0] and a[1] < b[3] and a[3] > b[1])


def test_layout_reserves_room_right_of_the_flat_pattern(tmp_path):
    ok, _, steps, app, out = export(tmp_path)
    flat = dict(app.drawings[0].views)["flat"].GetOutline()
    data = json.loads(out.with_name("PART-1_bend_probe.json").read_text(encoding="utf-8"))
    right = [p for p in data["plan"]["placements"] if p["side"] == "right"]
    assert right, "horizontal bend lines should be dimensioned on the right"
    assert all(p["text_box"][2] <= 0.420 - 0.010 for p in right)
    assert flat[2] < 0.420 - 0.010 - 0.018 + 1e-9             # at least the reserved right-hand room is free


def test_two_view_configuration_keeps_the_old_stacked_layout(tmp_path):
    ok, _, steps, app, _ = export(tmp_path, folded=("*Top", "*Front"))
    assert ok
    views = dict(app.drawings[0].views)
    top, front = views["*Top"].GetOutline(), views["*Front"].GetOutline()
    assert (top[0] + top[2]) / 2 == pytest.approx((front[0] + front[2]) / 2)   # same column
    assert top[1] > front[3]                                                   # Top above Front


def test_right_view_sits_right_of_front(tmp_path):
    ok, _, steps, app, _ = export(tmp_path)
    views = dict(app.drawings[0].views)
    front, right, top = (views[k].GetOutline() for k in ("*Front", "*Right", "*Top"))
    assert right[0] > front[2]                                                 # to the right
    assert (right[1] + right[3]) / 2 == pytest.approx((front[1] + front[3]) / 2)
    assert top[1] > front[3]


def test_failure_does_not_leave_a_partial_file(tmp_path):
    master = tmp_path / "missing.slddrw"
    app = FakeApp(PART, box=BOX)
    exporter = SwDrawingExporter(app, master_drawing=master)
    out = tmp_path / "x.SLDDRW"
    ok, message, _ = exporter.export_part({"name": "x", "path": str(tmp_path / "x.SLDPRT")}, out)
    assert not ok and "master drawing not found" in message and not out.exists()


# ------------------------------------------------------------------ earlier helpers keep working
def test_fit_scale_and_scale_text_regressions():
    ex = SwDrawingExporter(None)
    assert ex._fit_scale(0.37) == pytest.approx(1 / 3)
    assert ex._fit_scale(3.0) == 2 and ex._fit_scale(1.0) == 1
    assert ex._fit_scale(0) == SCALES[-1] and ex._fit_scale(-1) == SCALES[-1]
    assert SwDrawingExporter(None, layout={"scale_mode": "standard"})._fit_scale(0.37) == pytest.approx(1 / 3)
    assert ex._scale_text(1 / 3) == "1:3" and ex._scale_text(2) == "2:1"
    assert SwDrawingExporter._safe_name('a<b>:c.') == "a_b__c"


def test_old_order_names_still_accepted():
    for name in ("largest_first", "smallest_first", "largest_outside", "smallest_outside"):
        SwDrawingExporter(None, bend_dim_order=name)
