"""Pure geometry: transforms, outer boundary, bend -> outer edge measurement, tags, placement."""
import math

import pytest

from solidworks.bend_geometry import (Xform, assign_tags, classify_outer, detect_convention, measure_bends,
                                      order_key, plan_placements)
from fakesw import rect

MM = 0.001


def edges_of(*polys):
    out = []
    for poly in polys:
        for i in range(len(poly)):
            out.append({"id": len(out), "p0": poly[i], "p1": poly[(i + 1) % len(poly)], "is_line": True})
    return out


def run(polys, bends, merge=True):
    edges = edges_of(*polys)
    outer, info = classify_outer([(e["p0"], e["p1"]) for e in edges])
    res = measure_bends([{"id": f"#{i + 1}", "p0": a, "p1": b} for i, (a, b) in enumerate(bends)],
                        edges, outer, info["polygon"], merge_collinear=merge)
    return res, edges, outer, info


def mm(poly):
    return [(x * MM, y * MM) for x, y in poly]


# ------------------------------------------------------------------ transforms
def test_xform_inverse_and_composition():
    a = math.radians(37)
    X = Xform([[math.cos(a), math.sin(a), 0], [-math.sin(a), math.cos(a), 0], [0, 0, 1]], (0.1, -0.2, 0.3), 0.25)
    Y = Xform([[0, 0, 1], [1, 0, 0], [0, 1, 0]], (1, 2, 3), 2.0)
    p = (0.3, 0.7, -0.2)
    assert X.inverse().apply(X.apply(p)) == pytest.approx(p, abs=1e-12)
    assert X.then(Y).apply(p) == pytest.approx(Y.apply(X.apply(p)), abs=1e-12)
    assert Xform.from_array(X.to_array()).apply(p) == pytest.approx(X.apply(p), abs=1e-15)


def test_row_vector_convention_matches_documented_layout():
    # rows of R are the images of the axes: x-axis -> (0, 1, 0), plus translation, plus scale 2
    data = (0, 1, 0, -1, 0, 0, 0, 0, 1, 5, 6, 7, 2, 0, 0, 0)
    assert Xform.from_array(data, "row").apply((1, 0, 0)) == pytest.approx((5, 8, 7))
    assert Xform.from_array(data, "column").apply((1, 0, 0)) == pytest.approx((5, 4, 7))


def test_detect_convention():
    X = Xform([[0.6, 0.8, 0], [-0.8, 0.6, 0], [0, 0, 1]], (0.01, 0.02, 0.03), 0.5)
    probe = (0.1234, -0.5678, 0.9012)
    assert detect_convention(X.to_array(), probe, X.apply(probe)) == "row"
    transposed = Xform.from_array(X.to_array(), "column").to_array()
    assert detect_convention(transposed, probe, X.apply(probe)) == "column"
    assert detect_convention(X.to_array(), probe, (9, 9, 9)) is None


# ------------------------------------------------------------------ outer boundary + measurement
def test_outer_boundary_excludes_holes():
    _, edges, outer, info = run([mm(rect(100, 60)), mm(rect(10, 20, 35, 20))], [])
    assert outer == {0, 1, 2, 3}
    assert info["closed"] and info["method"] == "largest closed edge loop"


def test_cutout_edge_is_not_used_as_reference():
    # bend at x=30; a slot edge at x=35 is closer (5 mm) but is NOT the outer boundary
    res, *_ = run([mm(rect(100, 60)), mm(rect(10, 20, 35, 20))], [(mm([(30, 0)])[0], mm([(30, 60)])[0])])
    assert res[0]["status"] == "ok"
    assert res[0]["value"] == pytest.approx(30 * MM)


def test_outer_edge_beyond_a_notch_is_rejected():
    # U shape: notch 40..60 wide from y=20 up. Bend at x=85 in the right leg.
    u_shape = mm([(0, 0), (100, 0), (100, 60), (60, 60), (60, 20), (40, 20), (40, 60), (0, 60)])
    res, edges, *_ = run([u_shape], [(mm([(85, 0)])[0], mm([(85, 60)])[0])])
    r = res[0]
    assert r["status"] == "ok" and r["value"] == pytest.approx(15 * MM)        # right edge
    left_edge = next(i for i, e in enumerate(edges) if e["p0"] == u_shape[7])  # (0,60)->(0,0)
    assert any(x["edge"] == left_edge and "in between" in x["why"] for x in r["rejected"])


def test_equal_values_on_different_bends_are_all_kept():
    res, *_ = run([mm(rect(100, 60))], [(mm([(20, 0)])[0], mm([(20, 60)])[0]),
                                       (mm([(80, 60)])[0], mm([(80, 0)])[0])])
    assert [r["status"] for r in res] == ["ok", "ok"]
    assert [round(r["value"] / MM, 3) for r in res] == [20.0, 20.0]
    assert res[0]["edge"] != res[1]["edge"]


def test_duplicates_and_collinear_bends_share_one_dimension():
    plate = mm(rect(100, 60))
    a = (mm([(20, 0)])[0], mm([(20, 25)])[0])
    b = (mm([(20, 35)])[0], mm([(20, 60)])[0])   # same line, interrupted -> same reference edge
    res, *_ = run([plate], [a, b, a])
    assert [r["status"] for r in res] == ["ok", "merged", "duplicate"]
    assert res[0]["shared_with"] == ["#2", "#3"]
    res2, *_ = run([plate], [a, b], merge=False)
    assert [r["status"] for r in res2] == ["ok", "ok"]


def test_no_parallel_outer_edge_is_reported_not_guessed():
    trapezoid = mm([(0, 0), (100, 0), (80, 60), (20, 60)])
    res, *_ = run([trapezoid], [(mm([(50, 0)])[0], mm([(50, 60)])[0])])
    assert res[0]["status"] == "no_reference_edge" and res[0]["value"] is None


def test_bend_outside_boundary_is_flagged():
    res, *_ = run([mm(rect(100, 60))], [(mm([(150, 0)])[0], mm([(150, 60)])[0])])
    assert res[0]["status"] == "outside_boundary"


@pytest.mark.parametrize("angle", [0, 17, 45, 90, 133, 270])
def test_oblique_bends_measure_perpendicular_distance(angle):
    c, s = math.cos(math.radians(angle)), math.sin(math.radians(angle))
    rot = lambda pts: [(c * x - s * y + 0.05, s * x + c * y - 0.02) for x, y in pts]  # noqa: E731
    plate = rot(mm(rect(100, 60)))
    bend = rot(mm([(70, 0), (70, 60)]))
    res, *_ = run([plate], [(bend[0], bend[1])])
    assert res[0]["status"] == "ok"
    assert res[0]["value"] == pytest.approx(30 * MM, abs=1e-9)


def test_chamfered_outer_edge_parallel_to_oblique_bend():
    # 45 deg bend parallel to a 45 deg chamfer: distance measured perpendicular to both
    poly = mm([(0, 0), (100, 0), (100, 40), (60, 80), (0, 80)])
    bend = mm([(40, 80), (100, 20)])  # parallel to chamfer (100,40)-(60,80), offset inward
    res, *_ = run([poly], [(bend[0], bend[1])])
    assert res[0]["status"] == "ok"
    assert res[0]["value"] == pytest.approx(20 / math.sqrt(2) * MM, abs=1e-9)


# ------------------------------------------------------------------ tags
def test_assign_tags_nearest_and_unique():
    lines = {"#1": ((0, 0), (0, 1)), "#2": ((0.05, 0), (0.05, 1))}
    notes = [("B", (0.004, 0.5)), ("A", (0.047, 0.5)), ("Fixed face", (0.02, 0.5))]
    tags, unassigned, ambiguous = assign_tags(lines, notes, ["A", "B"], 0.015)
    assert tags == {"#1": "B", "#2": "A"} and unassigned == [] and ambiguous == {}


def test_assign_tags_respects_max_distance():
    tags, unassigned, _ = assign_tags({"#1": ((0, 0), (0, 1))}, [("A", (0.5, 0.5))], ["A"], 0.015)
    assert tags == {} and unassigned == ["A"]


def test_assign_tags_flags_ambiguous_pairs_instead_of_guessing():
    # tag "B" sits 3 mm from its line but only 3.5 mm from the neighbouring one
    lines = {"#1": ((0, 0), (0, 1)), "#2": ((0.0065, 0), (0.0065, 1))}
    tags, _, ambiguous = assign_tags(lines, [("A", (-0.003, 0.5)), ("B", (0.0035, 0.5))], ["A", "B"], 0.015)
    assert "#2" in ambiguous and "#2" not in tags


# ------------------------------------------------------------------ placement
def _boxes_overlap(a, b):
    return a[0] < b[2] and a[2] > b[0] and a[1] < b[3] and a[3] > b[1]


def test_dense_bends_are_stacked_without_text_collisions():
    part = mm(rect(100, 60))
    xs = [10, 13, 16, 19, 50, 53]       # distances measured from the left edge x=0 -> spans overlap
    dims = [{"key": f"#{i}", "b0": (x * MM, 0), "b1": (x * MM, 60 * MM), "e0": (0, 0), "e1": (0, 60 * MM)}
            for i, x in enumerate(xs)]
    placements, needs = plan_placements(dims, part, 8 * MM, 7 * MM, 14 * MM, 3.5 * MM)
    assert all(p["collisions"] == 0 for p in placements)
    assert all(p["side"] == "bottom" for p in placements)  # vertical bend lines -> dimensions below
    for i, p in enumerate(placements):
        for q in placements[i + 1:]:
            assert not _boxes_overlap(p["text_box"], q["text_box"])
    assert len({p["level"] for p in placements}) > 1
    assert needs["bottom"] > 8 * MM and needs["right"] == 0.0


def test_horizontal_bends_go_to_the_right_and_avoid_obstacles():
    part = mm(rect(100, 60))
    dims = [{"key": "#1", "b0": (0, 20 * MM), "b1": (100 * MM, 20 * MM), "e0": (0, 0), "e1": (100 * MM, 0)}]
    free, _ = plan_placements(dims, part, 8 * MM, 7 * MM, 14 * MM, 3.5 * MM)
    assert free[0]["side"] == "right" and free[0]["level"] == 0
    # an obstacle (e.g. a bend tag) where level 0 would put the text -> moved out to the next level
    tag = tuple(free[0]["text_box"])
    blocked, _ = plan_placements(dims, part, 8 * MM, 7 * MM, 14 * MM, 3.5 * MM, obstacles=[tag])
    assert blocked[0]["level"] == 1 and blocked[0]["collisions"] == 0


def test_placement_reports_unavoidable_collision_when_out_of_bounds():
    part = mm(rect(100, 60))
    dims = [{"key": "#1", "b0": (20 * MM, 0), "b1": (20 * MM, 60 * MM), "e0": (0, 0), "e1": (0, 60 * MM)}]
    tight = (0, 0, 0.1, 0.06)            # no room below the part at all
    placed, _ = plan_placements(dims, part, 8 * MM, 7 * MM, 14 * MM, 3.5 * MM, max_levels=3, bounds=tight)
    assert placed[0]["collisions"] > 0


def test_order_key_names_and_aliases():
    items = [{"id": "#1", "value": 3}, {"id": "#2", "value": 1}, {"id": "#3", "value": 2}]
    assert [i["id"] for i in sorted(items, key=order_key("largest_outside"))] == ["#2", "#3", "#1"]
    assert [i["id"] for i in sorted(items, key=order_key("largest_first"))] == ["#2", "#3", "#1"]
    assert [i["id"] for i in sorted(items, key=order_key("smallest_outside"))] == ["#1", "#3", "#2"]
    assert [i["id"] for i in sorted(items, key=order_key("smallest_first"))] == ["#1", "#3", "#2"]
