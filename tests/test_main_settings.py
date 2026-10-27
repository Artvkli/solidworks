"""main.export_slddrw_files wiring: settings reach the exporter, test mode never touches production output."""
import json

import main
from fakesw import FakeApp, FlatPart, rect

MM = 0.001
PART = FlatPart([(x * MM, y * MM) for x, y in rect(100, 60)], [((25 * MM, 0), (25 * MM, 60 * MM))])


def scan(names):
    return {"sheet_metal_parts": [{"name": n, "path": f"/parts/{n}.SLDPRT", "quantity": 2, "thickness": 2.0}
                                  for n in names]}


def run(tmp_path, monkeypatch, **settings):
    for key, value in settings.items():
        monkeypatch.setattr(main, key, value)
    inp, out = tmp_path / "input", tmp_path / "output"
    inp.mkdir()
    out.mkdir()
    (inp / main.MASTER_DRAWING_NAME).write_bytes(b"master")
    app = FakeApp(PART)
    result = main.export_slddrw_files(app, tmp_path / "a.SLDASM", scan(["P1", "P2", "P3"]), out, inp)
    return result, out, app


def test_limit_one_writes_only_to_the_test_folder(tmp_path, monkeypatch):
    result, out, app = run(tmp_path, monkeypatch, SLDDRW_LIMIT=1, SLDDRW_ONLY_PARTS=None)
    assert [r["part"] for r in result["created"]] == ["P1"]
    assert (out / "flat_drawings_test" / "P1.SLDDRW").exists()
    assert (out / "flat_drawings_test" / "P1_bend_probe.json").exists()
    assert not (out / "flat_drawings").exists()               # production folder untouched
    summary = json.loads((out / "flat_drawings_test" / "slddrw_summary.json").read_text(encoding="utf-8"))
    assert any("bend dimensions: 1/1 created and verified" in s for s in summary["created"][0]["steps"])


def test_only_parts_selects_a_representative_part(tmp_path, monkeypatch):
    result, out, _ = run(tmp_path, monkeypatch, SLDDRW_LIMIT=None, SLDDRW_ONLY_PARTS=("p3",))
    assert [r["part"] for r in result["created"]] == ["P3"]
    assert (out / "flat_drawings_test" / "P3.SLDDRW").exists()


def test_production_mode_uses_the_normal_folder(tmp_path, monkeypatch):
    result, out, app = run(tmp_path, monkeypatch, SLDDRW_LIMIT=None, SLDDRW_ONLY_PARTS=None)
    assert len(result["created"]) == 3 and (out / "flat_drawings" / "P2.SLDDRW").exists()
    names = [n for n, _ in app.drawings[0].views]
    assert names == ["flat", "*Top", "*Front", "*Right"]          # FOLDED_VIEWS default reaches the exporter
