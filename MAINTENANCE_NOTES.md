# Maintenance notes

Phase 1 (cleanup) and phase 2 (bend dimensions), 2026-10-27.

Nothing here has been run against a live SolidWorks; this environment has no Windows/COM.
All verification was offline (syntax, pyflakes/pylint, imports with stubbed pywin32, mock-COM unit checks).

## Active pipeline (what `python main.py` actually uses)

```
main.py
 ├─ solidworks/assembly.py      SolidWorksAssembly, find_assemblies  (→ component.py)
 ├─ solidworks/sheet_metal.py   SheetMetalDetector  (scan, thickness, mirror merge, CSV)
 ├─ solidworks/sw_drawings.py   SwDrawingExporter   (one .SLDDRW per part)
 │   └─ solidworks/bend_geometry.py  COM-free geometry for bend dimensions (phase 2)
 ├─ solidworks/dwg_exporter.py  DwgExporter  (always built; only activate_assembly() runs
 │                                            while SLDDRW_ONLY = True)
 └─ solidworks/thickness_dwg.py ThicknessDwgBuilder (only when SLDDRW_ONLY = False; needs ezdxf)
```
Runtime inputs: `input/*.SLDASM` (+ referenced parts) and `input/master.slddrw` (required).
Outputs: `output/flat_drawings/*.SLDDRW`, `slddrw_summary.json`, `scan_cache.json`,
`<asm>_sheet_metal.csv`, `run_summary.json`. All settings live at the top of `main.py`.

## Inactive modules (not imported by main.py — kept, not deleted)

| File | What it is | Note |
|---|---|---|
| `config.py` | Old settings module | Nothing imports it; `main.py` has its own constants |
| `solidworks/flat_view.py` | ReportLab PDF flat-pattern renderer | Has `match_bend_lines()` (8-orientation sketch→sheet fit) – useful reference for the next phase |
| `solidworks/bend_table.py` | BendTableReader (temp drawing → read bend table, probe JSON/DXF) | `_probe_view()` is a ready-made diagnostic dump of view/bend-line data |
| `solidworks/bend_table_pdf.py` | Bend-table PDF writer | uses flat_view |
| `solidworks/pdf_builder.py` | matplotlib PDF builder from DXF | |
| `solidworks/autocad.py` | AutoCAD COM DWG assembler | |
| `solidworks/connection.py`, `solidworks/utils.py` | Tiny helpers | unused; `com_value` is duplicated as `_get` in every module |
| `solidworks/test-export.py` | Stand-alone debug script, hard-coded `F:\...` paths | hyphen in name → not importable as a module |

`CHANGES_FA.txt` and the README "UPDATE" section describe an **older** sw_drawings
(`SwDrawingPdfExporter`, `_insert_dimensions`, PDF export). The current class is
`SwDrawingExporter` and does not export PDF.

## sw_drawings.py – flow (`_export_once`)

1. open part, activate configuration, rebuild, find Flat-Pattern feature
2. copy `master.slddrw` → open → `_apply_sheet_format` (custom props + placeholder notes)
3. `_remove_existing_views`
4. `_create_flat_view`, then the folded views listed in `FOLDED_VIEWS` (default Top, Front, Right)
5. `_insert_bend_table` (column cleanup, move bottom-right, **reads the tag column**)
6. `_analyze_bends` – measure every bend (layout-independent, done once)
7. `_layout_views` – folded views in a grid (Top over Front, Right beside Front), flat pattern with room
   reserved below/right for bend dimensions; re-run up to 3× with the room the planned dimensions need
8. `_plan_bend_dims` – map to sheet, pair tags, collision-aware placement
9. overall length/width dimensions for every view
10. `_finish_bend_dims` – create, read back and verify, report every bend not dimensioned
11. `<part>_bend_probe.json` (if `BEND_DEBUG_JSON`), QTY note, rebuild, SaveAs3

## Phase 2 – bend dimensions: what was wrong and what changed

**Root cause (verified in code and against the API docs).** `IView.GetBendLines()` returns sketch segments
whose `GetStartPoint2/GetEndPoint2` X/Y/Z are in the bend-line sketch's *local* frame. The old code pushed
them through `IView.ModelToViewTransform`, which expects *model* coordinates, while the outer edges
(`GetVisibleEntities2`) are in model space. Bend lines and edges were in different frames, so the
nearest-edge choice and the value were wrong (this is the ±1.044 vs −0.354…0.654 mismatch in CHANGES_FA §4).
The unused `via` parameter of the old mapper was meant for the missing sketch→model step but was never passed.
Second defect: the old fallback mapper ignored the translation and picked "transpose or not" from the
aspect ratio – undefined for square parts and 90°-rotated views.

**Coordinate chain now.**
* measuring frame = the first bend line's sketch frame (true size, sheet plane). Edges go model → frame with
  `ISketch.ModelToSketchTransform`; bend lines from another sketch go sketch_i → model → frame.
  **Dimension values never use the drawing-view transform or view scale.**
* checks: all edge end points must lie within the sheet thickness of the frame plane, and bend lines must lie
  inside the outer boundary – otherwise bend dimensions are aborted for that part with a message.
* sheet placement = frame → model → `ModelToViewTransform`; validated against `IView.GetOutline` (a uniform
  scale / position difference is corrected and logged; a one-axis size mismatch aborts placement).
* MathTransform matrix convention (`p' = s·(p·R) + T`) is verified once per session by letting SolidWorks
  transform a probe point (`IMathPoint.MultiplyTransform`); no aspect-ratio guessing anywhere.

**Edge selection.** Outer boundary = loop with `ILoop2.IsOuter()` of the largest planar face parallel to the
sheet (fallback: largest closed loop of the visible edges). A reference edge must be a straight outer edge,
parallel to the bend (±0.5°), overlap the bend line along its direction, and be reachable without crossing
another outer edge (edges beyond a notch are rejected). Holes/slots/cut-outs are never references.
Nearest such edge wins; equal distances on both sides are flagged as `tie`. No candidate → bend reported as
`no_reference_edge`, not dimensioned.

**Identity.** Bends are keyed by their position in `GetBendLines` (`#1, #2 …`). Equal values never merge.
Exact duplicates share one dimension; collinear bends measured to the same edge share one dimension when
`BEND_DIM_MERGE_COLLINEAR = True`. Every bend not dimensioned gets its own log line with the reason.

**Verification.** Each created dimension is read back (`GetDimension2(0).SystemValue`) and compared with the
computed value (0.01 mm). Mismatch → logged `MISMATCH` (`BEND_DIM_VERIFY = "flag"`) or also deleted (`"delete"`).

**Tags.** Tag names come from the bend table's tag column; tag notes of the flat view are paired with the
nearest bend line. A pairing is reported as AMBIGUOUS (not used) when another bend line is about as close.
Tags only label the log/probe: the dimension itself is attached to the bend line entity.

**Placement.** Dimension line perpendicular to the bend, beyond the part along the bend direction: below
for bends that are vertical on the sheet, right for horizontal ones, generic direction for oblique ones.
Levels assigned deterministically (`BEND_DIM_ORDER`: `largest_outside` = smallest nearest the part;
old names `largest_first`/`smallest_first` still accepted). Collisions checked between text boxes, dimension
lines, extension lines, bend tags, the overall dimensions, folded views and the sheet/table bounds; narrow
dimensions put the text just outside the extension lines. Unavoidable overlaps are counted in the log.

## Tests

`python -m pytest tests` (133 tests, offline; pywin32 is stubbed when absent). Fake SolidWorks objects
(`tests/fakesw.py`) model part embedding, sketch frame (rotated/translated/mirrored), view scale/rotation/
position and the row-vector MathTransform with their own math. **They prove the Python logic only, not
the behaviour of a real SolidWorks drawing.**

## Not verified yet – needs a real SolidWorks run

Run with `SLDDRW_LIMIT = 1` (or `SLDDRW_ONLY_PARTS`), output goes to `output/flat_drawings_test/`. Check:
1. log line `transform convention verified ...` (if it says "unavailable", MathUtility could not be used)
2. `outer boundary: ... via face outer loop` (else the geometric fallback was used)
3. `bend view transform: matches the view outline` (a "corrected" note means ModelToViewTransform differs
   from what the docs suggest – still usable, but report it)
4. every bend line `verified`; no `MISMATCH`; no `NOT dimensioned` you did not expect
5. `bend tags paired: n of n` and visually: each dimension is attached to the right bend line / tag
6. open the .SLDDRW: extension lines on the bend line and the outer edge, no overlapping text

Assumptions that only a live run can confirm: `ISketchSegment.GetSketch` + `ModelToSketchTransform` for the
flat-pattern bend-line sketch; `Select4` on a bend-line segment + model edge with `SelectData.View` followed by
`AddDimension2` gives a parallel distance; bend note positions from `IView.GetNotes` are sheet coordinates;
`GetTwoAdjacentFaces2` / `IFace2.Normal` / `ILoop2.IsOuter` on edges of a flat-pattern view.

## Real run 1 (SolidWorks 2025, AHU22000-D03-01) – findings and fixes

Confirmed by the log:
* bend measurement chain works on real geometry: outer loop via `ILoop2.IsOuter` (36 edges, face 2.49 m²),
  frame check passed; scale sanity: 97.7 mm × 22 = 2149 mm, 52.7 mm × 22 = 1159 mm → 2.49 m².
* `GetVisibleEntities2` types: 1 = edges (64; 36 with vertices, the rest closed hole edges), 2 = vertices, 3 = face.
* `IView.GetOutline` = geometry + **4.2 mm border per side** (8.4 mm total) on every view and scale.
* `IMathPoint.MultiplyTransform` failed with `Member not found` (DISP_E_MEMBERNOTFOUND).

Fixed:
* outline check accepts an equal border in x and y (a real mismatch is still rejected);
* convention check runs once, tries the pywin32 call variants, names the failing call; if it cannot run,
  both matrix forms are evaluated and the geometry decides (wrong form → edges far out of the sketch plane);
  if both fit with different results it stops instead of guessing;
* overall dimensions never fall back to model coordinates; each one is read back and verified;
* view sizes for layout use two probe scales so the border is not counted as geometry (was +8 % here);
* flatness tolerance 2·t + 0.5 mm (robust to an estimated thickness; a wrong frame gives hundreds of mm).
