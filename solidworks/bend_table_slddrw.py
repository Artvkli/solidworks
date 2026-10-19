"""
solidworks/bend_table_slddrw.py

The sheets of the bend table PDF (flat view, folded views, dimensions, bend tags, bend table, title block)
written as a native SolidWorks drawing: ONE .slddrw per thickness, ONE A4 sheet per part.

How: SwCanvas imitates the reportlab canvas methods the PDF code uses (line, rect, paths, text ...), but
creates SolidWorks objects instead:
    lines / outlines / arrows / dimension lines  -> sketch segments on the sheet (on layers OUTLINE / THIN / BEND)
    texts (numbers, tags, table cells, title block) -> SolidWorks notes
So the PDF drawing code (FlatPatternView.draw, draw_folded, BendTablePdfWriter._draw_page) is reused unchanged.

NOTE: the result is a "dumb" drawing (not linked to the 3D model), exactly like the PDF is.
"""
import math
import time
from pathlib import Path

from solidworks.bend_table_dxf import DxfCanvas, MM, A4_PT
from solidworks.bend_table_pdf import BendTablePdfWriter, safe_name

TEXT_ASCENT = 0.9      # note box top = baseline + 0.9 * font size (tune if texts sit too high / low)
FONT_NAME = "Arial"


class SwCanvas(DxfCanvas):
    """The subset of reportlab's Canvas API used by the bend table drawing code, drawing into SolidWorks."""

    def __init__(self, reader, drawing):
        # (no ezdxf document here: only the transform / state helpers of DxfCanvas are reused)
        self.reader = reader
        self.drawing = drawing
        self._pagesize = A4_PT
        self._ctm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        self._stack = []
        self._dash = False
        self._lw = 0.7
        self._font = ("Helvetica", 10.0)

        self.c = reader._constants()
        self.sm = drawing.SketchManager
        self._first_sheet = True
        self._sheet_ready = False
        self._names = set()
        self.sheet_name = "part"
        self._page_in_item = 0
        self._layer = None
        self.errors = 0
        self.first_error = None
        self.sheets = 0
        self._make_layers()
        try:
            self.sm.AddToDB = True              # much faster: no screen refresh per segment
            self.sm.DisplayWhenAdded = False
        except Exception:
            pass

    # ---------- helpers ----------
    def _const(self, name, default):
        try:
            return getattr(self.c, name)
        except Exception:
            return default

    def _fail(self, what, exc):
        self.errors += 1
        if self.first_error is None:
            self.first_error = f"{what}: {type(exc).__name__}: {exc}"

    def _pt(self, x, y):
        a, b, c, d, e, f = self._ctm
        return ((a * x + c * y + e) * MM / 1000.0, (b * x + d * y + f) * MM / 1000.0)   # metres

    def _make_layers(self):
        cont = self._const("swLineCONTINUOUS", 0)
        dashed = self._const("swLineHIDDEN", 1)
        thin = self._const("swLW_THIN", 0)
        normal = self._const("swLW_NORMAL", 1)
        thick = self._const("swLW_THICK", 2)
        for name, style, weight in (("OUTLINE", cont, thick), ("THIN", cont, thin), ("BEND", dashed, normal)):
            try:
                self.drawing.CreateLayer2(name, "", 0, style, weight, True, False)
            except Exception as exc:
                self._fail(f"CreateLayer2({name})", exc)

    def _use_layer(self):
        name = "BEND" if self._dash else ("OUTLINE" if self._lw >= 0.6 else "THIN")
        if name != self._layer:
            try:
                self.drawing.SetCurrentLayer(name)
            except Exception as exc:
                self._fail("SetCurrentLayer", exc)
            self._layer = name

    # ---------- sheets ----------
    def begin_item(self, name):
        self.sheet_name = safe_name(name)[:60]
        self._page_in_item = 0

    def _unique_sheet_name(self):
        self._page_in_item += 1
        base = self.sheet_name if self._page_in_item == 1 else f"{self.sheet_name}_{self._page_in_item}"
        name, n = base, 2
        while name.lower() in self._names:
            name = f"{base}_{n}"
            n += 1
        self._names.add(name.lower())
        return name

    def _ensure_sheet(self):
        if self._sheet_ready:
            return
        name = self._unique_sheet_name()
        paper = self._const("swDwgPaperA4sizeVertical", 7)
        none = self._const("swDwgTemplateNone", 7)
        width, height = A4_PT[0] * MM / 1000.0, A4_PT[1] * MM / 1000.0

        if self._first_sheet:
            sheet = self.drawing.GetCurrentSheet()
            sheet.SetProperties2(paper, none, 1.0, 1.0, True, "", width, height, True)
            sheet.SetName(name)
        else:
            sheet = None
            for property_view in ("Default", ""):
                try:
                    sheet = self.drawing.NewSheet4(
                        name, paper, none, 1.0, 1.0, True, "", width, height, property_view, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
                    )
                except Exception as exc:
                    self._fail("NewSheet4", exc)
                    sheet = None
                if sheet is not None:
                    break
            if sheet is None:
                raise RuntimeError(f"cannot create sheet {name}: {self.first_error}")
            try:
                self.drawing.ActivateSheet(name)
            except Exception:
                pass
        self._sheet_ready = True
        self.sheets += 1

    def showPage(self):
        self._sheet_ready = False
        self._first_sheet = False

    def finish(self):
        try:
            self.sm.AddToDB = False
            self.sm.DisplayWhenAdded = True
        except Exception:
            pass

    # ---------- drawing ----------
    def _seg(self, p, q):
        if math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-6:
            return
        self._ensure_sheet()
        self._use_layer()
        try:
            self.sm.CreateLine(p[0], p[1], 0.0, q[0], q[1], 0.0)
        except Exception as exc:
            self._fail("CreateLine", exc)

    def line(self, x1, y1, x2, y2):
        self._seg(self._pt(x1, y1), self._pt(x2, y2))

    def rect(self, x, y, w, h, stroke=1, fill=0):
        pts = [self._pt(x, y), self._pt(x + w, y), self._pt(x + w, y + h), self._pt(x, y + h)]
        for i in range(4):
            self._seg(pts[i], pts[(i + 1) % 4])

    def drawPath(self, path, stroke=1, fill=0):
        # (arrow heads, filled in the PDF, become small closed triangles here)
        for sub in path.subpaths:
            pts = [self._pt(x, y) for x, y in sub["pts"]]
            for i in range(len(pts) - 1):
                self._seg(pts[i], pts[i + 1])
            if sub["closed"] and len(pts) > 2:
                self._seg(pts[-1], pts[0])

    def _text(self, x, y, text, anchor):
        if not text:
            return
        self._ensure_sheet()
        name, size = self._font
        width = self.stringWidth(str(text), name, size)
        start = x - width / 2.0 if anchor == "c" else (x - width if anchor == "r" else x)

        a, b = self._ctm[0], self._ctm[1]
        scale = math.hypot(a, b) or 1.0
        top_left = self._pt(start, y + size * TEXT_ASCENT)
        try:
            note = self.drawing.InsertNote(str(text))
            if note is None:
                raise RuntimeError("InsertNote returned None")
            annotation = note.GetAnnotation()
            fmt = annotation.GetTextFormat(0)
            fmt.TypeFaceName = FONT_NAME
            fmt.CharHeightInPts = size * scale
            fmt.Bold = name.endswith("Bold")
            annotation.SetTextFormat(0, False, fmt)
            annotation.SetPosition2(top_left[0], top_left[1], 0.0)
            angle = math.atan2(b, a)
            if abs(angle) > 1e-9:
                note.Angle = angle
        except Exception as exc:
            self._fail("note", exc)

    def drawString(self, x, y, text):
        self._text(x, y, text, "l")

    def drawCentredString(self, x, y, text):
        self._text(x, y, text, "c")

    def drawRightString(self, x, y, text):
        self._text(x, y, text, "r")


class SlddrwWriter(BendTablePdfWriter):
    """Same pages as the PDF, saved as .slddrw: one file per thickness, one sheet per part."""

    def __init__(self, reader, group_by_thickness=True, **kwargs):
        super().__init__(**kwargs)
        self.reader = reader
        self.group_by_thickness = group_by_thickness

    @staticmethod
    def _tag(value):
        return "unknown_thickness" if value is None else f"{value:g}mm"

    def write_all_slddrw(self, items, output_folder, base_name="bend_tables"):
        print("\n" + "=" * 60 + "\nBEND TABLE DRAWINGS (.slddrw, by thickness)\n" + "=" * 60)
        reader = self.reader
        reader._resolve_templates([])
        if not reader.drawing_template:
            print("No drawing template found (set DRAWING_TEMPLATE_PATH).")
            return {"files": [], "error": "no template"}
        if reader._constants() is None:
            print("SolidWorks constants (swconst.tlb) could not be loaded - cannot create sheets.")
            return {"files": [], "error": "no constants"}

        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        with_table = [i for i in items if (i.get("bend_table") or {}).get("rows")]
        groups = {}
        for item in with_table:
            groups.setdefault(item.get("thickness") if self.group_by_thickness else None, []).append(item)

        files = []
        for key in sorted(groups, key=lambda k: (k is None, k or 0)):
            group = sorted(groups[key], key=lambda i: i["name"])
            tag = self._tag(key) if self.group_by_thickness else "all"
            path = output_folder / f"{safe_name(base_name)}_{tag}.slddrw"
            started = time.time()
            drawing = None
            try:
                drawing = reader.sw_app.NewDocument(reader.drawing_template, 0, 0.0, 0.0)
                if drawing is None:
                    raise RuntimeError("NewDocument returned None")
                steps = []
                reader._use_millimeters(drawing, steps)

                canvas = SwCanvas(reader, drawing)
                for item in group:
                    canvas.begin_item(item["name"])
                    before = len(self.messages)
                    self._draw_item(canvas, item)
                    for line in self.messages[before:]:
                        print(f"    {line}")
                canvas.finish()

                if path.exists():
                    path.unlink()
                saved = drawing.SaveAs3(str(path), 0, 1)
                ok = path.exists() and path.stat().st_size > 0
                print(
                    f"{tag} -> {len(group)} part(s), {canvas.sheets} sheet(s) -> {path.name} "
                    f"({'saved' if ok else f'NOT saved, SaveAs3={saved!r}'}, {time.time() - started:.0f} s)"
                )
                if canvas.errors:
                    print(f"    WARNING: {canvas.errors} SolidWorks call(s) failed, first: {canvas.first_error}")
                if ok:
                    files.append({"thickness": key, "path": path, "parts": len(group), "sheets": canvas.sheets})
            except Exception as exc:
                print(f"{tag} -> FAILED: {type(exc).__name__}: {exc}")
            finally:
                if drawing is not None:
                    reader._close(drawing)

        print(f"\n.slddrw files written: {len(files)}")
        return {"files": files, "error": None}