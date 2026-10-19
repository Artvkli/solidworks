"""
solidworks/bend_table_dxf.py

Writes the SAME sheets that BendTablePdfWriter draws on the PDF (flat view, folded views, dimensions,
bend tags, bend table, title block) into DXF / DWG files, one file per thickness.
Every sheet is an A4 page (210 x 297 mm, 1:1 with the PDF) placed in a grid in model space.

How: DxfCanvas imitates the few reportlab canvas methods that FlatPatternView and BendTablePdfWriter use,
but creates ezdxf entities instead. So the PDF drawing code is reused unchanged.
"""
import math
from pathlib import Path

from solidworks.bend_table_pdf import BendTablePdfWriter, safe_name

MM = 25.4 / 72.0  # PDF point -> mm
A4_PT = (595.2756, 841.8898)


class _Path:
    def __init__(self):
        self.subpaths = []

    def moveTo(self, x, y):
        self.subpaths.append({"pts": [(x, y)], "closed": False})

    def lineTo(self, x, y):
        if not self.subpaths:
            self.moveTo(x, y)
        else:
            self.subpaths[-1]["pts"].append((x, y))

    def close(self):
        if self.subpaths:
            self.subpaths[-1]["closed"] = True


class DxfCanvas:
    """The subset of reportlab's Canvas API used by the bend table drawing code."""

    def __init__(self, doc, pagesize=A4_PT, per_row=4, gap_mm=40.0):
        from ezdxf import const

        self.doc = doc
        self.msp = doc.modelspace()
        self._pagesize = pagesize
        self.per_row = per_row
        self.gap_mm = gap_mm
        self.index = 0                       # which sheet we are on (showPage moves to the next)
        self._ctm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        self._stack = []
        self._dash = False
        self._lw = 0.7
        self._font = ("Helvetica", 10.0)
        self._valid_lw = [w for w in const.VALID_DXF_LINEWEIGHTS if w >= 0]
        self.sheets = 0

    # ---------- geometry ----------
    def _origin(self):
        w, h = self._pagesize[0] * MM, self._pagesize[1] * MM
        col, row = self.index % self.per_row, self.index // self.per_row
        return col * (w + self.gap_mm), -row * (h + self.gap_mm)

    def _pt(self, x, y):
        a, b, c, d, e, f = self._ctm
        ox, oy = self._origin()
        return ((a * x + c * y + e) * MM + ox, (b * x + d * y + f) * MM + oy)

    def _attribs(self):
        attribs = {"linetype": "DASHED" if self._dash else "CONTINUOUS"}
        target = int(round(self._lw * MM * 100))
        attribs["lineweight"] = min(self._valid_lw, key=lambda w: abs(w - target))
        return attribs

    # ---------- state (reportlab names) ----------
    def saveState(self):
        self._stack.append((self._ctm, self._dash, self._lw, self._font))

    def restoreState(self):
        if self._stack:
            self._ctm, self._dash, self._lw, self._font = self._stack.pop()

    def translate(self, tx, ty):
        a, b, c, d, e, f = self._ctm
        self._ctm = (a, b, c, d, a * tx + c * ty + e, b * tx + d * ty + f)

    def rotate(self, degrees):
        a, b, c, d, e, f = self._ctm
        r = math.radians(degrees)
        cs, sn = math.cos(r), math.sin(r)
        self._ctm = (a * cs + c * sn, b * cs + d * sn, -a * sn + c * cs, -b * sn + d * cs, e, f)

    def setLineWidth(self, w):
        self._lw = w

    def setDash(self, *args):
        self._dash = bool(args)

    def setStrokeColorRGB(self, *a):
        pass

    def setFillColorRGB(self, *a):
        pass

    def setFont(self, name, size):
        self._font = (name, float(size))

    def stringWidth(self, text, font, size):
        from reportlab.pdfbase.pdfmetrics import stringWidth
        return stringWidth(text, font, size)

    def setTitle(self, *a):
        pass

    def showPage(self):
        self.index += 1

    def save(self):
        pass

    # ---------- drawing ----------
    def line(self, x1, y1, x2, y2):
        self.msp.add_line(self._pt(x1, y1), self._pt(x2, y2), dxfattribs=self._attribs())

    def rect(self, x, y, w, h, stroke=1, fill=0):
        pts = [self._pt(x, y), self._pt(x + w, y), self._pt(x + w, y + h), self._pt(x, y + h)]
        self.msp.add_lwpolyline(pts, close=True, dxfattribs=self._attribs())

    def beginPath(self):
        return _Path()

    def drawPath(self, path, stroke=1, fill=0):
        for sub in path.subpaths:
            pts = [self._pt(x, y) for x, y in sub["pts"]]
            if len(pts) < 2:
                continue
            if fill and len(pts) >= 3:       # arrow heads
                hatch = self.msp.add_hatch(color=7)
                hatch.paths.add_polyline_path(pts, is_closed=True)
            if stroke:
                self.msp.add_lwpolyline(pts, close=sub["closed"], dxfattribs=self._attribs())

    def _text(self, x, y, text, align):
        from ezdxf.enums import TextEntityAlignment

        if not text:
            return
        a, b, _c, _d, _e, _f = self._ctm
        scale = math.hypot(a, b) or 1.0
        height = self._font[1] * MM * scale * 0.72     # em size -> cap height
        entity = self.msp.add_text(
            str(text), height=height,
            dxfattribs={"rotation": math.degrees(math.atan2(b, a)), "style": "Standard"},
        )
        entity.set_placement(self._pt(x, y), align=align)

    def drawString(self, x, y, text):
        from ezdxf.enums import TextEntityAlignment
        self._text(x, y, text, TextEntityAlignment.LEFT)

    def drawCentredString(self, x, y, text):
        from ezdxf.enums import TextEntityAlignment
        self._text(x, y, text, TextEntityAlignment.CENTER)

    def drawRightString(self, x, y, text):
        from ezdxf.enums import TextEntityAlignment
        self._text(x, y, text, TextEntityAlignment.RIGHT)


class BendTableDxfWriter(BendTablePdfWriter):
    """Same pages as the PDF, but as DXF (and DWG if ODA File Converter is installed), one file per thickness."""

    def __init__(self, oda_path=None, sheets_per_row=4, group_by_thickness=True, **kwargs):
        super().__init__(**kwargs)
        self.oda_path = oda_path
        self.sheets_per_row = sheets_per_row
        self.group_by_thickness = group_by_thickness

    @staticmethod
    def _thickness_tag(value):
        return "unknown_thickness" if value is None else f"{value:g}mm"

    def _write_group(self, items, out_dxf):
        import ezdxf

        doc = ezdxf.new("R2010", setup=True)     # setup=True brings the DASHED linetype
        doc.header["$LTSCALE"] = 0.5             # dash length of the bend lines: change if too long / short
        canvas = DxfCanvas(doc, per_row=self.sheets_per_row)
        pages = 0
        for item in items:
            pages += self._draw_item(canvas, item)
        doc.saveas(str(out_dxf))
        return pages

    def write_all_dxf(self, items, output_folder, base_name="bend_tables"):
        print("\n" + "=" * 60 + "\nBEND TABLE DRAWINGS (DXF/DWG, by thickness)\n" + "=" * 60)
        try:
            import ezdxf  # noqa: F401
        except ImportError:
            print("The 'ezdxf' package is required.  pip install ezdxf")
            return {"files": [], "error": "ezdxf missing"}

        from solidworks.thickness_dwg import ThicknessDwgBuilder

        converter = ThicknessDwgBuilder(None, oda_path=self.oda_path)
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        with_table = [i for i in items if (i.get("bend_table") or {}).get("rows")]
        groups = {}
        for item in with_table:
            key = item.get("thickness") if self.group_by_thickness else None
            groups.setdefault(key, []).append(item)

        files = []
        for key in sorted(groups, key=lambda k: (k is None, k or 0)):
            group = sorted(groups[key], key=lambda i: i["name"])
            tag = self._thickness_tag(key) if self.group_by_thickness else "all"
            out_dxf = output_folder / f"{safe_name(base_name)}_{tag}_drawing.dxf"
            try:
                before = len(self.messages)
                pages = self._write_group(group, out_dxf)
            except Exception as exc:
                print(f"{tag} -> FAILED: {type(exc).__name__}: {exc}")
                continue

            for line in self.messages[before:]:
                print(f"    {line}")

            dwg, message = converter._convert_to_dwg(out_dxf)
            final = dwg or out_dxf
            print(f"{tag} -> {len(group)} part(s), {pages} sheet(s) -> {final.name}"
                  + ("" if dwg else f"  (DXF only: {message})"))
            files.append({"thickness": key, "path": final, "parts": len(group), "sheets": pages})

        return {"files": files, "error": None}