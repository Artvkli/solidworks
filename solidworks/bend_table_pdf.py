import re
from pathlib import Path

DEFAULT_HEADER = ["Tag", "Direction", "Angle", "Inner Radius"]


class BendTablePdfWriter:
    """
    Writes the bend table of every sheet metal part to a PDF, laid out like the sample drawing:
    A4 page, "QTY=n" on the left and the bend table (Tag / Direction / Angle / Inner Radius) in the
    lower-right corner above a small title block. One PDF per part, plus (optionally) one combined PDF.

    The values come from item["bend_table"], which BendTableReader reads from SolidWorks.
    """

    FONT = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"

    def __init__(
        self, margin=22.0, row_height=17.0, font_size=10.5, max_rows_per_page=36
    ):
        self.margin = margin
        self.row_height = row_height
        self.font_size = font_size
        self.max_rows_per_page = max_rows_per_page

    # =====================================================
    # HELPERS
    # =====================================================

    @staticmethod
    def _safe(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    @staticmethod
    def _qty_text(item):
        quantity = int(item.get("quantity", 1))
        mirror = int(item.get("mirror_quantity", 0) or 0)
        if mirror:
            return f"QTY={quantity} ({quantity - mirror}+{mirror} MIRROR)"
        return f"QTY={quantity}"

    @staticmethod
    def _table_data(item):
        table = item.get("bend_table") or {}
        rows = [[str(c) for c in row] for row in table.get("rows", [])]
        header = [str(c) for c in (table.get("header") or [])]
        columns = max([len(header)] + [len(r) for r in rows]) if rows else len(header)
        if not header or all(not h for h in header):
            header = DEFAULT_HEADER[:columns] + [""] * max(
                0, columns - len(DEFAULT_HEADER)
            )
        header += [""] * (columns - len(header))
        rows = [r + [""] * (columns - len(r)) for r in rows]
        return header, rows

    def _column_widths(self, header, rows):
        from reportlab.pdfbase.pdfmetrics import stringWidth

        widths = []
        for c in range(len(header)):
            texts = [header[c]] + [r[c] for r in rows]
            longest = max(stringWidth(t, self.FONT, self.font_size) for t in texts)
            widths.append(max(55.0, longest + 26.0))
        return widths

    # =====================================================
    # ONE PAGE
    # =====================================================

    def _draw_page(self, c, item, header, rows, widths, page_no, page_count):
        page_w, page_h = c._pagesize
        m = self.margin
        block_h = 56.0  # title block height
        rh = self.row_height

        # sheet border
        c.setLineWidth(0.8)
        c.rect(m, m, page_w - 2 * m, page_h - 2 * m)

        # ---- title block (bottom) ----
        c.setLineWidth(0.6)
        c.rect(m, m, page_w - 2 * m, block_h)
        x_mid = m + (page_w - 2 * m) * 0.55
        x_right = m + (page_w - 2 * m) * 0.82
        c.line(x_mid, m, x_mid, m + block_h)
        c.line(x_right, m, x_right, m + block_h)

        def cell(x, label, value, size):
            c.setFont(self.FONT, 6)
            c.drawString(x + 5, m + block_h - 10, label)
            c.setFont(self.FONT_BOLD, size)
            c.drawString(x + 5, m + 14, value)

        name = str(item["name"])
        name_size = 14 if len(name) <= 28 else 10
        cell(m, "DWG NO. / PART", name, name_size)

        thickness = item.get("thickness")
        thickness_text = f"{thickness:g} mm" if thickness is not None else "-"
        cell(x_mid, "THICKNESS", thickness_text, 12)
        cell(x_right, "SHEET", f"{page_no} OF {page_count}", 12)

        # ---- bend table (lower right, above the title block) ----
        table_w = sum(widths)
        table_h = rh * (len(rows) + 1)
        x0 = page_w - m - table_w
        y0 = m + block_h + 14.0  # bottom of the table
        top = y0 + table_h

        c.setLineWidth(0.6)
        for i in range(len(rows) + 2):
            y = top - i * rh
            c.line(x0, y, x0 + table_w, y)

        x = x0
        for w in [0.0] + widths:
            x += w
            c.line(x, top, x, y0)

        for r, cells in enumerate([header] + rows):
            baseline = top - r * rh - rh / 2.0 - self.font_size * 0.35
            x = x0
            for col, text in enumerate(cells):
                c.setFont(self.FONT_BOLD if r == 0 else self.FONT, self.font_size)
                c.drawCentredString(x + widths[col] / 2.0, baseline, text)
                x += widths[col]

        # ---- quantity, left of the table (like the sample) ----
        c.setFont(self.FONT, 17)
        c.drawString(m + 34, y0 + rh * 1.2, self._qty_text(item))

    def _draw_item(self, c, item):
        header, rows = self._table_data(item)
        widths = self._column_widths(header, rows)

        chunks = [
            rows[i : i + self.max_rows_per_page]
            for i in range(0, len(rows), self.max_rows_per_page)
        ] or [[]]

        for number, chunk in enumerate(chunks, start=1):
            self._draw_page(c, item, header, chunk, widths, number, len(chunks))
            c.showPage()

        return len(chunks)

    # =====================================================
    # FILES
    # =====================================================

    def write_part(self, item, path):
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

        canvas_obj = canvas.Canvas(str(path), pagesize=A4)
        canvas_obj.setTitle(f"Bend table - {item['name']}")
        pages = self._draw_item(canvas_obj, item)
        canvas_obj.save()
        return pages

    def write_all(self, items, output_folder, combined_name=None):
        print("\n" + "=" * 60 + "\nBEND TABLE PDF\n" + "=" * 60)

        try:
            import reportlab  # noqa: F401
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas
        except ImportError:
            print("The 'reportlab' package is required for PDF output.")
            print("Install it with:  pip install reportlab")
            return {
                "written": 0,
                "files": [],
                "combined": None,
                "skipped": [],
                "error": "reportlab missing",
            }

        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        with_table = [i for i in items if (i.get("bend_table") or {}).get("rows")]
        skipped = [i["name"] for i in items if i not in with_table]

        files = []
        used = set()
        for item in sorted(with_table, key=lambda i: i["name"]):
            base = self._safe(item["name"])
            unique = base
            counter = 2
            while unique.lower() in used:
                unique = f"{base}_{counter}"
                counter += 1
            used.add(unique.lower())

            path = output_folder / f"{unique}.pdf"
            try:
                pages = self.write_part(item, path)
                files.append(path)
                print(
                    f"{item['name']} -> {path.name} ({len(item['bend_table']['rows'])} bends, {pages} page(s))"
                )
            except Exception as exc:
                skipped.append(item["name"])
                print(f"{item['name']} -> FAILED: {type(exc).__name__}: {exc}")

        combined = None
        if combined_name and with_table:
            combined = output_folder / f"{self._safe(combined_name)}.pdf"
            try:
                canvas_obj = canvas.Canvas(str(combined), pagesize=A4)
                canvas_obj.setTitle("Bend tables")
                for item in sorted(with_table, key=lambda i: i["name"]):
                    self._draw_item(canvas_obj, item)
                canvas_obj.save()
                print(f"\nCombined PDF: {combined}")
            except Exception as exc:
                combined = None
                print(f"Could not write the combined PDF: {type(exc).__name__}: {exc}")

        print(f"\nPDF files written: {len(files)}")
        if skipped:
            print(f"Parts without a bend table (no PDF): {len(skipped)}")
            for name in skipped:
                print(f"  - {name}")

        return {
            "written": len(files),
            "files": files,
            "combined": combined,
            "skipped": skipped,
            "error": None,
        }
