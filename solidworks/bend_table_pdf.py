import re
from pathlib import Path

DEFAULT_HEADER = ["Tag", "Direction", "Angle", "Inner Radius"]


def safe_name(name):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
    return re.sub(r"\s+", " ", cleaned) or "part"


def attach_flat_dxfs(items, search_folders, export_folder, exporter=None):
    """
    Gives every part that has a bend table the DXF of its flat pattern (item["flat_dxf"]).

    Existing DXFs are reused (the ones made by an earlier full run live in
    .../_parts_dxf/<part>.dxf). Only the missing ones are exported by SolidWorks.
    """
    export_folder = Path(export_folder)
    export_folder.mkdir(parents=True, exist_ok=True)
    reused = exported = failed = 0

    for item in items:
        if not (item.get("bend_table") or {}).get("rows"):
            continue

        filename = f"{safe_name(item['name'])}.dxf"
        found = next(
            (
                Path(f) / filename
                for f in search_folders
                if (Path(f) / filename).exists()
            ),
            None,
        )

        if found is not None:
            item["flat_dxf"] = str(found)
            reused += 1
            continue

        if exporter is None:
            failed += 1
            continue

        ok, message = exporter.export_part(item, export_folder / filename)
        if ok:
            item["flat_dxf"] = str(export_folder / filename)
            exported += 1
        else:
            failed += 1
            print(
                f"  flat pattern DXF for {item['name']} could not be exported: {str(message)[:160]}"
            )

    print(
        f"Flat pattern DXFs: {reused} reused, {exported} exported now, {failed} missing"
    )
    return {"reused": reused, "exported": exported, "failed": failed}


class BendTablePdfWriter:
    """
    Writes one A4 page per part, laid out like the sample drawing:
      - the flat pattern view with its overall length / width and the positions of the bend lines
        (when item["flat_dxf"] is available),
      - "QTY=n" on the left and the bend table (Tag / Direction / Angle / Inner Radius) in the
        lower-right corner above a small title block.
    One PDF per part, plus (optionally) one combined PDF.
    """

    FONT = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"
    MIN_VIEW_HEIGHT = 260.0  # space kept free for the view on the first page

    def __init__(
        self, margin=22.0, row_height=17.0, font_size=10.5, max_rows_per_page=36
    ):
        self.margin = margin
        self.row_height = row_height
        self.font_size = font_size
        self.max_rows_per_page = max_rows_per_page
        self.messages = []

    # =====================================================
    # HELPERS
    # =====================================================

    @staticmethod
    def _safe(name):
        return safe_name(name)

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

    def _load_view(self, item):
        """The flat pattern view of this part, or None (with the reason kept in self.messages)."""
        path = item.get("flat_dxf")
        if not path:
            return None
        if not Path(path).exists():
            self.messages.append(f"{item['name']}: flat pattern DXF not found ({path})")
            return None
        from solidworks.flat_view import FlatPatternView

        view, message = FlatPatternView.load(path)
        if view is None:
            self.messages.append(f"{item['name']}: flat pattern not drawn - {message}")
            return None
        for line in view.info:
            if "could not be identified" in line:
                self.messages.append(f"{item['name']}: {line}")
        return view

    # =====================================================
    # ONE PAGE
    # =====================================================

    def _draw_page(self, c, item, header, rows, widths, page_no, page_count, view=None):
        page_w, page_h = c._pagesize
        m = self.margin
        block_h = 56.0  # title block height
        rh = self.row_height

        # sheet border
        c.setLineWidth(0.8)
        c.rect(m, m, page_w - 2 * m, page_h - 2 * m)

        # ---- bend table (lower right, above the title block) ----
        table_w = sum(widths)
        table_h = rh * (len(rows) + 1)
        x0 = page_w - m - table_w
        y0 = m + block_h + 14.0  # bottom of the table
        top = y0 + table_h

        # ---- flat pattern view (above the table), first page only ----
        scale_text = "-"
        if view is not None and page_no == 1:
            area = (m + 14.0, top + 26.0, page_w - m - 14.0, page_h - m - 14.0)
            try:
                description = view.draw(c, area)
                scale_text = view.scale_label
                self.messages.append(f"{item['name']}: view drawn - {description}")
            except Exception as exc:
                self.messages.append(
                    f"{item['name']}: view drawing failed - {type(exc).__name__}: {exc}"
                )

        # ---- title block (bottom) ----
        c.setLineWidth(0.6)
        c.setStrokeColorRGB(0, 0, 0)
        c.setFillColorRGB(0, 0, 0)
        c.rect(m, m, page_w - 2 * m, block_h)
        inner = page_w - 2 * m
        cuts = [m + inner * 0.50, m + inner * 0.68, m + inner * 0.84]
        for x in cuts:
            c.line(x, m, x, m + block_h)

        def cell(x, label, value, size):
            c.setFont(self.FONT, 6)
            c.drawString(x + 5, m + block_h - 10, label)
            c.setFont(self.FONT_BOLD, size)
            c.drawString(x + 5, m + 14, value)

        name = str(item["name"])
        cell(m, "DWG NO. / PART", name, 14 if len(name) <= 28 else 10)

        thickness = item.get("thickness")
        cell(
            cuts[0],
            "THICKNESS",
            f"{thickness:g} mm" if thickness is not None else "-",
            12,
        )
        cell(cuts[1], "SCALE", scale_text, 12)
        cell(cuts[2], "SHEET", f"{page_no} OF {page_count}", 12)

        # ---- the table itself ----
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
        view = self._load_view(item)

        # with a view, the first page keeps room for it, so it holds fewer table rows
        page_h = c._pagesize[1]
        room = page_h - 2 * self.margin - 56.0 - 14.0 - 26.0 - self.MIN_VIEW_HEIGHT
        first = (
            max(5, min(self.max_rows_per_page, int(room // self.row_height) - 1))
            if view
            else self.max_rows_per_page
        )

        chunks, index = [], 0
        while index < len(rows) or not chunks:
            size = first if not chunks else self.max_rows_per_page
            chunks.append(rows[index : index + size])
            index += size

        for number, chunk in enumerate(chunks, start=1):
            self._draw_page(c, item, header, chunk, widths, number, len(chunks), view)
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
                before = len(self.messages)
                pages = self.write_part(item, path)
                files.append(path)
                print(
                    f"{item['name']} -> {path.name} ({len(item['bend_table']['rows'])} bends, {pages} page(s))"
                )
                for line in self.messages[before:]:
                    print(f"    {line}")
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
