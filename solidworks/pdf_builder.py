import csv
import re
from datetime import date
from pathlib import Path

PAPER_SIZES_MM = {"A4": (210.0, 297.0), "A3": (297.0, 420.0)}  # (short side, long side)

# drawing scales: 5:1 ... 1:200
SCALES = [
    5,
    4,
    2,
    1,
    1 / 2,
    1 / 2.5,
    1 / 4,
    1 / 5,
    1 / 7.5,
    1 / 10,
    1 / 15,
    1 / 20,
    1 / 25,
    1 / 30,
    1 / 40,
    1 / 50,
    1 / 75,
    1 / 100,
    1 / 150,
    1 / 200,
]

FRAME_MARGIN = 10.0  # mm, border to paper edge
TITLE_BLOCK_H = 38.0  # mm
NOTES_H = 26.0  # mm, room above the title block for QTY / THK notes
DIM_MARGIN_LEFT = 28.0  # room for the vertical overall dimension
DIM_MARGIN_RIGHT = 14.0
DIM_MARGIN_TOP = 20.0  # room for the horizontal overall dimension

DASHED_KEYWORDS = ("DASH", "CENTER", "PHANTOM", "HIDDEN")
UNIT_TO_MM = {
    1: 25.4,
    2: 304.8,
    4: 1.0,
    5: 10.0,
    6: 1000.0,
}  # ezdxf units: inch, foot, mm, cm, m


class PdfDrawingExporter:
    """
    Builds one PDF drawing per sheet metal part from its flat pattern DXF:
    flat pattern view to scale, overall dimensions, QTY / thickness notes and a title block.
    Files are sorted into one folder per thickness.
    """

    def __init__(
        self,
        paper="A4",
        orientation="portrait",
        company="TECHNICAL DEPARTMENT",
        drawn_by="",
        checked_by="",
        approved_by="",
        material="",
        deburr_note="DEBURR AND BREAK SHARP EDGES",
        date_text="",
    ):
        if paper not in PAPER_SIZES_MM:
            raise ValueError(f"paper must be one of {sorted(PAPER_SIZES_MM)}")
        if orientation not in ("portrait", "landscape", "auto"):
            raise ValueError("orientation must be 'portrait', 'landscape' or 'auto'")
        self.paper = paper
        self.orientation = orientation
        self.company = company
        self.drawn_by = drawn_by
        self.checked_by = checked_by
        self.approved_by = approved_by
        self.material = material
        self.deburr_note = deburr_note
        self.date_text = date_text

    # =====================================================
    # HELPERS
    # =====================================================

    @staticmethod
    def _safe(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    @staticmethod
    def thickness_folder(thickness):
        return "unknown_thickness" if thickness is None else f"{thickness:g}mm"

    @staticmethod
    def _scale_label(scale):
        return f"{scale:g}:1" if scale >= 1 else f"1:{1 / scale:g}"

    @staticmethod
    def _is_dashed(doc, entity):
        layer_name = str(entity.dxf.get("layer", "0") or "0")
        if "bend" in layer_name.lower():
            return True

        linetype = str(entity.dxf.get("linetype", "BYLAYER") or "BYLAYER")
        if linetype.upper() == "BYLAYER":
            try:
                linetype = str(doc.layers.get(layer_name).dxf.linetype)
            except Exception:
                linetype = "CONTINUOUS"

        return any(key in linetype.upper() for key in DASHED_KEYWORDS)

    # =====================================================
    # READ THE FLAT PATTERN DXF
    # =====================================================

    def _read_geometry(self, dxf_path):
        """Returns (solid_polylines, dashed_polylines, texts, bbox) in millimetres."""
        import ezdxf
        from ezdxf import path as ezpath

        doc = ezdxf.readfile(str(dxf_path))
        factor = UNIT_TO_MM.get(int(getattr(doc, "units", 0) or 0), 1.0)

        solid, dashed, texts = [], [], []

        for entity in doc.modelspace():
            kind = entity.dxftype()

            if kind in ("TEXT", "MTEXT"):
                try:
                    text = entity.dxf.text if kind == "TEXT" else entity.text
                    insert = entity.dxf.insert
                    texts.append((insert.x * factor, insert.y * factor, str(text)))
                except Exception:
                    pass
                continue

            try:
                geometry = ezpath.make_path(entity)
            except Exception:
                continue  # entity type without outline (point, dimension, ...)

            target = dashed if self._is_dashed(doc, entity) else solid

            sub_paths = (
                list(geometry.sub_paths()) if geometry.has_sub_paths else [geometry]
            )
            for sub in sub_paths:
                points = [
                    (v.x * factor, v.y * factor)
                    for v in sub.flattening(0.05 / max(factor, 1e-9))
                ]
                if len(points) >= 2:
                    target.append(points)

        all_points = [p for line in solid + dashed for p in line]
        if not all_points:
            raise ValueError("flat pattern DXF contains no geometry")

        xs = [p[0] for p in all_points]
        ys = [p[1] for p in all_points]
        return solid, dashed, texts, (min(xs), min(ys), max(xs), max(ys))

    # =====================================================
    # LAYOUT (orientation + scale)
    # =====================================================

    def _best_layout(self, width, height):
        short, long_ = PAPER_SIZES_MM[self.paper]
        best = None

        options = {"portrait": (True,), "landscape": (False,), "auto": (True, False)}[
            self.orientation
        ]
        for portrait in options:
            paper_w, paper_h = (short, long_) if portrait else (long_, short)

            area_w = paper_w - 2 * FRAME_MARGIN - DIM_MARGIN_LEFT - DIM_MARGIN_RIGHT
            area_h = (
                paper_h - 2 * FRAME_MARGIN - TITLE_BLOCK_H - NOTES_H - DIM_MARGIN_TOP
            )

            scale = None
            for candidate in SCALES:
                if width * candidate <= area_w and height * candidate <= area_h:
                    scale = candidate
                    break
            if scale is None:
                scale = SCALES[-1]

            option = (scale, portrait, paper_w, paper_h, area_w, area_h)
            if best is None or option[0] > best[0] + 1e-12:
                best = option

        return best

    # =====================================================
    # DRAWING PRIMITIVES
    # =====================================================

    @staticmethod
    def _box(ax, x, y, w, h, lw=0.8):
        from matplotlib.patches import Rectangle

        ax.add_patch(
            Rectangle((x, y), w, h, fill=False, edgecolor="black", linewidth=lw)
        )

    @staticmethod
    def _text(
        ax,
        x,
        y,
        text,
        size=7,
        ha="left",
        va="center",
        family="serif",
        rotation=0,
        color="black",
    ):
        ax.text(
            x,
            y,
            text,
            fontsize=size,
            ha=ha,
            va=va,
            family=family,
            rotation=rotation,
            color=color,
        )

    @staticmethod
    def _dimension(ax, p1, p2, text, vertical=False):
        ax.annotate(
            "",
            xy=p2,
            xytext=p1,
            arrowprops=dict(
                arrowstyle="<|-|>",
                color="0.3",
                lw=0.6,
                shrinkA=0,
                shrinkB=0,
                mutation_scale=7,
            ),
        )
        mx, my = (p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2
        if vertical:
            ax.text(
                mx - 1.4,
                my,
                text,
                fontsize=8,
                ha="right",
                va="center",
                rotation=90,
                color="0.2",
            )
        else:
            ax.text(
                mx, my + 1.4, text, fontsize=8, ha="center", va="bottom", color="0.2"
            )

    @staticmethod
    def _line(ax, p1, p2, lw=0.4, color="0.45"):
        ax.plot(
            [p1[0], p2[0]],
            [p1[1], p2[1]],
            color=color,
            linewidth=lw,
            solid_capstyle="butt",
        )

    # =====================================================
    # TITLE BLOCK (similar to the company drawing)
    # =====================================================

    def _draw_frame_and_title_block(self, ax, paper_w, paper_h, name, scale_text):
        m = FRAME_MARGIN
        self._box(ax, m, m, paper_w - 2 * m, paper_h - 2 * m, lw=1.2)

        x0, x1 = m, paper_w - m
        y0, y1 = m, m + TITLE_BLOCK_H
        total = x1 - x0
        c1 = x0 + 0.29 * total
        c2 = c1 + 0.27 * total

        self._box(ax, x0, y0, total, TITLE_BLOCK_H, lw=1.0)

        # ---- left block ----
        self._box(ax, x0, y0 + 28, c1 - x0, 10)
        self._text(ax, (x0 + c1) / 2, y0 + 33, self.company, size=7.5, ha="center")

        half = x0 + (c1 - x0) * 0.5
        self._box(ax, x0, y0 + 15, half - x0, 13)
        self._text(ax, x0 + 1.5, y0 + 26.3, "SIGNATURE", size=4.5)
        self._box(ax, half, y0 + 15, c1 - half, 13)
        self._text(
            ax,
            half + 1.5,
            y0 + 24,
            "\n".join(self._wrap(self.deburr_note, 16)),
            size=5,
            va="center",
        )

        rows = (
            ("DRAWN BY", self.drawn_by, 10),
            ("CHECKED BY", self.checked_by, 5),
            ("APPROVED BY", self.approved_by, 0),
        )
        label_w = (c1 - x0) * 0.38
        for label, value, dy in rows:
            self._box(ax, x0, y0 + dy, label_w, 5)
            self._box(ax, x0 + label_w, y0 + dy, (c1 - x0) - label_w, 5)
            self._text(ax, x0 + 1, y0 + dy + 2.5, label, size=4.5)
            self._text(ax, x0 + label_w + 1.5, y0 + dy + 2.5, value, size=5.5)

        # ---- middle block ----
        self._box(ax, c1, y0, c2 - c1, TITLE_BLOCK_H)
        self._box(ax, c1, y0 + 6, c2 - c1, 6)
        self._text(ax, c1 + 1.5, y0 + 9, f"MATERIAL: {self.material}", size=5)
        self._text(ax, c1 + 1.5, y0 + 3, "WEIGHT:", size=5)

        # ---- right block ----
        top = y1
        self._box(ax, c2, top - 6, (x1 - c2) * 0.62, 6)
        self._box(ax, c2 + (x1 - c2) * 0.62, top - 6, (x1 - c2) * 0.38, 6)
        self._text(ax, c2 + 1.5, top - 3, "DO NOT SCALE DRAWING", size=4.5)
        self._text(ax, c2 + (x1 - c2) * 0.62 + 1.5, top - 3, "REVISION", size=4.5)
        self._box(ax, c2, top - 16, x1 - c2, 10)
        self._box(ax, c2, top - 22, x1 - c2, 6)
        self._text(ax, c2 + 1.5, top - 19, "TITLE:", size=4.5)

        dwg_h = 10
        dwg_y = top - 32
        self._box(ax, c2, dwg_y, (x1 - c2) * 0.88, dwg_h)
        self._box(ax, c2 + (x1 - c2) * 0.88, dwg_y, (x1 - c2) * 0.12, dwg_h)
        self._text(ax, c2 + 1.5, dwg_y + dwg_h - 1.8, "DWG NO.", size=4.5)
        self._text(ax, c2 + (x1 - c2) * 0.44, dwg_y + 3.8, name, size=10, ha="center")
        self._text(
            ax,
            c2 + (x1 - c2) * 0.94,
            dwg_y + dwg_h - 2.5,
            self.paper,
            size=5,
            ha="center",
        )

        self._box(ax, c2, y0, (x1 - c2) * 0.62, 6)
        self._box(ax, c2 + (x1 - c2) * 0.62, y0, (x1 - c2) * 0.38, 6)
        self._text(ax, c2 + 1.5, y0 + 3, f"SCALE: {scale_text}", size=5)
        self._text(ax, c2 + (x1 - c2) * 0.62 + 1.5, y0 + 3, "SHEET 1 OF 1", size=5)

    @staticmethod
    def _wrap(text, width):
        words, lines, current = str(text).split(), [], ""
        for word in words:
            if current and len(current) + 1 + len(word) > width:
                lines.append(current)
                current = word
            else:
                current = f"{current} {word}".strip()
        if current:
            lines.append(current)
        return lines or [""]

    # =====================================================
    # ONE PART
    # =====================================================

    def export_part(self, item, dxf_path, pdf_path):
        """Returns (ok, message)."""
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.collections import LineCollection

        try:
            solid, dashed, texts, (min_x, min_y, max_x, max_y) = self._read_geometry(
                dxf_path
            )
        except Exception as exc:
            return False, f"cannot read flat pattern DXF: {exc}"

        width = max_x - min_x
        height = max_y - min_y
        if width <= 0 or height <= 0:
            return False, "flat pattern has no size"

        scale, _, paper_w, paper_h, area_w, area_h = self._best_layout(width, height)
        scale_text = self._scale_label(scale)

        m = FRAME_MARGIN
        area_x0 = m + DIM_MARGIN_LEFT
        area_y0 = m + TITLE_BLOCK_H + NOTES_H
        offset_x = area_x0 + (area_w - width * scale) / 2
        offset_y = area_y0 + (area_h - height * scale) / 2

        def tx(point):
            return (
                offset_x + (point[0] - min_x) * scale,
                offset_y + (point[1] - min_y) * scale,
            )

        fig = plt.figure(figsize=(paper_w / 25.4, paper_h / 25.4))
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, paper_w)
        ax.set_ylim(0, paper_h)
        ax.set_aspect("equal")
        ax.axis("off")

        try:
            self._draw_frame_and_title_block(
                ax, paper_w, paper_h, str(item["name"]), scale_text
            )

            # ---- flat pattern ----
            if solid:
                ax.add_collection(
                    LineCollection(
                        [[tx(p) for p in line] for line in solid],
                        colors="black",
                        linewidths=0.9,
                        capstyle="round",
                    )
                )
            if dashed:
                ax.add_collection(
                    LineCollection(
                        [[tx(p) for p in line] for line in dashed],
                        colors="black",
                        linewidths=0.4,
                        linestyles=(0, (6, 2.5)),
                    )
                )
            for x, y, text in texts:
                px, py = tx((x, y))
                ax.text(px, py, text, fontsize=4.5, color="0.25")

            # ---- overall dimensions ----
            left, bottom = tx((min_x, min_y))
            right, top = tx((max_x, max_y))

            y_dim = top + 9
            self._line(ax, (left, top + 1), (left, y_dim + 2))
            self._line(ax, (right, top + 1), (right, y_dim + 2))
            self._dimension(ax, (left, y_dim), (right, y_dim), f"{width:.1f}")

            x_dim = left - 9
            self._line(ax, (left - 1, top), (x_dim - 2, top))
            self._line(ax, (left - 1, bottom), (x_dim - 2, bottom))
            self._dimension(
                ax, (x_dim, bottom), (x_dim, top), f"{height:.1f}", vertical=True
            )

            # ---- notes ----
            quantity = int(item.get("quantity", 1))
            mirror = int(item.get("mirror_quantity", 0) or 0)
            note_x = m + 8
            note_y = m + TITLE_BLOCK_H + NOTES_H - 6
            self._text(ax, note_x, note_y, f"QTY={quantity}", size=15)
            if mirror:
                self._text(
                    ax,
                    note_x,
                    note_y - 5.5,
                    f"({quantity - mirror} + {mirror} MIRROR)",
                    size=7.5,
                )

            thickness = item.get("thickness")
            print("=" * 50)
            print("PART:", item.get("name"))
            print("THICKNESS:", thickness)
            print("THICKNESS TYPE:", type(thickness))
            print("THICKNESS SOURCE:", item.get("thickness_source"))
            print("ITEM:", item)
            print("=" * 50)
        
            if thickness is not None:

                estimated = (
                    " (estimated)"
                    if item.get("thickness_source") == "estimated"
                    else ""
                )
                self._text(
                    ax,
                    note_x,
                    note_y - (11 if mirror else 6.5),
                    f"THK = {thickness:g} mm{estimated}",
                    size=9,
                )

            fig.savefig(
                str(pdf_path),
                format="pdf",
                metadata={
                    "Title": str(item["name"]),
                    "Creator": "SolidWorks sheet metal exporter",
                },
            )
        except Exception as exc:
            return False, f"PDF drawing failed: {exc}"
        finally:
            plt.close(fig)

        if not Path(pdf_path).exists() or Path(pdf_path).stat().st_size == 0:
            return False, "PDF file was not written"
        return True, f"scale {scale_text}"

    # =====================================================
    # ALL PARTS
    # =====================================================

    def export_all(self, exported_parts, output_folder):
        """
        exported_parts: [{"item": {...}, "dxf": Path}, ...]  (as returned by ThicknessDwgBuilder)
        Writes <output_folder>/<thickness>mm/<part>.pdf
        """
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        total = len(exported_parts)
        ok_list, failed = [], []
        used = set()

        print()
        print("=" * 60)
        print("PDF DRAWINGS (one folder per thickness)")
        print("=" * 60)

        for index, entry in enumerate(
            sorted(
                exported_parts,
                key=lambda e: (e["item"].get("thickness") or 0, e["item"]["name"]),
            ),
            start=1,
        ):
            item, dxf_path = entry["item"], entry["dxf"]
            name = item["name"]
            folder = output_folder / self.thickness_folder(item.get("thickness"))
            folder.mkdir(parents=True, exist_ok=True)

            base = self._safe(name)
            unique, counter = base, 2
            while (folder / f"{unique}.pdf").as_posix().lower() in used:
                unique = f"{base}_{counter}"
                counter += 1
            pdf_path = folder / f"{unique}.pdf"
            used.add(pdf_path.as_posix().lower())

            label = f"[{index}/{total}] {name}"
            ok, message = self.export_part(item, dxf_path, pdf_path)
            if ok:
                print(f"{label} -> {folder.name}/{pdf_path.name} ({message})")
                ok_list.append(
                    (name, item.get("thickness"), item.get("quantity"), pdf_path)
                )
            else:
                print(f"{label} -> FAILED: {message}")
                failed.append((name, message))

        report = output_folder / "pdf_report.csv"
        try:
            with open(report, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["Part", "Thickness (mm)", "Quantity", "PDF", "Status"])
                for name, thickness, quantity, path in ok_list:
                    writer.writerow(
                        [
                            name,
                            thickness,
                            quantity,
                            str(path.relative_to(output_folder)),
                            "OK",
                        ]
                    )
                for name, message in failed:
                    writer.writerow([name, "", "", "", f"FAILED: {message}"])
        except Exception as exc:
            print(f"Could not save PDF report: {exc}")
            report = None

        print()
        print(f"PDFs created: {len(ok_list)} / {total}")
        if failed:
            print("Failed:")
            for name, message in failed:
                print(f"  - {name}: {message}")

        return {
            "created": len(ok_list),
            "failed": len(failed),
            "failed_parts": failed,
            "files": [p for *_, p in ok_list],
            "report_path": report,
        }
