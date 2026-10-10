import math
from pathlib import Path

PT_PER_MM = 72.0 / 25.4  # 1:1 on paper
BEND_LAYER_WORDS = ("BEND",)
DASHED_WORDS = ("DASH", "CENTER", "PHANTOM", "HIDDEN", "DOT")
AXIS_TOLERANCE = 0.01  # a bend line is "vertical/horizontal" if it leans less than 1 %
POSITION_TOLERANCE = 0.05  # mm: bend lines closer than this share one dimension


def _round_pos(value):
    return round(value / POSITION_TOLERANCE) * POSITION_TOLERANCE


class FlatPatternView:
    """
    The flat pattern of one part (read from the DXF that SolidWorks exported), ready to be drawn on a
    PDF page with its overall length / width and the positions of the bend lines.
    """

    def __init__(self, paths, info=None):
        # paths: list of {"points": [(x, y), ...], "bend": bool}
        self.paths = paths
        self.info = info or []
        self.scale_label = ""

    # =====================================================
    # LOADING
    # =====================================================

    @staticmethod
    def _is_bend(entity, doc):
        layer = str(entity.dxf.layer or "").upper()
        if any(word in layer for word in BEND_LAYER_WORDS):
            return True

        linetype = str(entity.dxf.get("linetype", "BYLAYER") or "BYLAYER").upper()
        if linetype in ("BYLAYER", "BYBLOCK", ""):
            try:
                linetype = str(doc.layers.get(entity.dxf.layer).dxf.linetype).upper()
            except Exception:
                linetype = "CONTINUOUS"
        return any(word in linetype for word in DASHED_WORDS)

    @classmethod
    def load(cls, dxf_path):
        """Returns (view_or_None, message)."""
        try:
            import ezdxf
            from ezdxf import path as ezpath
        except ImportError:
            return None, "ezdxf is not installed (pip install ezdxf)"

        try:
            doc = ezdxf.readfile(str(dxf_path))
        except Exception as exc:
            return None, f"cannot read {Path(dxf_path).name}: {exc}"

        paths = []
        layers = {}
        for entity in doc.modelspace():
            kind = entity.dxftype()
            if kind in ("TEXT", "MTEXT", "DIMENSION", "POINT", "INSERT", "HATCH"):
                continue
            try:
                geometry = ezpath.make_path(entity)
            except Exception:
                continue

            is_bend = cls._is_bend(entity, doc)
            layers[str(entity.dxf.layer)] = layers.get(str(entity.dxf.layer), 0) + 1

            for sub in geometry.sub_paths():
                points = [(v.x, v.y) for v in sub.flattening(0.05)]
                if len(points) >= 2:
                    paths.append({"points": points, "bend": is_bend})

        if not paths:
            return None, "the DXF has no drawable geometry"

        info = [f"layers: {layers}"]
        view = cls(paths, info)
        if not view.bend_lines():
            info.append(
                "bend lines could not be identified (no layer containing 'BEND' and no dashed linetype)"
            )
        return view, "OK"

    # =====================================================
    # GEOMETRY
    # =====================================================

    def _outline_points(self):
        pts = [p for path in self.paths if not path["bend"] for p in path["points"]]
        return pts or [p for path in self.paths for p in path["points"]]

    def bbox(self):
        pts = self._outline_points()
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        return min(xs), min(ys), max(xs), max(ys)

    def bend_lines(self):
        """Straight bend lines as ((x1, y1), (x2, y2))."""
        lines = []
        for path in self.paths:
            if path["bend"]:
                lines.append((path["points"][0], path["points"][-1]))
        return lines

    def rotated(self):
        """The same view turned 90 degrees (x, y) -> (-y, x)."""
        paths = [
            {"points": [(-y, x) for x, y in p["points"]], "bend": p["bend"]}
            for p in self.paths
        ]
        return FlatPatternView(paths, list(self.info))

    def bend_positions(self):
        """(x positions of vertical bend lines, y positions of horizontal bend lines, number of slanted ones)."""
        xs, ys, slanted = set(), set(), 0
        for (x1, y1), (x2, y2) in self.bend_lines():
            dx, dy = abs(x2 - x1), abs(y2 - y1)
            length = math.hypot(dx, dy)
            if length < 1e-6:
                continue
            if dx <= AXIS_TOLERANCE * length:
                xs.add(_round_pos((x1 + x2) / 2.0))
            elif dy <= AXIS_TOLERANCE * length:
                ys.add(_round_pos((y1 + y2) / 2.0))
            else:
                slanted += 1
        return sorted(xs), sorted(ys), slanted

    # =====================================================
    # DRAWING
    # =====================================================

    @staticmethod
    def _arrow(c, x, y, ux, uy, size=4.0):
        """Filled arrow head with its tip at (x, y), pointing along (ux, uy)."""
        px, py = -uy, ux
        p = c.beginPath()
        p.moveTo(x, y)
        p.lineTo(
            x - ux * size * 2.2 + px * size * 0.55,
            y - uy * size * 2.2 + py * size * 0.55,
        )
        p.lineTo(
            x - ux * size * 2.2 - px * size * 0.55,
            y - uy * size * 2.2 - py * size * 0.55,
        )
        p.close()
        c.drawPath(p, stroke=0, fill=1)

    def _hdim(self, c, xa, xb, y_dim, y_from, text, font, size):
        """Horizontal dimension between page x positions xa < xb, dimension line at y_dim."""
        sign = 1.0 if y_dim >= y_from else -1.0
        for x in (xa, xb):
            c.line(x, y_from + sign * 3.0, x, y_dim + sign * 3.0)
        c.line(xa, y_dim, xb, y_dim)
        if xb - xa > 9.0:
            self._arrow(c, xa, y_dim, 1, 0)
            self._arrow(c, xb, y_dim, -1, 0)
        c.setFont(font, size)
        c.drawCentredString(
            (xa + xb) / 2.0, y_dim + (2.5 if sign > 0 else -size - 1.5), text
        )

    def _vdim(self, c, ya, yb, x_dim, x_from, text, font, size):
        """Vertical dimension between page y positions ya < yb, dimension line at x_dim."""
        sign = 1.0 if x_dim >= x_from else -1.0
        for y in (ya, yb):
            c.line(x_from + sign * 3.0, y, x_dim + sign * 3.0, y)
        c.line(x_dim, ya, x_dim, yb)
        if yb - ya > 9.0:
            self._arrow(c, x_dim, ya, 0, 1)
            self._arrow(c, x_dim, yb, 0, -1)
        c.saveState()
        c.setFont(font, size)
        c.translate(x_dim + (-2.5 if sign > 0 else size + 1.5), (ya + yb) / 2.0)
        c.rotate(90)
        c.drawCentredString(0, 0, text)
        c.restoreState()

    @staticmethod
    def _chain_marks(low, high, positions):
        inside = [p for p in positions if low + 1e-3 < p < high - 1e-3]
        return [low] + inside + [high] if positions else None

    @staticmethod
    def _levels(marks, k, font, size):
        """
        Row (0, 1, 2, ...) for each dimension of a chain, so that neighbouring numbers never overlap.
        A number goes to the first row where it does not touch the previous number of that row.
        """
        from reportlab.pdfbase.pdfmetrics import stringWidth

        levels, last_end = [], []
        for a, b in zip(marks, marks[1:]):
            text_width = stringWidth(f"{b - a:.1f}", font, size) + 3.0
            centre = (a + b) / 2.0 * k
            start, end = centre - text_width / 2.0, centre + text_width / 2.0
            level = 0
            while True:
                if level >= len(last_end):
                    last_end.append(-1e12)
                if start >= last_end[level]:
                    last_end[level] = end
                    break
                level += 1
            levels.append(level)
        return levels

    def _plan(self, area_size, font, size):
        """Scale and margins for this orientation. The chain margins grow with the rows they need."""
        aw, ah = area_size
        xs, ys, _ = self.bend_positions()
        x0, y0, x1, y1 = self.bbox()
        width, height = max(x1 - x0, 1e-6), max(y1 - y0, 1e-6)
        xmarks = self._chain_marks(x0, x1, xs)
        ymarks = self._chain_marks(y0, y1, ys)
        top, right = 30.0, 30.0
        left = 44.0 if ymarks else 14.0
        bottom = 44.0 if xmarks else 14.0
        k = 1.0

        for _ in range(4):
            k = min(
                (aw - left - right) / width,
                (ah - top - bottom) / height,
                5.0 * PT_PER_MM,
            )
            k = max(k, 1e-6)
            new_bottom = (
                22.0 + 13.0 * (max(self._levels(xmarks, k, font, size)) + 1)
                if xmarks
                else 14.0
            )
            new_left = (
                22.0 + 13.0 * (max(self._levels(ymarks, k, font, size)) + 1)
                if ymarks
                else 14.0
            )
            if abs(new_bottom - bottom) < 0.5 and abs(new_left - left) < 0.5:
                break
            bottom, left = new_bottom, new_left
        return k, (left, bottom, top, right)

    def draw(self, c, area, font="Helvetica", font_size=8.5):
        """
        Draws the view (with dimensions) centred in area = (x0, y0, x1, y1) in PDF points.
        Returns a short description of what was drawn.
        """
        ax0, ay0, ax1, ay1 = area
        aw, ah = ax1 - ax0, ay1 - ay0

        view = self
        k, margins = self._plan((aw, ah), font, font_size)
        turned = self.rotated()
        k_turned, margins_turned = turned._plan((aw, ah), font, font_size)
        if k_turned > k * 1.2:
            view, k, margins = turned, k_turned, margins_turned

        left, bottom, top, right = margins
        x0, y0, x1, y1 = view.bbox()
        width, height = x1 - x0, y1 - y0
        drawn_w, drawn_h = width * k, height * k

        # centre the drawing (without its dimension margins) inside the free area
        ox = ax0 + left + (aw - left - right - drawn_w) / 2.0
        oy = ay0 + bottom + (ah - top - bottom - drawn_h) / 2.0

        def px(x):
            return ox + (x - x0) * k

        def py(y):
            return oy + (y - y0) * k

        # ---- geometry ----
        c.saveState()
        c.setStrokeColorRGB(0, 0, 0)
        for path in view.paths:
            pts = path["points"]
            c.setLineWidth(0.35 if path["bend"] else 0.7)
            c.setDash(3, 2) if path["bend"] else c.setDash()
            p = c.beginPath()
            p.moveTo(px(pts[0][0]), py(pts[0][1]))
            for x, y in pts[1:]:
                p.lineTo(px(x), py(y))
            c.drawPath(p, stroke=1, fill=0)
        c.setDash()

        # ---- dimensions ----
        gray = 0.25
        c.setStrokeColorRGB(gray, gray, gray)
        c.setFillColorRGB(gray, gray, gray)
        c.setLineWidth(0.4)

        left_x, right_x = px(x0), px(x1)
        bottom_y, top_y = py(y0), py(y1)

        self._hdim(
            c, left_x, right_x, top_y + 22.0, top_y, f"{width:.1f}", font, font_size
        )
        self._vdim(
            c,
            bottom_y,
            top_y,
            right_x + 22.0,
            right_x,
            f"{height:.1f}",
            font,
            font_size,
        )

        xs, ys, slanted = view.bend_positions()
        xmarks = self._chain_marks(x0, x1, xs)
        ymarks = self._chain_marks(y0, y1, ys)

        if xmarks:
            levels = self._levels(xmarks, k, font, font_size)
            for i in range(len(xmarks) - 1):
                self._hdim(
                    c,
                    px(xmarks[i]),
                    px(xmarks[i + 1]),
                    bottom_y - (16.0 + 13.0 * levels[i]),
                    bottom_y,
                    f"{xmarks[i + 1] - xmarks[i]:.1f}",
                    font,
                    font_size,
                )
        if ymarks:
            levels = self._levels(ymarks, k, font, font_size)
            for i in range(len(ymarks) - 1):
                self._vdim(
                    c,
                    py(ymarks[i]),
                    py(ymarks[i + 1]),
                    left_x - (16.0 + 13.0 * levels[i]),
                    left_x,
                    f"{ymarks[i + 1] - ymarks[i]:.1f}",
                    font,
                    font_size,
                )
        c.restoreState()

        ratio = k / PT_PER_MM
        self.scale_label = f"{ratio:.3g}:1" if ratio >= 1 else f"1:{1.0 / ratio:.3g}"
        return (
            f"{width:.1f} x {height:.1f} mm, scale {self.scale_label}, "
            f"{len(xs)} vertical / {len(ys)} horizontal / {slanted} slanted bend line(s)"
            + (", turned 90 degrees" if view is not self else "")
        )
