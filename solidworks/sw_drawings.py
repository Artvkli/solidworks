"""
Real SolidWorks drawings (.slddrw) of the sheet metal parts, and the PDF saved FROM SolidWorks.

STEP 1 (this file): the drawing sheet with the flat pattern view (at a standard scale), SolidWorks' own bend
table with the A, B, C... tags, the quantity note and optional custom properties; saved as .slddrw and as PDF.
Later steps add the folded views, the dimensions and the title block fields.

It reuses the helpers of the BendTableReader you already have (templates, millimetres, bend table insertion),
and the view sizes that reader saved in output/flat_view/<part>_probe.json.
"""
import json
import re
from pathlib import Path

try:
    import pythoncom
    from win32com.client import VARIANT
except ImportError:  # allows importing outside Windows
    pythoncom = None
    VARIANT = None

# standard drawing scales, as decimals (sheet size / real size), biggest first
STANDARD_SCALES = (
    5.0, 2.0, 1.0, 1 / 2.0, 1 / 2.5, 1 / 5.0, 1 / 10.0, 1 / 15.0, 1 / 20.0, 1 / 25.0, 1 / 30.0,
    1 / 40.0, 1 / 50.0, 1 / 75.0, 1 / 100.0, 1 / 150.0, 1 / 200.0, 1 / 250.0, 1 / 500.0,
)
FOLDED_BAND = 0.38            # share of the free height kept for the folded views (they come in step 2)
DIMENSION_ROOM = 0.030        # metres kept free around a view for its dimensions (they come in step 3)
TABLE_ROW_HEIGHT = 0.007      # metres per row of the bend table (estimate, used only to keep the space free)
DEFAULT_SHEET = (0.210, 0.297)  # A4 portrait, used only if the sheet size cannot be read

SAVE_SILENT = 1               # swSaveAsOptions_Silent
CUSTOM_INFO_TEXT = 30         # swCustomInfoText
REPLACE_VALUE = 2             # swCustomPropertyReplaceValue
ANCHOR_BOTTOM_RIGHT = 4       # swBOMConfigurationAnchor_BottomRight


def safe_name(name):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
    return re.sub(r"\s+", " ", cleaned) or "part"


def choose_scale(width, height, fit_width, fit_height):
    """The biggest standard scale at which a width x height view (metres) fits in fit_width x fit_height."""
    if fit_width <= 0 or fit_height <= 0 or width <= 0 or height <= 0:
        return STANDARD_SCALES[-1]
    for scale in STANDARD_SCALES:
        if width * scale <= fit_width and height * scale <= fit_height:
            return scale
    return STANDARD_SCALES[-1]


def scale_text(scale):
    return f"{scale:.3g}:1" if scale >= 1 else f"1:{1.0 / scale:.3g}"


def plan_flat_view(sheet, margins_mm, table_rows, flat_size_mm):
    """
    Where the flat pattern view goes and at which scale. All lengths in metres.

    sheet = (width, height); margins_mm = (left, right, top, bottom) where bottom includes the title block;
    flat_size_mm = (width, height) of the flat pattern at 1:1. Returns {"scale", "center", "free_area"}.
    """
    left, right, top, bottom = [m / 1000.0 for m in margins_mm]
    sheet_w, sheet_h = sheet

    table_h = (table_rows + 1) * TABLE_ROW_HEIGHT
    x0, x1 = left, sheet_w - right
    y0 = bottom + table_h + 0.010          # above the bend table
    y1 = sheet_h - top
    band = FOLDED_BAND * (y1 - y0)         # the top part is for the folded views
    flat_y1 = y1 - band

    fit_w = (x1 - x0) - 2 * DIMENSION_ROOM
    fit_h = (flat_y1 - y0) - 2 * DIMENSION_ROOM
    scale = choose_scale(flat_size_mm[0] / 1000.0, flat_size_mm[1] / 1000.0, fit_w, fit_h)

    return {
        "scale": scale,
        "center": ((x0 + x1) / 2.0, (y0 + flat_y1) / 2.0),
        "free_area": (x0, y0, x1, y1),
    }


class SwDrawingBuilder:
    def __init__(self, reader, margins_mm=(10, 10, 10, 50), properties=None, save_pdf=True):
        """
        reader      the BendTableReader already used in main.py (its templates and helpers are reused)
        margins_mm  (left, right, top, bottom) of the free area of the sheet; the bottom one must cover the
                    title block of the template
        properties  custom properties written to every drawing, e.g.
                    {"DWG NO": "{name}", "THICKNESS": "{thickness} mm", "QTY": "{quantity}"}
        """
        self.reader = reader
        self.sw_app = reader.sw_app
        self.margins_mm = tuple(margins_mm)
        self.properties = dict(properties or {})
        self.save_pdf = save_pdf

    # =====================================================
    # SMALL HELPERS
    # =====================================================

    def _get(self, obj, name, default=None):
        return self.reader._get(obj, name, default)

    def _call(self, obj, name, *args, default=None):
        return self.reader._call(obj, name, *args, default=default)

    @staticmethod
    def _qty_text(item):
        quantity = int(item.get("quantity", 1))
        mirror = int(item.get("mirror_quantity", 0) or 0)
        return f"QTY={quantity} ({quantity - mirror}+{mirror} MIRROR)" if mirror else f"QTY={quantity}"

    def _sheet_size(self, drawing, steps):
        sheet = self._call(drawing, "GetCurrentSheet")
        props = self._call(sheet, "GetProperties2")
        try:
            width, height = float(props[5]), float(props[6])
            if width > 0 and height > 0:
                steps.append(f"sheet: {width * 1000:.0f} x {height * 1000:.0f} mm")
                return width, height
        except Exception:
            pass
        steps.append(f"sheet size unknown - assuming A4 portrait {DEFAULT_SHEET[0] * 1000:.0f} x {DEFAULT_SHEET[1] * 1000:.0f} mm")
        return DEFAULT_SHEET

    @staticmethod
    def _flat_size_mm(item):
        """Width x height of the flat pattern view at 1:1, from the data the reader saved."""
        path = item.get("view_json")
        if not path or not Path(path).exists():
            return None
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            x0, y0, x1, y1 = [float(v) for v in data["outline"]][:4]
            scale = float(data.get("scale") or 1.0) or 1.0
            return (x1 - x0) * 1000.0 / scale, (y1 - y0) * 1000.0 / scale
        except Exception:
            return None

    def _set_scale(self, view, scale, steps):
        """Give the view a custom scale. Tries the decimal scale first, then the ratio."""
        try:
            view.ScaleDecimal = scale
        except Exception as exc:
            steps.append(f"view scale: ScaleDecimal could not be set ({exc})")
        current = self._get(view, "ScaleDecimal")
        if isinstance(current, (int, float)) and abs(current - scale) < scale * 0.01:
            steps.append(f"view scale set to {scale_text(scale)}")
            return True

        if VARIANT is not None:
            try:
                view.ScaleRatio = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (1.0, 1.0 / scale))
                current = self._get(view, "ScaleDecimal")
                if isinstance(current, (int, float)) and abs(current - scale) < scale * 0.01:
                    steps.append(f"view scale set to {scale_text(scale)} (ratio)")
                    return True
            except Exception as exc:
                steps.append(f"view scale: ScaleRatio could not be set ({exc})")

        steps.append(f"view scale could NOT be set (now {current}) - the view keeps the sheet scale")
        return False

    def _place_view(self, view, center, steps):
        if VARIANT is None:
            return
        try:
            view.Position = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (center[0], center[1]))
        except Exception as exc:
            steps.append(f"view position could not be set ({exc})")

    def _insert_table(self, view, x, y, steps):
        """Bend table with its bottom-right corner at (x, y). Falls back to the reader's own insertion."""
        constants = self.reader._constants()
        anchor = getattr(constants, "swBOMConfigurationAnchor_BottomRight", ANCHOR_BOTTOM_RIGHT) if constants else ANCHOR_BOTTOM_RIGHT
        template = self.reader.bend_template or ""
        try:
            table = view.InsertBendTable(False, x, y, anchor, "A", template)
            if table is not None:
                steps.append("bend table inserted at the lower right")
                return table
            steps.append("bend table: InsertBendTable returned None")
        except Exception as exc:
            steps.append(f"bend table at the lower right failed ({exc}) - trying the reader's way")
        return self.reader._insert_bend_table(view, steps)

    def _add_note(self, drawing, text, x, y, steps):
        try:
            note = drawing.InsertNote(text)
            annotation = self._call(note, "GetAnnotation")
            annotation.SetPosition2(x, y, 0.0)
            steps.append(f"note '{text}' placed")
            return True
        except Exception as exc:
            steps.append(f"note '{text}' failed: {exc}")
            return False

    def _write_properties(self, drawing, item, steps):
        if not self.properties:
            return
        values = {
            "name": item["name"],
            "thickness": f"{item['thickness']:g}" if item.get("thickness") is not None else "",
            "quantity": item.get("quantity", ""),
            "configuration": item.get("configuration") or "",
        }
        try:
            manager = self._call(self._call(drawing, "Extension"), "CustomPropertyManager", "")
            written = 0
            for field, template in self.properties.items():
                try:
                    manager.Add3(str(field), CUSTOM_INFO_TEXT, str(template).format(**values), REPLACE_VALUE)
                    written += 1
                except Exception as exc:
                    steps.append(f"property '{field}' failed: {exc}")
            steps.append(f"custom properties written: {written} of {len(self.properties)}")
        except Exception as exc:
            steps.append(f"custom properties failed: {exc}")

    def _save(self, drawing, path, steps, label):
        path = Path(path)
        try:
            if path.exists():
                path.unlink()
        except Exception as exc:
            steps.append(f"{label}: cannot replace the old file ({exc})")
        try:
            drawing.SaveAs3(str(path), 0, SAVE_SILENT)
        except Exception as exc:
            steps.append(f"{label}: SaveAs3 failed: {exc}")
            return False
        ok = path.exists() and path.stat().st_size > 0
        steps.append(f"{label}: {'saved' if ok else 'NOT written'} -> {path.name}")
        return ok

    # =====================================================
    # ONE PART
    # =====================================================

    def build_part(self, item, output_folder):
        """Returns {"slddrw": Path or None, "pdf": Path or None, "message": str, "steps": [...]}."""
        steps = []
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        base = safe_name(item["name"])
        drawing_path = output_folder / f"{base}.SLDDRW"
        pdf_path = output_folder / f"{base}.pdf"

        path = Path(item["path"]).resolve()
        config = item.get("configuration") or ""
        rows = len((item.get("bend_table") or {}).get("rows", []))

        result = {"slddrw": None, "pdf": None, "message": "", "steps": steps}
        reader = self.reader
        model, opened_by_us = reader._open_part(path)
        drawing = None
        try:
            if model is None:
                result["message"] = "part could not be opened"
                return result
            reader._resolve_templates(steps)
            if not reader.drawing_template:
                result["message"] = "no drawing template found (set DRAWING_TEMPLATE_PATH)"
                return result

            self._call(model, "ForceRebuild3", True)

            drawing = self.sw_app.NewDocument(reader.drawing_template, 0, 0.0, 0.0)
            if drawing is None:
                result["message"] = "NewDocument returned None"
                return result
            reader._use_millimeters(drawing, steps)
            sheet = self._sheet_size(drawing, steps)

            # ---- where the flat pattern goes, and at which scale ----
            flat_size = self._flat_size_mm(item)
            if flat_size is None:
                steps.append("flat pattern size unknown (no saved view data) - using 1:10 in the middle of the sheet")
                plan = {"scale": 0.1, "center": (sheet[0] / 2.0, sheet[1] / 2.0)}
            else:
                plan = plan_flat_view(sheet, self.margins_mm, rows, flat_size)
                steps.append(
                    f"flat pattern {flat_size[0]:.1f} x {flat_size[1]:.1f} mm -> scale {scale_text(plan['scale'])}, "
                    f"centre ({plan['center'][0] * 1000:.0f}, {plan['center'][1] * 1000:.0f}) mm"
                )

            cx, cy = plan["center"]
            view = reader._create_flat_view(drawing, path, config)
            if view is None:
                result["message"] = "flat pattern view was not created"
                return result
            # (the view was created by the reader's helper at its fixed spot; scale and position follow)
            self._set_scale(view, plan["scale"], steps)
            self._place_view(view, (cx, cy), steps)

            # ---- bend table (tags appear in the view by themselves), bottom right above the title block ----
            left, right, top, bottom = [m / 1000.0 for m in self.margins_mm]
            self._insert_table(view, sheet[0] - right, bottom + 0.004, steps)

            # ---- quantity (bottom left, like the PDF) and the custom properties ----
            self._add_note(drawing, self._qty_text(item), left + 0.012, bottom + 0.010, steps)
            self._write_properties(drawing, item, steps)

            self._call(drawing, "ForceRebuild3", False)

            # ---- save the drawing, then the PDF FROM SolidWorks ----
            if not self._save(drawing, drawing_path, steps, "drawing"):
                result["message"] = "the drawing could not be saved"
                return result
            result["slddrw"] = drawing_path
            if self.save_pdf and self._save(drawing, pdf_path, steps, "pdf"):
                result["pdf"] = pdf_path

            result["message"] = "OK"
            return result

        except Exception as exc:
            result["message"] = f"error: {type(exc).__name__}: {exc}"
            return result
        finally:
            if drawing is not None:
                reader._close(drawing)
            if opened_by_us and model is not None:
                reader._close(model)

    # =====================================================
    # ALL PARTS
    # =====================================================

    def build_all(self, items, output_folder, limit=None):
        targets = [i for i in items if (i.get("bend_table") or {}).get("rows")]
        if limit:
            targets = targets[: int(limit)]

        print("\n" + "=" * 60 + "\nSOLIDWORKS DRAWINGS (.slddrw + PDF from SolidWorks)\n" + "=" * 60)
        if limit:
            print(f"Test mode: only the first {len(targets)} part(s)")

        made = failed = 0
        detailed = 0
        for index, item in enumerate(targets, 1):
            result = self.build_part(item, output_folder)
            label = f"[{index}/{len(targets)}] {item['name']}"
            if result["slddrw"]:
                made += 1
                print(f"{label} -> {result['slddrw'].name}" + (f" + {result['pdf'].name}" if result["pdf"] else " (PDF not written)"))
                if not result["pdf"] and self.save_pdf:
                    for line in result["steps"][-3:]:
                        print(f"    [debug] {line}")
            else:
                failed += 1
                print(f"{label} -> FAILED: {result['message']}")
            if (not result["slddrw"] or detailed < 1) and detailed < 2:
                detailed += 1
                for line in result["steps"]:
                    print(f"    [debug] {line}")

        print(f"\nDrawings saved: {made} / {len(targets)}   (folder: {output_folder})")
        return {"made": made, "failed": failed}