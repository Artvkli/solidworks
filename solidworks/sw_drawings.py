import csv
import os
import re
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

from .dwg_exporter import DwgExporter

# ---------------------------------------------------------
# SolidWorks constants
# ---------------------------------------------------------
SW_DOC_PART = 1
SW_DWG_PAPER_A4_VERTICAL = 7  # swDwgPaperA4sizeVertical
SW_IMPORT_FROM_ENTIRE_MODEL = 1  # swImportModelItemsFromEntireModel
ANNOTATION_TYPE_CANDIDATES = (
    1,
    2,
    4,
    8,
)  # swInsertAnnotation_e values tried until dimensions appear
ANCHOR_BOTTOM_RIGHT = 4  # bend table anchor corner (bottom right)

# ---------------------------------------------------------
# Layout (metres on the sheet)
# ---------------------------------------------------------
SCALE_DENOMINATORS = (
    1,
    2,
    3,
    4,
    5,
    8,
    10,
    15,
    20,
    25,
    30,
    40,
    50,
    60,
    75,
    100,
    150,
    200,
    300,
    500,
)
MARGIN_LEFT = 0.030
MARGIN_RIGHT = 0.022
MARGIN_TOP = 0.030
TITLE_BLOCK_TOP = 0.056  # the default sheet format's title block ends about here
VIEW_GAP = 0.020  # free space around each view (room for dimensions)
ROW_HEIGHT_GUESS = 0.0065  # bend table row height when it cannot be read


class SwDrawingPdfExporter(DwgExporter):
    """
    Creates a real SolidWorks drawing for every sheet metal part (default drawing template):
    flat pattern + folded views + bend table + model dimensions + QTY note,
    saves the .slddrw and exports the PDF.  Every optional step is best effort and logged.
    """

    def __init__(
        self,
        sw_app,
        assembly_path=None,
        reconnect=None,
        template=None,
        paper_size=SW_DWG_PAPER_A4_VERTICAL,
        folded_views=True,
        bend_table=True,
        dimensions=True,
        view_names=("Front", "Isometric"),
    ):
        super().__init__(sw_app, assembly_path=assembly_path, reconnect=reconnect)
        self.template = template
        self.paper_size = paper_size
        self.folded_views = folded_views
        self.view_names = tuple(view_names)
        self.bend_table = bend_table
        self.dimensions = dimensions
        self._template_path = None

    # =====================================================
    # TEMPLATE
    # =====================================================

    def find_template(self):
        if self._template_path:
            return self._template_path

        candidates = []
        if self.template:
            candidates.append(str(self.template))

        # 1) the default drawing template from the user's SolidWorks options
        #    (the only string option that ends with .drwdot is the drawing template)
        for index in range(0, 100):
            try:
                value = self.sw_app.GetUserPreferenceStringValue(index)
            except Exception:
                continue
            if value and str(value).lower().endswith(".drwdot"):
                candidates.append(str(value))

        # 2) documented helper
        try:
            value = self.sw_app.GetDocumentTemplate(3, "", self.paper_size, 0.0, 0.0)
            if value:
                candidates.append(str(value))
        except Exception:
            pass

        # 3) standard install folder
        program_data = os.environ.get("ProgramData")
        if program_data:
            base = Path(program_data) / "SolidWorks"
            if base.exists():
                candidates += [
                    str(p)
                    for p in sorted(
                        base.glob("SOLIDWORKS*/templates/*.drwdot"), reverse=True
                    )
                ]

        for candidate in candidates:
            if Path(candidate).exists():
                self._template_path = candidate
                return candidate
        return None

    def _find_bend_table_template(self):
        program_data = os.environ.get("ProgramData")
        if not program_data:
            return ""
        base = Path(program_data) / "SolidWorks"
        if base.exists():
            found = sorted(base.glob("**/*.sldbndtbt"), reverse=True)
            if found:
                return str(found[0])
        return ""

    # =====================================================
    # SMALL SOLIDWORKS HELPERS
    # =====================================================

    @staticmethod
    def _r8_array(*values):
        return VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in values)
        )

    def _sheet_size(self, drawing):
        sheet = self._call(drawing, "GetCurrentSheet")
        props = self._call(sheet, "GetProperties2")
        try:
            width, height = float(props[5]), float(props[6])
            if width > 0 and height > 0:
                return sheet, width, height
        except Exception:
            pass
        return sheet, 0.210, 0.297

    def _outline(self, view):
        values = self._call(view, "GetOutline")
        try:
            x0, y0, x1, y1 = (float(v) for v in values[:4])
            return x0, y0, x1, y1
        except Exception:
            return None

    def _position(self, view):
        values = self._call(view, "Position")
        try:
            return float(values[0]), float(values[1])
        except Exception:
            return None

    def _move_view_center(self, view, cx, cy):
        outline = self._outline(view)
        position = self._position(view)
        if outline is None or position is None:
            return False

        dx = cx - (outline[0] + outline[2]) / 2
        dy = cy - (outline[1] + outline[3]) / 2
        target = (position[0] + dx, position[1] + dy)

        for value in (self._r8_array(*target), target):
            try:
                view.Position = value
                break
            except Exception:
                continue
        else:
            return False

        after = self._outline(view)
        if after is None:
            return True
        return (
            abs((after[0] + after[2]) / 2 - cx) < 0.002
            and abs((after[1] + after[3]) / 2 - cy) < 0.002
        )

    def _create_view(self, drawing, kind, model_path, configuration, x, y):
        try:
            if kind == "flat":
                view = drawing.CreateFlatPatternViewFromModelView3(
                    str(model_path), str(configuration or ""), x, y, 0.0, False, False
                )
                if view is not None:
                    return view
                return drawing.CreateDrawViewFromModelView3(
                    str(model_path), "*Flat pattern", x, y, 0.0
                )
            return drawing.CreateDrawViewFromModelView3(
                str(model_path), f"*{kind}", x, y, 0.0
            )
        except Exception:
            if kind == "flat":
                try:
                    return drawing.CreateDrawViewFromModelView3(
                        str(model_path), "*Flat pattern", x, y, 0.0
                    )
                except Exception:
                    return None
            return None

    # =====================================================
    # LAYOUT
    # =====================================================

    @staticmethod
    def pack_views(sizes, region, denominator, gap=VIEW_GAP):
        """
        sizes: {key: (width, height)} at 1:1.  Returns {key: (cx, cy)} at 1:denominator or None.
        Views keep the order of `sizes` (flat pattern first, then the others); rows are filled
        left to right and every row and the whole block are centred in the region.
        """
        x0, y0, x1, y1 = region
        avail_w, avail_h = x1 - x0, y1 - y0

        items = [
            (key, w / denominator, h / denominator) for key, (w, h) in sizes.items()
        ]
        if any(w > avail_w or h > avail_h for _, w, h in items):
            return None

        rows, current, current_w = [], [], 0.0
        for item in items:
            needed = item[1] + (gap if current else 0.0)
            if current and current_w + needed > avail_w:
                rows.append(current)
                current, current_w = [item], item[1]
            else:
                current.append(item)
                current_w += needed
        if current:
            rows.append(current)

        row_heights = [max(h for _, _, h in row) for row in rows]
        total_h = sum(row_heights) + gap * (len(rows) - 1)
        if total_h > avail_h:
            return None

        positions = {}
        y_top = y1 - (avail_h - total_h) / 2
        for row, row_h in zip(rows, row_heights):
            row_w = sum(w for _, w, _ in row) + gap * (len(row) - 1)
            x = x0 + (avail_w - row_w) / 2
            for key, w, h in row:
                positions[key] = (x + w / 2, y_top - row_h / 2)
                x += w + gap
            y_top -= row_h + gap
        return positions

    def choose_scale(self, sizes, region):
        for denominator in SCALE_DENOMINATORS:
            positions = self.pack_views(sizes, region, denominator)
            if positions is not None:
                return denominator, positions
        denominator = SCALE_DENOMINATORS[-1]
        return denominator, self.pack_views(
            sizes, (region[0], region[1], region[2], region[3] + 10), denominator
        )

    # =====================================================
    # BEND TABLE / DIMENSIONS / NOTE
    # =====================================================

    @staticmethod
    def _method_signature(obj, method_name):
        """(parameter names, parameter variant types) of a COM method, read from its type library."""
        try:
            typeinfo = obj._oleobj_.GetTypeInfo()
            attr = typeinfo.GetTypeAttr()
            for index in range(attr.cFuncs):
                desc = typeinfo.GetFuncDesc(index)
                names = typeinfo.GetNames(desc.memid)
                if names and names[0].lower() == method_name.lower():
                    args = getattr(desc, "args", None) or desc[2]
                    types = []
                    for arg in args:
                        try:
                            types.append(int(arg[0][0]))
                        except Exception:
                            types.append(None)
                    return list(names[1:]), types
        except Exception:
            pass
        return None, None

    @staticmethod
    def _guess_arguments(names, types, x, y, template):
        """One value per parameter, chosen from its name (and variant type when the name says nothing)."""
        values = []
        for position, name in enumerate(names):
            key = str(name).lower()
            vt = types[position] if position < len(types) else None
            if "anchorpoint" in key or "useanchor" in key:
                values.append(False)
            elif key == "x":
                values.append(x)
            elif key == "y":
                values.append(y)
            elif "anchortype" in key or key == "anchor":
                values.append(ANCHOR_BOTTOM_RIGHT)
            elif "template" in key:
                values.append(template)
            elif vt == 11:  # VT_BOOL
                values.append(False)
            elif vt == 5:  # VT_R8
                values.append(0.0)
            elif vt == 8:  # VT_BSTR
                values.append("")
            else:
                values.append(0)
        return values

    def _insert_bend_table(self, flat_view, sheet_w, steps):
        x = sheet_w - 0.012
        y = TITLE_BLOCK_TOP + 0.004
        template = self._find_bend_table_template()

        attempts = []

        names, types = self._method_signature(flat_view, "InsertBendTable")
        if names:
            arguments = self._guess_arguments(names, types, x, y, template)
            steps.append(
                f"bend table signature: InsertBendTable({', '.join(map(str, names))})"
            )
            attempts.append(lambda: flat_view.InsertBendTable(*arguments))
        else:
            steps.append("bend table signature: could not be read from SolidWorks")

        attempts += [
            lambda: flat_view.InsertBendTable(
                False, x, y, ANCHOR_BOTTOM_RIGHT, template
            ),
            lambda: flat_view.InsertBendTable(False, x, y, ANCHOR_BOTTOM_RIGHT, ""),
            lambda: flat_view.InsertBendTable(False, ANCHOR_BOTTOM_RIGHT, template),
            lambda: flat_view.InsertBendTable(False, ANCHOR_BOTTOM_RIGHT, ""),
        ]
        errors = []
        for number, attempt in enumerate(attempts, start=1):
            try:
                table = attempt()
                if table is not None:
                    steps.append(f"bend table: inserted (variant {number})")
                    return table
                errors.append(f"v{number}: returned nothing")
            except Exception as exc:
                errors.append(f"v{number}: {str(exc)[:70]}")
        steps.append("bend table: FAILED (" + " | ".join(errors) + ")")
        return None

    def _bend_table_height(self, table):
        rows = self._get(table, "RowCount")
        try:
            rows = int(rows)
        except Exception:
            return None
        total = 0.0
        for row in range(rows):
            height = self._call(table, "GetRowHeight", row)
            try:
                total += float(height)
            except Exception:
                total += ROW_HEIGHT_GUESS
        return total if total > 0 else rows * ROW_HEIGHT_GUESS

    def _insert_dimensions(self, drawing, steps):
        errors = []
        for types in ANNOTATION_TYPE_CANDIDATES:
            try:
                result = drawing.InsertModelAnnotations3(
                    SW_IMPORT_FROM_ENTIRE_MODEL, types, True, False, False, False
                )
                count = len(result) if result else 0
                if count:
                    steps.append(f"dimensions: {count} model items (type {types})")
                    return count
                errors.append(f"type {types}: none")
            except Exception as exc:
                errors.append(f"type {types}: {str(exc)[:60]}")
        steps.append("dimensions: none inserted (" + " | ".join(errors) + ")")
        return 0

    def _insert_note(self, drawing, text, x, y, steps):
        try:
            note = drawing.InsertNote(text)
            if note is None:
                steps.append("note: InsertNote returned nothing")
                return False

            placed = False
            try:
                note.SetTextPoint(x, y, 0.0)
                placed = True
            except Exception:
                pass
            if not placed:
                try:
                    annotation = note.GetAnnotation()
                    annotation.SetPosition2(x, y, 0.0)
                    placed = True
                except Exception:
                    pass

            steps.append(
                "note: QTY inserted" + ("" if placed else " (position not set)")
            )
            return True
        except Exception as exc:
            steps.append(f"note: FAILED ({str(exc)[:70]})")
            return False

    # =====================================================
    # ONE PART
    # =====================================================

    def _build_drawing(self, item, slddrw_path, pdf_path):
        steps = []
        path = Path(item["path"]).resolve()
        configuration = item.get("configuration")
        model = None
        opened_by_us = False
        previous_config = None
        drawing = None
        drawing_title = None

        try:
            template = self.find_template()
            if not template:
                return (
                    False,
                    "no drawing template found (set SW_DRAWING_TEMPLATE in main.py)",
                    steps,
                )
            steps.append(f"template: {Path(template).name}")

            model, opened_by_us, open_status = self._open_part(path)
            if model is None:
                return False, f"cannot open part: {open_status}", steps

            _, previous_config = self._activate_configuration(model, configuration)

            drawing = self.sw_app.NewDocument(template, self.paper_size, 0.0, 0.0)
            if drawing is None:
                return False, "NewDocument returned nothing", steps
            drawing_title = self._get(drawing, "GetTitle")

            sheet, sheet_w, sheet_h = self._sheet_size(drawing)
            steps.append(f"sheet: {sheet_w * 1000:.0f} x {sheet_h * 1000:.0f} mm")
            self._call(sheet, "SetScale", 1.0, 1.0, True, True)

            # ---- create the views at 1:1 (provisional spot) to learn their sizes ----
            kinds = ["flat"] + (list(self.view_names) if self.folded_views else [])
            views, sizes = {}, {}
            for kind in kinds:
                view = self._create_view(drawing, kind, path, configuration, 0.1, 0.15)
                outline = self._outline(view) if view is not None else None
                if outline is None:
                    steps.append(f"view {kind}: FAILED")
                    continue
                views[kind] = view
                sizes[kind] = (outline[2] - outline[0], outline[3] - outline[1])
            steps.append("views: " + ", ".join(views) if views else "views: none")

            if "flat" not in views:
                return False, "flat pattern view could not be created", steps

            # ---- bend table (also tells us how much room the bottom needs) ----
            table_height = 0.0
            if self.bend_table:
                table = self._insert_bend_table(views["flat"], sheet_w, steps)
                if table is not None:
                    table_height = self._bend_table_height(table) or 0.05
            bottom = TITLE_BLOCK_TOP + max(table_height + 0.010, 0.030) + 0.006
            region = (MARGIN_LEFT, bottom, sheet_w - MARGIN_RIGHT, sheet_h - MARGIN_TOP)

            # ---- scale + positions ----
            denominator, positions = self.choose_scale(sizes, region)
            if positions is None:
                return False, "views do not fit on the sheet", steps
            steps.append(f"scale: 1:{denominator}")
            self._call(sheet, "SetScale", 1.0, float(denominator), True, True)

            # sizes changed with the scale: place by centre
            for kind, (cx, cy) in positions.items():
                if not self._move_view_center(views[kind], cx, cy):
                    return False, f"could not move view {kind}", steps

            # ---- dimensions from the model ----
            if self.dimensions:
                self._insert_dimensions(drawing, steps)

            # ---- QTY note ----
            quantity = int(item.get("quantity", 1))
            mirror = int(item.get("mirror_quantity", 0) or 0)
            text = f"QTY={quantity}" + (
                f" ({quantity - mirror} + {mirror} MIRROR)" if mirror else ""
            )
            self._insert_note(
                drawing, text, MARGIN_LEFT + 0.004, TITLE_BLOCK_TOP + 0.012, steps
            )

            # ---- save the drawing, then the PDF ----
            slddrw_path.parent.mkdir(parents=True, exist_ok=True)
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            for old in (slddrw_path, pdf_path):
                if old.exists():
                    try:
                        old.unlink()
                    except Exception:
                        pass

            try:
                drawing.SaveAs3(str(slddrw_path), 0, 1)
            except Exception as exc:
                steps.append(f"save .slddrw: FAILED ({str(exc)[:60]})")
            else:
                steps.append(
                    "saved .slddrw"
                    if self._written(slddrw_path)
                    else "save .slddrw: no file"
                )

            saved_pdf = False
            for attempt in (
                lambda: drawing.SaveAs3(str(pdf_path), 0, 1 | 2),
                lambda: drawing.Extension.SaveAs3(
                    str(pdf_path),
                    0,
                    1 | 2,
                    VARIANT(pythoncom.VT_DISPATCH, None),
                    VARIANT(pythoncom.VT_DISPATCH, None),
                    self._byref_i4(),
                    self._byref_i4(),
                ),
            ):
                try:
                    attempt()
                except Exception as exc:
                    steps.append(f"pdf save attempt: {str(exc)[:60]}")
                if self._written(pdf_path):
                    saved_pdf = True
                    break

            if not saved_pdf:
                return False, "PDF was not written", steps
            return True, "SolidWorks drawing", steps

        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}", steps
        finally:
            try:
                if drawing_title:
                    self.sw_app.CloseDoc(str(drawing_title))
            except Exception:
                pass
            if model is not None:
                self._restore_configuration(model, previous_config)
            self._close_part(model, opened_by_us)

    def export_part(self, item, slddrw_path, pdf_path):
        ok, message, steps = self._build_drawing(
            item, Path(slddrw_path), Path(pdf_path)
        )

        if not ok and self._is_connection_error(
            RuntimeError(message + " " + " ".join(steps))
        ):
            print("  SolidWorks connection problem. Reconnecting and retrying once...")
            time.sleep(3)
            if self._reconnect():
                ok, message, steps = self._build_drawing(
                    item, Path(slddrw_path), Path(pdf_path)
                )

        return ok, message, steps

    # =====================================================
    # ALL PARTS
    # =====================================================

    @staticmethod
    def _safe_name(name):
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().rstrip(".")
        return re.sub(r"\s+", " ", cleaned) or "part"

    def export_all(self, exported_parts, output_folder, fallback=None, limit=None):
        """
        exported_parts: [{"item": {...}, "dxf": Path}] as returned by ThicknessDwgBuilder.
        PDFs go to <output_folder>/<thickness>mm/<part>.pdf, editable drawings to
        <output_folder>/_drawings/<thickness>mm/<part>.slddrw.
        fallback: a PdfDrawingExporter used when the SolidWorks drawing fails.
        """
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)

        entries = sorted(
            exported_parts,
            key=lambda e: (e["item"].get("thickness") or 0, e["item"]["name"]),
        )
        if limit:
            entries = entries[: int(limit)]
        total = len(entries)

        print()
        print("=" * 60)
        print("PDF DRAWINGS (SolidWorks drawings, one folder per thickness)")
        print("=" * 60)
        template = self.find_template()
        print(f"Drawing template: {template or 'NOT FOUND'}")
        if limit:
            print(f"TEST MODE: only the first {total} parts")

        rows, used = [], set()
        sw_count = fallback_count = 0
        failed = []

        for index, entry in enumerate(entries, start=1):
            item, dxf_path = entry["item"], entry["dxf"]
            name = item["name"]
            thickness = item.get("thickness")
            folder_name = (
                "unknown_thickness" if thickness is None else f"{thickness:g}mm"
            )

            base = self._safe_name(name)
            unique, counter = base, 2
            while f"{folder_name}/{unique}".lower() in used:
                unique, counter = f"{base}_{counter}", counter + 1
            used.add(f"{folder_name}/{unique}".lower())

            pdf_path = output_folder / folder_name / f"{unique}.pdf"
            slddrw_path = output_folder / "_drawings" / folder_name / f"{unique}.slddrw"

            pdf_path.parent.mkdir(
                parents=True, exist_ok=True
            )  # the fallback needs the folder even if SolidWorks failed early

            label = f"[{index}/{total}] {name}"
            started = time.time()
            ok, message, steps = self.export_part(item, slddrw_path, pdf_path)

            if ok:
                sw_count += 1
                print(
                    f"{label} -> {folder_name}/{pdf_path.name} ({time.time() - started:.0f}s)"
                )
                for line in steps:
                    print(f"      {line}")
                rows.append(
                    [
                        name,
                        thickness,
                        item.get("quantity"),
                        f"{folder_name}/{pdf_path.name}",
                        "SolidWorks",
                        " | ".join(steps),
                    ]
                )
                continue

            print(f"{label} -> SolidWorks drawing FAILED: {message}")
            for line in steps:
                print(f"      {line}")

            if fallback is not None:
                ok2, message2 = fallback.export_part(item, dxf_path, pdf_path)
                if ok2:
                    fallback_count += 1
                    print(f"      fallback PDF created ({message2})")
                    rows.append(
                        [
                            name,
                            thickness,
                            item.get("quantity"),
                            f"{folder_name}/{pdf_path.name}",
                            "fallback",
                            f"{message} | " + " | ".join(steps),
                        ]
                    )
                    continue
                message = f"{message}; fallback: {message2}"

            failed.append((name, message))
            rows.append([name, thickness, item.get("quantity"), "", "FAILED", message])

        report = output_folder / "pdf_report.csv"
        try:
            with open(report, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(
                    ["Part", "Thickness (mm)", "Quantity", "PDF", "Method", "Details"]
                )
                writer.writerows(rows)
        except Exception as exc:
            print(f"Could not save PDF report: {exc}")
            report = None

        print()
        print(
            f"PDFs: {sw_count} SolidWorks drawings, {fallback_count} fallback, {len(failed)} failed (of {total})"
        )

        return {
            "created": sw_count + fallback_count,
            "solidworks": sw_count,
            "fallback": fallback_count,
            "failed": len(failed),
            "failed_parts": failed,
            "report_path": report,
        }
