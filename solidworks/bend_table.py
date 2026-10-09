import csv
import gc
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

SW_DOC_PART = 1
SW_OPEN_SILENT = 1
SW_OPEN_READONLY = 2
ANCHOR_TOP_LEFT = 1  # swBOMConfigurationAnchor_TopLeft
MAX_PREFERENCE_INDEX = 80


class BendTableReader:
    """
    Reads the bend table (Tag / Direction / Angle / Inner Radius) of every sheet metal part.

    The values are NOT calculated here. SolidWorks itself builds them: for each part a
    temporary drawing is created, the flat pattern view and its bend table are inserted,
    the table cells are read, and the drawing is closed WITHOUT saving. So tags, UP/DOWN
    directions, angles and radii are exactly what SolidWorks shows (same as the sample PDF).
    """

    def __init__(
        self, sw_app, assembly_path=None, drawing_template=None, bend_template=None
    ):
        self.sw_app = sw_app
        self.assembly_path = Path(assembly_path).resolve() if assembly_path else None
        self.drawing_template = drawing_template
        self.bend_template = bend_template
        self._templates_resolved = False

    # =====================================================
    # SAFE HELPERS
    # =====================================================

    @staticmethod
    def _get(obj, name, default=None):
        if obj is None:
            return default
        try:
            value = getattr(obj, name)
            return value() if callable(value) else value
        except Exception:
            return default

    @staticmethod
    def _call(obj, name, *args, default=None):
        if obj is None:
            return default
        try:
            attr = getattr(obj, name)
            if args:
                return attr(*args) if callable(attr) else default
            if hasattr(attr, "_oleobj_"):
                return attr
            result = attr() if callable(attr) else attr
            return default if result is None else result
        except Exception:
            return default

    @staticmethod
    def _byref_i4():
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    # =====================================================
    # TEMPLATES
    # =====================================================

    def _resolve_templates(self, steps):
        if self._templates_resolved:
            return
        self._templates_resolved = True

        install = None
        exe = self._get(self.sw_app, "GetExecutablePath")
        if exe:
            install = Path(str(exe))
            if install.suffix:
                install = install.parent

        if not self.drawing_template:
            # the index of "default drawing template" differs between versions, so look at the
            # values and take the one that is a .drwdot file
            for index in range(MAX_PREFERENCE_INDEX):
                try:
                    value = self.sw_app.GetUserPreferenceStringValue(index)
                except Exception:
                    continue
                if (
                    isinstance(value, str)
                    and value.lower().endswith(".drwdot")
                    and Path(value).exists()
                ):
                    self.drawing_template = value
                    break

        if not self.drawing_template and install:
            found = sorted(install.glob("lang/*/*.drwdot"))
            if found:
                self.drawing_template = str(found[0])

        if not self.bend_template and install:
            found = sorted(install.glob("lang/english/*.sldbndtbt")) or sorted(
                install.glob("lang/*/*.sldbndtbt")
            )
            if found:
                self.bend_template = str(found[0])

        steps.append(f"drawing template: {self.drawing_template}")
        steps.append(
            f"bend table template: {self.bend_template or '(SolidWorks default)'}"
        )

    # =====================================================
    # PART OPEN / CLOSE
    # =====================================================

    def _open_part(self, path):
        try:
            existing = self.sw_app.GetOpenDocumentByName(str(path))
            if existing is not None:
                return existing, False
        except Exception:
            pass

        try:
            model = self.sw_app.OpenDoc6(
                str(path),
                SW_DOC_PART,
                SW_OPEN_SILENT | SW_OPEN_READONLY,
                "",
                self._byref_i4(),
                self._byref_i4(),
            )
            return model, model is not None
        except Exception:
            return None, False

    def _close(self, model):
        title = self._get(model, "GetTitle")
        try:
            if title:
                self.sw_app.CloseDoc(str(title))
        except Exception:
            pass

    # =====================================================
    # READ THE TABLE
    # =====================================================

    def _cell_text(self, table, row, col):
        for name in ("DisplayedText", "Text", "GetCellText"):
            value = self._call(table, name, row, col)
            if isinstance(value, str):
                return value.strip()
        return ""

    def _insert_bend_table(self, view, steps):
        templates = []
        if self.bend_template:
            templates.append(self.bend_template)
        templates.append("")

        for template in templates:
            try:
                table = view.InsertBendTable(True, 0.2, 0.2, ANCHOR_TOP_LEFT, template)
            except Exception as exc:
                steps.append(f"InsertBendTable(template={template!r}) failed: {exc}")
                continue
            if table is not None:
                return table
            steps.append(f"InsertBendTable(template={template!r}) returned None")
        return None

    def _read_table(self, table, steps):
        row_count = self._get(table, "RowCount")
        col_count = self._get(table, "ColumnCount")
        steps.append(f"table size: rows={row_count} columns={col_count}")

        if (
            not isinstance(row_count, int)
            or not isinstance(col_count, int)
            or row_count < 1
            or col_count < 1
        ):
            return None

        header = [self._cell_text(table, 0, c) for c in range(col_count)]
        rows = []
        for r in range(1, row_count):
            cells = [self._cell_text(table, r, c) for c in range(col_count)]
            if any(cells):
                rows.append(cells)

        return {"header": header, "rows": rows}

    def read_part(self, item):
        """Returns (table_dict_or_None, message, steps)."""
        steps = []
        path = Path(item["path"]).resolve()
        config = item.get("configuration") or ""

        model, opened_by_us = self._open_part(path)
        drawing = None
        try:
            if model is None:
                return None, "part could not be opened", steps

            self._resolve_templates(steps)
            if not self.drawing_template:
                return (
                    None,
                    "no drawing template found (set DRAWING_TEMPLATE_PATH in main.py)",
                    steps,
                )

            drawing = self.sw_app.NewDocument(self.drawing_template, 0, 0.0, 0.0)
            if drawing is None:
                return None, "NewDocument returned None", steps
            steps.append("temporary drawing created")

            view = drawing.CreateFlatPatternViewFromModelView3(
                str(path), str(config), 0.15, 0.15, 0.0, False, False
            )
            if view is None:
                return (
                    None,
                    "flat pattern view was not created (does the part have a flat pattern?)",
                    steps,
                )
            steps.append("flat pattern view created")

            table_obj = self._insert_bend_table(view, steps)
            if table_obj is None:
                return None, "bend table could not be inserted", steps
            steps.append("bend table inserted")

            table = self._read_table(table_obj, steps)
            if table is None:
                return None, "bend table could not be read", steps
            if not table["rows"]:
                return None, "bend table is empty (no bends?)", steps

            return table, "OK", steps

        except Exception as exc:
            return None, f"error: {type(exc).__name__}: {exc}", steps
        finally:
            if drawing is not None:
                self._close(drawing)  # closed without saving
            if opened_by_us and model is not None:
                self._close(model)
            del drawing, model
            gc.collect()

    # =====================================================
    # ALL PARTS
    # =====================================================

    def read_all(self, items, output_folder=None, limit=None):
        targets = list(items)
        if limit:
            targets = targets[: int(limit)]

        print("\n" + "=" * 60 + "\nBEND TABLES (read from SolidWorks)\n" + "=" * 60)
        if limit:
            print(f"Test mode: only the first {len(targets)} part(s)")

        ok = 0
        failed = []
        detailed = 0

        for index, item in enumerate(targets, 1):
            table, message, steps = self.read_part(item)
            label = f"[{index}/{len(targets)}] {item['name']}"

            if table:
                item["bend_table"] = table
                ok += 1
                print(f"{label} -> OK ({len(table['rows'])} bends)")
            else:
                failed.append((item["name"], message))
                print(f"{label} -> no bend table: {message}")
                if detailed < 2:
                    detailed += 1
                    for line in steps:
                        print(f"    [debug] {line}")

        print(f"\nBend tables read: {ok} / {len(targets)}")

        csv_path = None
        if output_folder:
            csv_path = self._write_csv(Path(output_folder), targets)

        return {"read": ok, "failed": failed, "csv_path": csv_path}

    @staticmethod
    def _write_csv(output_folder, items):
        path = output_folder / "bend_tables.csv"
        try:
            output_folder.mkdir(parents=True, exist_ok=True)
            with path.open("w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(
                    [
                        "Part",
                        "Configuration",
                        "Tag",
                        "Direction",
                        "Angle",
                        "Inner Radius",
                    ]
                )
                for item in items:
                    table = item.get("bend_table")
                    if not table:
                        continue
                    for row in table["rows"]:
                        writer.writerow(
                            [item["name"], item.get("configuration") or ""] + list(row)
                        )
            print(f"Bend table CSV saved: {path}")
            return path
        except Exception as exc:
            print(f"Could not save bend table CSV: {exc}")
            return None
