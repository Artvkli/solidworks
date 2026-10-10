import csv
import gc
import time
from pathlib import Path

import pythoncom
from win32com.client import VARIANT

SW_DOC_PART = 1
SW_OPEN_SILENT = 1
SW_OPEN_READONLY = 2
ANCHOR_TOP_LEFT = 1  # swBOMConfigurationAnchor_TopLeft
FIRST_TAG = "A"  # first bend tag letter
SW_UNSUPPRESS = 0
SW_SUPPRESS = 1
SW_THIS_CONFIGURATION = 1
NO_OPTION_SPECIFIED = 0  # swUserPreferenceOption_e.swDetailingNoOptionSpecified
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
        self._last_accessor = None
        self._constants_cache = None  # None = not tried yet, False = not available

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
    # UNITS  (the bend table shows values in the DRAWING's units)
    # =====================================================

    def _constants(self):
        """
        SolidWorks enumeration values (swUnitSystem, swUnitsLinear, ...), read from the
        swconst.tlb that is installed with YOUR SolidWorks, so the numbers always match its version.
        """
        if self._constants_cache is not None:
            return self._constants_cache or None
        self._constants_cache = False

        try:
            import win32com.client as client
            from win32com.client import gencache

            exe = self._get(self.sw_app, "GetExecutablePath")
            if not exe:
                return None
            install = Path(str(exe))
            if install.suffix:
                install = install.parent

            for tlb_path in sorted(install.glob("*.tlb")):
                if "swconst" not in tlb_path.name.lower():
                    continue
                tlb = pythoncom.LoadTypeLib(str(tlb_path))
                guid, lcid, _syskind, major, minor, _flags = tlb.GetLibAttr()
                gencache.EnsureModule(str(guid), lcid, major, minor)
                try:
                    getattr(client.constants, "swUnitSystem_MMGS")
                except AttributeError:
                    continue
                self._constants_cache = client.constants
                return self._constants_cache
        except Exception:
            pass
        return None

    def _use_millimeters(self, drawing, steps):
        """
        The default drawing template may be in INCHES: a 2 mm radius then shows as 0.08.
        Switch the temporary drawing to MMGS (millimetres) before the bend table is created.
        """
        constants = self._constants()
        if constants is None:
            steps.append(
                "units: SolidWorks constants (swconst.tlb) could not be loaded - the drawing units were NOT "
                "changed. If radii look like inches (2 mm -> 0.08), set DRAWING_TEMPLATE_PATH to a metric template."
            )
            return False

        try:
            extension = self._call(drawing, "Extension")
            extension.SetUserPreferenceInteger(
                getattr(constants, "swUnitSystem"),
                NO_OPTION_SPECIFIED,
                getattr(constants, "swUnitSystem_MMGS"),
            )
            extension.SetUserPreferenceInteger(
                getattr(constants, "swUnitsLinear"),
                NO_OPTION_SPECIFIED,
                getattr(constants, "swMM"),
            )
            current = extension.GetUserPreferenceInteger(
                getattr(constants, "swUnitsLinear"), NO_OPTION_SPECIFIED
            )
            ok = current == getattr(constants, "swMM")
            steps.append(
                f"units: drawing linear unit = {'millimetres' if ok else f'NOT millimetres (code {current})'}"
            )
            return ok
        except Exception as exc:
            steps.append(f"units: could not set millimetres: {exc}")
            return False

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
        # ITableAnnotation: DisplayedText2(Row, Column, IncludeHidden) is the current accessor,
        # DisplayedText(Row, Column) is the older (obsolete) one; Text* give the driving string.
        attempts = (
            ("DisplayedText2", (row, col, True)),
            ("DisplayedText", (row, col)),
            ("Text2", (row, col, True)),
            ("Text", (row, col)),
        )
        for name, args in attempts:
            value = self._call(table, name, *args)
            if isinstance(value, str):
                self._last_accessor = name
                return value.strip()
        return ""

    def _insert_bend_table(self, view, steps):
        # IView.InsertBendTable(UseAnchorPoint, X, Y, AnchorType, StartValue, TableTemplate)
        # StartValue is the first tag ("A"); TableTemplate is the full path of a .sldbndtbt file.
        templates = [self.bend_template] if self.bend_template else [""]

        for template in templates:
            for anchor in (ANCHOR_TOP_LEFT, 0, 2):
                try:
                    table = view.InsertBendTable(
                        False, 0.05, 0.05, anchor, FIRST_TAG, template
                    )
                except Exception as exc:
                    steps.append(f"InsertBendTable(anchor={anchor}) failed: {exc}")
                    continue
                if table is not None:
                    return table
                steps.append(f"InsertBendTable(anchor={anchor}) returned None")
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

        self._last_accessor = None
        header = [self._cell_text(table, 0, c) for c in range(col_count)]
        rows = []
        for r in range(1, row_count):
            cells = [self._cell_text(table, r, c) for c in range(col_count)]
            if any(cells):
                rows.append(cells)

        steps.append(f"cells read with: {self._last_accessor}")
        return {"header": header, "rows": rows}

    def _create_flat_view(self, drawing, path, config):
        """CreateFlatPatternViewFromModelView3, first with the part's configuration, then with ''."""
        configs = []
        for name in (config, ""):
            if name not in configs:
                configs.append(name)

        for round_number in range(2):
            for name in configs:
                try:
                    view = drawing.CreateFlatPatternViewFromModelView3(
                        str(path), str(name), 0.15, 0.15, 0.0, False, False
                    )
                except Exception:
                    view = None
                if view is not None:
                    return view
            if round_number == 0:
                time.sleep(0.5)
        return None

    def read_part(self, item):
        """Returns (table_dict_or_None, message, steps)."""
        steps = []
        path = Path(item["path"]).resolve()
        config = item.get("configuration") or ""

        model, opened_by_us = self._open_part(path)
        drawing = None
        restore_flat = None
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

            # make sure the part (and its flat pattern) is up to date before a drawing uses it
            self._call(model, "ForceRebuild3", True)

            # In two real runs the FIRST part of the run failed to get a flat pattern view while all the
            # following ones worked. So: if the view is not created, close that drawing, wait a moment
            # and try again with a brand new temporary drawing.
            view = None
            for attempt in (1, 2):
                drawing = self.sw_app.NewDocument(self.drawing_template, 0, 0.0, 0.0)
                if drawing is None:
                    return None, "NewDocument returned None", steps
                steps.append(f"temporary drawing created (attempt {attempt})")
                self._use_millimeters(drawing, steps)

                view = self._create_flat_view(drawing, path, config)
                if view is None:
                    flat = self._call(model, "FeatureByName", "Flat-Pattern1")
                    if (
                        flat is not None
                        and restore_flat is None
                        and self._get(flat, "IsSuppressed", False)
                    ):
                        try:
                            flat.SetSuppression2(
                                SW_UNSUPPRESS, SW_THIS_CONFIGURATION, None
                            )
                            restore_flat = flat
                            steps.append(
                                "flat pattern was suppressed - un-suppressed for the view"
                            )
                        except Exception as exc:
                            steps.append(
                                f"could not un-suppress the flat pattern: {exc}"
                            )
                        view = self._create_flat_view(drawing, path, config)

                if view is not None:
                    break

                steps.append(f"flat pattern view was not created on attempt {attempt}")
                if attempt == 1:
                    self._close(drawing)
                    drawing = None
                    time.sleep(1.0)

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
            if restore_flat is not None:
                try:
                    restore_flat.SetSuppression2(
                        SW_SUPPRESS, SW_THIS_CONFIGURATION, None
                    )
                except Exception:
                    pass
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
