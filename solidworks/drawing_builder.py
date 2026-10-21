"""Create an editable native SolidWorks drawing from a sheet-metal part.

Target: SolidWorks 2025 + Windows + pywin32.

This first-stage builder creates a native SLDDRW with a flat-pattern view and
Front/Top/Right model views, adds editable notes, saves the drawing, and exports
PDF. Bend-table insertion and automatic dimensioning are intentionally isolated
as follow-up API-specific steps because SolidWorks drawing-table APIs depend on
installed templates and version/API signatures.
"""
from __future__ import annotations

import gc
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client
from win32com.client import VARIANT

# SolidWorks constants (numeric values used to avoid requiring generated constants)
SW_DOC_DRAWING = 3
SW_OPEN_SILENT = 1
SW_SAVE_SILENT = 1
SW_SAVE_CURRENT_VERSION = 0
SW_EXPORT_PDF_DATA = 1
SW_STANDARD_SHEET_SIZE_A3 = 8  # fallback only; configured template is preferred


class SolidWorksDrawingBuilder:
    """Build an editable drawing using SolidWorks' native drawing-view objects."""

    def __init__(self, sw_app: Any, output_dir: str | Path,
                 drawing_template: str | Path | None = None,
                 sheet_name: str = "Sheet1"):
        self.sw = sw_app
        self.output_dir = Path(output_dir).resolve()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.template = Path(drawing_template).resolve() if drawing_template else None
        self.sheet_name = sheet_name

    @staticmethod
    def _byref_i4():
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    @staticmethod
    def _safe_call(obj, method: str, *args, default=None):
        try:
            return getattr(obj, method)(*args)
        except Exception:
            return default

    def _open_part(self, part_path: Path, configuration: str | None):
        errors, warnings = self._byref_i4(), self._byref_i4()
        model = self.sw.GetOpenDocumentByName(str(part_path))
        opened_here = False
        if model is None:
            model = self.sw.OpenDoc6(str(part_path), 1, SW_OPEN_SILENT, "", errors, warnings)
            opened_here = True
        if model is None:
            raise RuntimeError(
                f"SolidWorks could not open part: {part_path}; "
                f"errors={getattr(errors, 'value', errors)}"
            )
        # Some scan-cache records use the placeholder "Default", although
        # the part file has a differently named configuration.
        if configuration:
            requested = str(configuration).strip()
            ok = self._safe_call(model, "ShowConfiguration2", requested, default=None)
            active = self._safe_call(model, "GetActiveConfiguration", default=None)
            active_name = self._safe_call(active, "GetName", default=None)

            if not (ok is True or (active_name and str(active_name) == requested)):
                raw_names = self._safe_call(model, "GetConfigurationNames", default=[]) or []
                config_names = [str(n) for n in raw_names if n]
                if requested in config_names:
                    ok = self._safe_call(model, "ShowConfiguration2", requested, default=False)
                    active = self._safe_call(model, "GetActiveConfiguration", default=None)
                    active_name = self._safe_call(active, "GetName", default=None)
                    if not (ok is True or (active_name and str(active_name) == requested)):
                        raise RuntimeError(
                            f"Could not activate configuration '{requested}'. "
                            f"Available configurations: {config_names or 'unknown'}"
                        )
                else:
                    # Prefer the configuration already active in the document;
                    # otherwise use the first configuration actually defined.
                    fallback = str(active_name) if active_name else (
                        config_names[0] if config_names else None
                    )
                    if not fallback or fallback == requested:
                        raise RuntimeError(
                            f"Could not activate requested configuration '{requested}'. "
                            f"Available configurations: {config_names or 'unknown'}"
                        )
                    fallback_ok = self._safe_call(
                        model, "ShowConfiguration2", fallback, default=False
                    )
                    active = self._safe_call(model, "GetActiveConfiguration", default=None)
                    active_name = self._safe_call(active, "GetName", default=None)
                    if fallback_ok is True or (active_name and str(active_name) == fallback):
                        print(
                            f"  [config] '{requested}' is not available; "
                            f"using existing configuration '{fallback}'"
                        )
                    else:
                        raise RuntimeError(
                            f"Could not activate requested configuration '{requested}'. "
                            f"Available configurations: {config_names or 'unknown'}"
                        )
        self._safe_call(model, "ForceRebuild3", True)
        return model, opened_here

    def _new_drawing(self):
        template = str(self.template) if self.template and self.template.exists() else ""
        drawing = self.sw.NewDocument(template, 0, 0.0, 0.0)
        if drawing is None:
            raise RuntimeError(
                "NewDocument failed. Set drawing_template to a valid .drwdot file "
                "from this SolidWorks installation."
            )
        return drawing

    def _add_note(self, drawing, text: str, x: float, y: float, height: float = 0.006):
        """Insert an editable note at sheet coordinates in metres."""
        drawing.ClearSelection2(True)
        note = drawing.InsertNote(str(text))
        if note is None:
            return False
        ann = self._safe_call(note, "GetAnnotation")
        if ann is not None:
            self._safe_call(ann, "SetPosition", float(x), float(y), 0.0)
            self._safe_call(ann, "SetTextHeight", float(height))
        return True

    def build(self, part_path: str | Path, configuration: str | None = None,
              part_number: str | None = None, quantity: int = 1,
              thickness_mm: float | None = None, export_pdf: bool = True,
              flat_xy=(0.18, 0.19), front_xy=(0.18, 0.055),
              top_xy=(0.38, 0.055), right_xy=(0.56, 0.055),
              view_scale: float = 0.12) -> dict[str, str]:
        """Create drawing and return output paths.

        Coordinates are in metres. Adjust positions/scale to match the chosen
        drawing template. This creates native linked views, not DXF linework.
        """
        part_path = Path(part_path).resolve()
        if not part_path.is_file():
            raise FileNotFoundError(part_path)
        stem = part_number or part_path.stem
        drawing_path = self.output_dir / f"{stem}.SLDDRW"
        pdf_path = self.output_dir / f"{stem}.PDF"

        part = drawing = None
        opened_here = False
        try:
            part, opened_here = self._open_part(part_path, configuration)
            drawing = self._new_drawing()

            # Flat-pattern view: model/configuration-linked and editable in drawing.
            flat = drawing.CreateFlatPatternViewFromModelView3(
                str(part_path), configuration or "", float(flat_xy[0]),
                float(flat_xy[1]), 0.0, False
            )
            if flat is None:
                raise RuntimeError(
                    "CreateFlatPatternViewFromModelView3 failed. Confirm the part has "
                    "a valid Flat-Pattern feature and try its flat-pattern configuration."
                )
            self._safe_call(flat, "SetScale", True, float(view_scale))

            # Folded views are native model views. View names are standard SW names.
            created_views = {"Flat-Pattern": flat}
            for view_name, xy in (("*Front", front_xy), ("*Top", top_xy), ("*Right", right_xy)):
                view = drawing.CreateDrawViewFromModelView3(
                    str(part_path), view_name, float(xy[0]), float(xy[1]), 0.0
                )
                if view is None:
                    raise RuntimeError(f"Could not create native drawing view {view_name}")
                self._safe_call(view, "SetScale", True, float(view_scale))
                created_views[view_name] = view

            # Editable text annotations. These are deliberately separate from geometry.
            self._add_note(drawing, "FLAT PATTERN", flat_xy[0], flat_xy[1] + 0.09)
            self._add_note(drawing, "FRONT VIEW", front_xy[0], front_xy[1] - 0.035)
            self._add_note(drawing, "TOP VIEW", top_xy[0], top_xy[1] - 0.035)
            self._add_note(drawing, "RIGHT VIEW", right_xy[0], right_xy[1] - 0.035)
            self._add_note(drawing, f"DWG NO. / PART: {stem}", 0.025, 0.018)
            self._add_note(drawing, f"QTY={int(quantity)}", 0.025, 0.045)
            if thickness_mm is not None:
                self._add_note(drawing, f"THICKNESS: {float(thickness_mm):g} mm", 0.22, 0.018)

            self._safe_call(drawing, "ForceRebuild3", True)
            self._safe_call(drawing, "ViewZoomtofit2")
            errors, warnings = self._byref_i4(), self._byref_i4()
            ok = drawing.Extension.SaveAs(
                str(drawing_path), SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT,
                None, errors, warnings
            )
            if not ok:
                raise RuntimeError(
                    f"SLDDRW save failed; errors={getattr(errors, 'value', errors)}, "
                    f"warnings={getattr(warnings, 'value', warnings)}"
                )

            result = {"slddrw": str(drawing_path)}
            if export_pdf:
                pdf_errors, pdf_warnings = self._byref_i4(), self._byref_i4()
                pdf_ok = drawing.Extension.SaveAs(
                    str(pdf_path), SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT,
                    None, pdf_errors, pdf_warnings
                )
                if pdf_ok:
                    result["pdf"] = str(pdf_path)
                else:
                    result["pdf_error"] = (
                        f"PDF export failed; errors={getattr(pdf_errors, 'value', pdf_errors)}"
                    )
            return result
        finally:
            # Leave the resulting drawing open for inspection; close only the part
            # if this builder opened it and it is not the active drawing document.
            if opened_here and part is not None:
                try:
                    title = part.GetTitle()
                    self.sw.CloseDoc(title)
                except Exception:
                    pass
            gc.collect()
            try:
                pythoncom.CoFreeUnusedLibraries()
            except Exception:
                pass


def build_from_scan_record(sw_app, record: dict, output_dir: str | Path,
                           drawing_template: str | Path | None = None,
                           export_pdf: bool = True) -> dict[str, str]:
    """Convenience adapter for SheetMetalDetector.scan_sheet_metal records."""
    builder = SolidWorksDrawingBuilder(sw_app, output_dir, drawing_template)
    return builder.build(
        part_path=record["path"],
        configuration=record.get("configuration"),
        part_number=Path(record["path"]).stem,
        quantity=record.get("quantity", 1),
        thickness_mm=record.get("thickness"),
        export_pdf=export_pdf,
    )
