from pathlib import Path

import pythoncom
import win32com.client

from solidworks.connection import SolidWorksConnection


class SolidWorksDrawing:

    SW_DOC_DRAWING = 3
    SW_OPEN_DOC_OPTIONS_SILENT = 1

    SW_SAVE_CURRENT_VERSION = 0
    SW_SAVE_SILENT = 1

    def __init__(self, connection: SolidWorksConnection):
        self.connection = connection
        self.document = None
        self.source_path: Path | None = None

    def open(self, drawing_path: Path):

        drawing_path = drawing_path.resolve()

        if not drawing_path.exists():
            raise FileNotFoundError(
                f"Drawing file does not exist: {drawing_path}"
            )

        if drawing_path.suffix.lower() != ".slddrw":
            raise ValueError(
                f"Unsupported file type: {drawing_path.suffix}"
            )

        sw = self.connection.app

        if sw is None:
            raise RuntimeError(
                "SOLIDWORKS is not connected."
            )

        sw.SetCurrentWorkingDirectory(
            str(drawing_path.parent)
        )

        errors = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0
        )

        warnings = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0
        )

        self.document = sw.OpenDoc6(
            str(drawing_path),
            self.SW_DOC_DRAWING,
            self.SW_OPEN_DOC_OPTIONS_SILENT,
            "",
            errors,
            warnings
        )

        if self.document is None:
            raise RuntimeError(
                f"Could not open drawing: {drawing_path}"
            )

        self.source_path = drawing_path

        return self.document

    def get_sheet_names(self) -> tuple[str, ...]:

        if self.document is None:
            raise RuntimeError(
                "No drawing is currently open."
            )

        return tuple(
            self.document.GetSheetNames
        )

    def activate_sheet(self, sheet_name: str):

        if self.document is None:
            raise RuntimeError(
                "No drawing is currently open."
            )

        result = self.document.ActivateSheet(
            sheet_name
        )

        if result is False:
            raise RuntimeError(
                f"Could not activate sheet: {sheet_name}"
            )

    def export_sheet(
        self,
        sheet_name: str,
        output_path: Path,
    ) -> Path:

        if self.document is None:
            raise RuntimeError(
                "No drawing is currently open."
            )

        output_path = output_path.resolve()

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.activate_sheet(sheet_name)

        errors = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0
        )

        warnings = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0
        )

        success = self.document.Extension.SaveAs(
            str(output_path),
            self.SW_SAVE_CURRENT_VERSION,
            self.SW_SAVE_SILENT,
            None,
            errors,
            warnings,
        )

        if not success:
            raise RuntimeError(
                f"Failed to export sheet '{sheet_name}'. "
                f"Errors: {errors}, "
                f"Warnings: {warnings}"
            )

        if not output_path.exists():
            raise RuntimeError(
                f"SOLIDWORKS reported success but output "
                f"was not created: {output_path}"
            )

        return output_path

    def close(self):

        if self.document is None:
            return

        sw = self.connection.app

        if sw is None:
            self.document = None
            self.source_path = None
            return

        try:
            title = self.document.GetTitle
            sw.CloseDoc(title)

        finally:
            self.document = None
            self.source_path = None