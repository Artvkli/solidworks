from pathlib import Path

import pythoncom
import win32com.client

from solidworks.connection import SolidWorksConnection


class SolidWorksDrawing:

    def __init__(self, connection: SolidWorksConnection):
        self.connection = connection
        self.document = None

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

        # SOLIDWORKS constants
        SW_DOC_DRAWING = 3
        SW_OPEN_DOC_OPTIONS_SILENT = 1

        # OpenDoc6 output parameters must be passed ByRef
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
            SW_DOC_DRAWING,
            SW_OPEN_DOC_OPTIONS_SILENT,
            "",
            errors,
            warnings
        )

        if self.document is None:
            raise RuntimeError(
                f"Could not open drawing: {drawing_path}"
            )

        print("OpenDoc6 errors:", errors)
        print("OpenDoc6 warnings:", warnings)

        return self.document

    def close(self):

        if self.document is None:
            return

        sw = self.connection.app

        if sw is None:
            self.document = None
            return

        title = self.document.

        sw.CloseDoc(title)

        self.document = None