from pathlib import Path

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

        document_type = 3
        open_options = 1

        self.document = sw.OpenDoc6(
            str(drawing_path),
            document_type,
            open_options,
            "",
            0,
            0
        )

        if self.document is None:
            raise RuntimeError(
                f"Could not open drawing: {drawing_path}"
            )

        return self.document

    def close(self):

        if self.document is None:
            return

        sw = self.connection.app

        if sw is None:
            self.document = None
            return

        title = self.document.GetTitle()

        sw.CloseDoc(title)

        self.document = None