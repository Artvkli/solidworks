import win32com.client


class SolidWorksConnection:

    def __init__(self):
        self.app = None

    def connect(self):
        if self.app is not None:
            return self.app

        try:
            self.app = win32com.client.GetActiveObject(
                "SldWorks.Application"
            )

        except Exception as exc:
            raise RuntimeError(
                "Could not connect to SOLIDWORKS. "
                "Make sure SOLIDWORKS is running."
            ) from exc

        return self.app

    def is_connected(self) -> bool:
        return self.app is not None