import win32com.client


class SolidWorksConnection:

    def __init__(self):
        self.sw_app = None

    def connect(self):
        """Connect to an already running SolidWorks instance."""

        try:
            self.sw_app = win32com.client.GetActiveObject(
                "SldWorks.Application"
            )

            print("Connected to SolidWorks successfully!")
            print("Version:", self.sw_app.RevisionNumber)

            return True

        except Exception as e:
            print("Could not connect to SolidWorks.")
            print("Error:", e)

            return False

    def get_application(self):
        """Return SolidWorks application object."""

        return self.sw_app