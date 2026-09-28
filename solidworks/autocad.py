import os
import time
import pythoncom
import win32com.client


class AutoCADExporter:

    def __init__(self):
        self.app = None

    def connect(self):
        """Connect to AutoCAD."""
        try:
            try:
                self.app = win32com.client.GetActiveObject(
                    "AutoCAD.Application"
                )
            except Exception:
                self.app = win32com.client.Dispatch(
                    "AutoCAD.Application"
                )

            self.app.Visible = True

            print("Connected to AutoCAD successfully.")

            return True

        except Exception as e:
            print("Could not connect to AutoCAD.")
            print(e)
            return False

    def _point(self, x, y, z=0.0):
        """Create AutoCAD point."""
        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (float(x), float(y), float(z))
        )

    def create_dwg(self, output_path):
        """Create a new DWG document."""
        try:
            document = self.app.Documents.Add()

            document.SaveAs(str(output_path))

            return document

        except Exception as e:
            print("Could not create DWG:")
            print(e)
            return None

    def insert_dwg(self, document, dwg_path, x, y):
        """Insert a DWG as a block."""

        try:
            model_space = document.ModelSpace

            insertion_point = self._point(x, y)

            block = model_space.InsertBlock(
                insertion_point,
                str(dwg_path),
                1.0,
                1.0,
                1.0,
                0.0
            )

            time.sleep(0.2)

            return block

        except Exception as e:
            print()
            print("Could not insert DWG:")
            print(dwg_path)
            print(e)
            return None

    def get_extents(self, entity):
        """Get entity bounding box."""

        try:
            min_point, max_point = entity.GetBoundingBox()

            return (
                min_point[0],
                min_point[1],
                max_point[0],
                max_point[1]
            )

        except Exception as e:
            print("Could not get entity extents:")
            print(e)
            return None

    def add_qty_text(
        self,
        document,
        x,
        y,
        quantity
    ):
        """Add QTY text under a part."""

        try:
            model_space = document.ModelSpace

            text_point = self._point(x, y)

            text = model_space.AddText(
                f"QTY: {quantity}",
                text_point,
                10.0
            )

            return text

        except Exception as e:
            print("Could not add QTY text:")
            print(e)
            return None

    def create_thickness_dwg(
        self,
        output_path,
        parts,
        spacing=100.0
    ):
        """
        Create one DWG for one thickness.

        Each unique part is inserted once.
        Quantity is written underneath.
        """

        document = self.create_dwg(output_path)

        if document is None:
            return False

        current_x = 0.0

        try:

            for part in parts:

                dwg_path = part["dwg_path"]
                quantity = part["quantity"]

                print()
                print("Adding:", part["name"])
                print("QTY:", quantity)

                if not os.path.exists(dwg_path):
                    print("[MISSING DWG]")
                    continue

                # --------------------------------------
                # Insert part DWG
                # --------------------------------------

                block = self.insert_dwg(
                    document,
                    dwg_path,
                    current_x,
                    0.0
                )

                if block is None:
                    print("[INSERT FAILED]")
                    continue

                # --------------------------------------
                # Get dimensions
                # --------------------------------------

                extents = self.get_extents(block)

                if extents is None:
                    print("[EXTENTS FAILED]")
                    continue

                min_x, min_y, max_x, max_y = extents

                width = max_x - min_x

                # --------------------------------------
                # Add Quantity
                # --------------------------------------

                text_x = current_x
                text_y = min_y - 20.0

                self.add_qty_text(
                    document,
                    text_x,
                    text_y,
                    quantity
                )

                # --------------------------------------
                # Move next part
                # --------------------------------------

                current_x += width + spacing

            # ------------------------------------------
            # Save
            # ------------------------------------------

            document.Save()

            print()
            print("[DWG CREATED]")
            print(output_path)

            return True

        except Exception as e:

            print()
            print("Error while creating thickness DWG:")
            print(e)

            return False

        finally:

            try:
                document.Close(False)
            except Exception:
                pass
