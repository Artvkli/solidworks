from config import INPUT_DIR
from core.scanner import DrawingScanner
from solidworks.connection import SolidWorksConnection
from solidworks.drawing import SolidWorksDrawing


def main():

    scanner = DrawingScanner(INPUT_DIR)
    print("INPUT_DIR =", INPUT_DIR)

    jobs = scanner.create_jobs()

    print(f"Found {len(jobs)} drawing(s).")

    if not jobs:
        print("No drawings found.")
        return

    connection = SolidWorksConnection()
    connection.connect()

    print("Connected to SOLIDWORKS.")

    drawing_manager = SolidWorksDrawing(connection)

    for job in jobs:

        print()
        print(f"Opening: {job.source}")

        try:
            document = drawing_manager.open(
                job.source
            )

            print("Opened successfully.")
            sheet_names = document.GetSheetNames()
            print("Sheets:")
            for sheet_name in sheet_names:
                print("  -", sheet_name)
            print("Title:", document.GetTitle)
            print("Output:", job.output)

        except Exception as exc:

            print(
                f"ERROR: {job.source}"
            )

            print(exc)

        finally:

            drawing_manager.close()

            print("Drawing closed.")


if __name__ == "__main__":
    main()

