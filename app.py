import logging
from pathlib import Path

from config import (
    INPUT_DIR,
    OUTPUT_DIR,
    LOG_DIR,
    SKIP_EXISTING,
    CREATE_OUTPUT_DIRS,
)

from core.scanner import DrawingScanner
from core.models import (
    ConversionResult,
    SheetResult,
)

from solidworks.connection import SolidWorksConnection
from solidworks.drawing import SolidWorksDrawing


def setup_logging():

    if CREATE_OUTPUT_DIRS:
        LOG_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        OUTPUT_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

    log_file = LOG_DIR / "exporter.log"

    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | "
            "%(levelname)s | "
            "%(message)s"
        ),
        handlers=[
            logging.FileHandler(
                log_file,
                encoding="utf-8",
            ),
            logging.StreamHandler(),
        ],
    )


def safe_filename(name: str) -> str:

    invalid = '<>:"/\\|?*'

    for char in invalid:
        name = name.replace(char, "_")

    name = name.strip().rstrip(".")

    if not name:
        name = "Sheet"

    return name


def process_drawing(
    drawing_manager: SolidWorksDrawing,
    source: Path,
) -> ConversionResult:

    result = ConversionResult(
        source=source,
        success=True,
    )

    logging.info(
        "Opening drawing: %s",
        source,
    )

    try:

        drawing_manager.open(source)

        sheets = drawing_manager.get_sheet_names()

        logging.info(
            "Found %d sheet(s): %s",
            len(sheets),
            sheets,
        )

        drawing_output_dir = (
            OUTPUT_DIR / source.stem
        )

        drawing_output_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        for sheet_name in sheets:

            safe_sheet_name = safe_filename(
                sheet_name
            )

            output_path = (
                drawing_output_dir
                / f"{safe_sheet_name}.dwg"
            )

            if (
                SKIP_EXISTING
                and output_path.exists()
            ):

                logging.info(
                    "SKIPPED existing: %s",
                    output_path,
                )

                result.sheets.append(
                    SheetResult(
                        sheet_name=sheet_name,
                        output=output_path,
                        success=True,
                        skipped=True,
                    )
                )

                continue

            try:

                logging.info(
                    "Exporting sheet '%s' -> %s",
                    sheet_name,
                    output_path,
                )

                exported = (
                    drawing_manager.export_sheet(
                        sheet_name,
                        output_path,
                    )
                )

                logging.info(
                    "Exported successfully: %s",
                    exported,
                )

                result.sheets.append(
                    SheetResult(
                        sheet_name=sheet_name,
                        output=exported,
                        success=True,
                    )
                )

            except Exception as exc:

                result.success = False

                logging.exception(
                    "Failed to export sheet '%s'",
                    sheet_name,
                )

                result.sheets.append(
                    SheetResult(
                        sheet_name=sheet_name,
                        output=None,
                        success=False,
                        error=str(exc),
                    )
                )

    except Exception as exc:

        result.success = False
        result.error = str(exc)

        logging.exception(
            "Failed to process drawing: %s",
            source,
        )

    finally:

        drawing_manager.close()

        logging.info(
            "Drawing closed: %s",
            source,
        )

    return result


def print_summary(results):

    total_drawings = len(results)

    successful_drawings = sum(
        1
        for result in results
        if result.success
    )

    total_sheets = sum(
        len(result.sheets)
        for result in results
    )

    successful_sheets = sum(
        1
        for result in results
        for sheet in result.sheets
        if sheet.success
    )

    skipped_sheets = sum(
        1
        for result in results
        for sheet in result.sheets
        if sheet.skipped
    )

    failed_sheets = sum(
        1
        for result in results
        for sheet in result.sheets
        if not sheet.success
    )

    print()
    print("=" * 60)
    print("EXPORT SUMMARY")
    print("=" * 60)

    print(
        f"Drawings : {total_drawings}"
    )

    print(
        f"Successful drawings : "
        f"{successful_drawings}"
    )

    print(
        f"Sheets : {total_sheets}"
    )

    print(
        f"Successful sheets : "
        f"{successful_sheets}"
    )

    print(
        f"Skipped sheets : "
        f"{skipped_sheets}"
    )

    print(
        f"Failed sheets : "
        f"{failed_sheets}"
    )

    print("=" * 60)

    for result in results:

        if result.success:
            status = "SUCCESS"
        else:
            status = "FAILED"

        print(
            f"{status}: {result.source}"
        )

        for sheet in result.sheets:

            if sheet.skipped:
                status = "SKIPPED"

            elif sheet.success:
                status = "OK"

            else:
                status = "FAILED"

            print(
                f"    [{status}] "
                f"{sheet.sheet_name}"
            )

            if sheet.error:
                print(
                    f"        {sheet.error}"
                )


def main():

    setup_logging()

    print(
        f"INPUT_DIR = {INPUT_DIR}"
    )

    scanner = DrawingScanner(
        INPUT_DIR
    )

    jobs = scanner.create_jobs()

    print(
        f"Found {len(jobs)} drawing(s)."
    )

    if not jobs:
        print(
            "No drawings found."
        )
        return

    connection = SolidWorksConnection()

    connection.connect()

    print(
        "Connected to SOLIDWORKS."
    )

    drawing_manager = SolidWorksDrawing(
        connection
    )

    results = []

    for job in jobs:

        print()
        print(
            f"Opening: {job.source}"
        )

        result = process_drawing(
            drawing_manager,
            job.source,
        )

        results.append(result)

    print_summary(results)


if __name__ == "__main__":
    main()