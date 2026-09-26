import logging

from config import (
    INPUT_DIR,
    OUTPUT_DIR,
    LOG_DIR,
    SKIP_EXISTING,
    CREATE_OUTPUT_DIRS,
    DXF_SHEET_METAL_OPTIONS,
    DEDUPLICATE_COMPONENTS,
)

from core.scanner import AssemblyScanner

from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly
from solidworks.sheet_metal import SheetMetalProcessor


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

    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | "
            "%(levelname)s | "
            "%(message)s"
        ),
        handlers=[
            logging.FileHandler(
                LOG_DIR / "exporter.log",
                encoding="utf-8",
            ),
            logging.StreamHandler(),
        ],
    )


def main():

    setup_logging()

    print(
        f"INPUT_DIR = {INPUT_DIR}"
    )

    scanner = AssemblyScanner(
        INPUT_DIR
    )

    jobs = scanner.create_jobs()

    print(
        f"Found {len(jobs)} assembly(s)."
    )

    if not jobs:
        print(
            "No assemblies found."
        )
        return

    connection = SolidWorksConnection()

    connection.connect()

    print(
        "Connected to SOLIDWORKS."
    )

    assembly_manager = SolidWorksAssembly(
        connection
    )

    processor = SheetMetalProcessor(
        connection
    )

    total_components = 0
    sheet_metal_count = 0
    exported_count = 0
    skipped_count = 0
    failed_count = 0

    for job in jobs:

        print()
        print("=" * 70)
        print(
            f"Assembly: {job.source}"
        )
        print("=" * 70)

        logging.info(
            "Opening assembly: %s",
            job.source,
        )

        try:

            assembly_manager.open(
                job.source
            )

            assembly_manager.resolve_components()

            components = (
                assembly_manager.get_components()
            )

            total_components += len(
                components
            )

            logging.info(
                "Found %d component(s).",
                len(components),
            )

            processed = set()

            for index, component in enumerate(
                components,
                start=1,
            ):

                try:

                    data = (
                        processor.process_component(
                            component,
                            job.output_dir,
                            DXF_SHEET_METAL_OPTIONS,
                        )
                    )

                    if data is None:
                        continue

                    sheet_metal_count += 1

                    path = data["path"]
                    configuration = (
                        data["configuration"]
                    )
                    thickness = data["thickness"]
                    output_path = data["output"]

                    unique_key = (
                        str(path.resolve()).lower(),
                        configuration.lower(),
                    )

                    logging.info(
                        "[%d/%d] Sheet Metal: %s | "
                        "Configuration: %s | "
                        "Thickness: %s",
                        index,
                        len(components),
                        path.name,
                        configuration,
                        thickness,
                    )

                    if (
                        DEDUPLICATE_COMPONENTS
                        and unique_key in processed
                    ):

                        logging.info(
                            "Duplicate component skipped: %s",
                            path.name,
                        )

                        skipped_count += 1
                        continue

                    processed.add(unique_key)

                    if (
                        SKIP_EXISTING
                        and output_path.exists()
                    ):

                        logging.info(
                            "Existing DXF skipped: %s",
                            output_path,
                        )

                        skipped_count += 1
                        continue

                    logging.info(
                        "Exporting: %s",
                        output_path,
                    )

                    processor.export_dxf(
                        data["model"],
                        path,
                        output_path,
                        DXF_SHEET_METAL_OPTIONS,
                    )

                    exported_count += 1

                    print(
                        f"  OK  | "
                        f"{path.name} | "
                        f"{thickness} mm | "
                        f"{output_path}"
                    )

                except Exception as exc:

                    failed_count += 1

                    logging.exception(
                        "Component processing failed."
                    )

                    print(
                        f"  ERROR | "
                        f"Component #{index} | "
                        f"{exc}"
                    )

        except Exception as exc:

            failed_count += 1

            logging.exception(
                "Assembly processing failed: %s",
                job.source,
            )

            print(
                f"ERROR: {exc}"
            )

        finally:

            assembly_manager.close()

            logging.info(
                "Assembly closed: %s",
                job.source,
            )

    print()
    print("=" * 70)
    print("EXPORT SUMMARY")
    print("=" * 70)

    print(
        f"Assemblies       : {len(jobs)}"
    )

    print(
        f"Components       : {total_components}"
    )

    print(
        f"Sheet Metal      : {sheet_metal_count}"
    )

    print(
        f"Exported DXF     : {exported_count}"
    )

    print(
        f"Skipped          : {skipped_count}"
    )

    print(
        f"Failed           : {failed_count}"
    )

    print("=" * 70)


if __name__ == "__main__":
    main()