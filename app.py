import logging
import sys
from pathlib import Path

from config import (
    INPUT_DIR,
    OUTPUT_DIR,
    LOG_DIR,
    CREATE_OUTPUT_DIRS,
    DEDUPLICATE_COMPONENTS,
)

from core.scanner import AssemblyScanner

from solidworks.connection import (
    SolidWorksConnection,
)

from solidworks.assembly import (
    SolidWorksAssembly,
)

from solidworks.sheet_metal import (
    SheetMetalProcessor,
)


def setup_logging():

    LOG_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    log_file = (
        LOG_DIR
        / "solidworks_exporter.log"
    )

    logger = logging.getLogger()

    logger.setLevel(
        logging.INFO
    )

    formatter = logging.Formatter(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(message)s"
    )

    file_handler = logging.FileHandler(
        log_file,
        encoding="utf-8",
    )

    file_handler.setFormatter(
        formatter
    )

    console_handler = logging.StreamHandler(
        sys.stdout
    )

    console_handler.setFormatter(
        formatter
    )

    logger.handlers.clear()

    logger.addHandler(
        file_handler
    )

    logger.addHandler(
        console_handler
    )


def component_key(
    component,
    processor,
):

    path = processor.get_component_path(
        component
    )

    configuration = (
        processor.get_configuration(
            component
        )
    )

    if path is None:
        return None

    return (
        str(path.resolve()).lower(),
        configuration.lower(),
    )


def main():

    setup_logging()

    logger = logging.getLogger(
        __name__
    )

    print()
    print("=" * 70)
    print("SOLIDWORKS SHEET METAL EXPORTER")
    print("=" * 70)
    print()

    logger.info(
        "INPUT_DIR = %s",
        INPUT_DIR,
    )

    # -----------------------------------------------------
    # Scan assemblies
    # -----------------------------------------------------

    scanner = AssemblyScanner(
        INPUT_DIR
    )

    jobs = scanner.create_jobs()

    print(
        f"Found {len(jobs)} assembly(s)."
    )

    if not jobs:

        logger.warning(
            "No .SLDASM files found."
        )

        return

    # -----------------------------------------------------
    # Connect SOLIDWORKS
    # -----------------------------------------------------

    connection = (
        SolidWorksConnection()
    )

    connection.connect()

    print(
        "Connected to SOLIDWORKS."
    )

    # -----------------------------------------------------
    # Process
    # -----------------------------------------------------

    assembly_manager = (
        SolidWorksAssembly(
            connection
        )
    )

    processor = (
        SheetMetalProcessor()
    )

    total_assemblies = 0
    total_components = 0
    total_sheet_metal = 0
    total_failed = 0

    for job in jobs:

        total_assemblies += 1

        print()
        print("=" * 70)
        print(
            f"Assembly: {job.source}"
        )
        print("=" * 70)

        try:

            logger.info(
                "Opening assembly: %s",
                job.source,
            )

            assembly_manager.open(
                job.source
            )

            # Resolve lightweight components
            assembly_manager.resolve_components()

            components = (
                assembly_manager.get_components()
            )

            logger.info(
                "Found %d component(s).",
                len(components),
            )

            total_components += len(
                components
            )

            # ---------------------------------------------
            # Deduplication
            # ---------------------------------------------

            processed_keys = set()

            for index, component in enumerate(
                components,
                start=1,
            ):

                try:

                    name = (
                        processor
                        .get_component_name(
                            component
                        )
                    )

                    logger.info(
                        "[%d/%d] Component: %s",
                        index,
                        len(components),
                        name,
                    )

                    path = (
                        processor
                        .get_component_path(
                            component
                        )
                    )

                    logger.info(
                        "Component path: %s",
                        path,
                    )

                    # -------------------------------------
                    # Deduplicate Part + Configuration
                    # -------------------------------------

                    if DEDUPLICATE_COMPONENTS:

                        key = component_key(
                            component,
                            processor,
                        )

                        if key is not None:

                            if key in processed_keys:

                                logger.info(
                                    "Skipping duplicate: %s",
                                    name,
                                )

                                continue

                            processed_keys.add(
                                key
                            )

                    # -------------------------------------
                    # Process Sheet Metal
                    # -------------------------------------

                    result = (
                        processor
                        .process_component(
                            component,
                            job.output_dir,
                        )
                    )

                    if result is None:
                        continue

                    total_sheet_metal += 1

                    logger.info(
                        "Sheet Metal part found:"
                    )

                    logger.info(
                        "  Name        : %s",
                        result["name"],
                    )

                    logger.info(
                        "  Path        : %s",
                        result["path"],
                    )

                    logger.info(
                        "  Configuration: %s",
                        result["configuration"],
                    )

                    logger.info(
                        "  Thickness   : %.4f mm",
                        result["thickness_mm"],
                    )

                except Exception as exc:

                    total_failed += 1

                    logger.exception(
                        "Component processing failed: %s",
                        exc,
                    )

        except Exception as exc:

            total_failed += 1

            logger.exception(
                "Assembly processing failed: %s",
                exc,
            )

        finally:

            assembly_manager.close()

            logger.info(
                "Assembly closed: %s",
                job.source,
            )

    # -----------------------------------------------------
    # Summary
    # -----------------------------------------------------

    print()
    print("=" * 70)
    print("SCAN SUMMARY")
    print("=" * 70)

    print(
        f"Assemblies       : {total_assemblies}"
    )

    print(
        f"Components       : {total_components}"
    )

    print(
        f"Sheet Metal      : {total_sheet_metal}"
    )

    print(
        f"Failed           : {total_failed}"
    )

    print("=" * 70)


if __name__ == "__main__":
    main()