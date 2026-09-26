from pathlib import Path

from config import (
    DRAWING_EXTENSION,
    OUTPUT_DIR,
    OUTPUT_EXTENSION,
    RECURSIVE_SCAN,
)

from core.models import DrawingJob


class DrawingScanner:

    def __init__(self, input_dir: Path):
        self.input_dir = input_dir

    def validate(self) -> None:
        if not self.input_dir.exists():
            raise FileNotFoundError(
                f"Input directory does not exist: {self.input_dir}"
            )

        if not self.input_dir.is_dir():
            raise NotADirectoryError(
                f"Input path is not a directory: {self.input_dir}"
            )

    def scan(self) -> list[Path]:
        self.validate()

        if RECURSIVE_SCAN:
            candidates = self.input_dir.rglob("*")
        else:
            candidates = self.input_dir.glob("*")

        drawings = [
            path
            for path in candidates
            if (
                path.is_file()
                and not path.name.startswith("~$")
                and path.suffix.lower() == DRAWING_EXTENSION.lower()
            )
        ]

        return sorted(drawings)

    def create_jobs(self) -> list[DrawingJob]:
        drawings = self.scan()

        jobs = []

        for drawing in drawings:

            relative_path = drawing.relative_to(
                self.input_dir
            )

            output_relative = relative_path.with_suffix(
                OUTPUT_EXTENSION
            )

            output_path = OUTPUT_DIR / output_relative

            jobs.append(
                DrawingJob(
                    source=drawing,
                    output=output_path,
                )
            )

        return jobs