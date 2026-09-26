from pathlib import Path

from config import (
    ASSEMBLY_EXTENSION,
    OUTPUT_DIR,
    RECURSIVE_SCAN,
)

from core.models import AssemblyJob


class AssemblyScanner:

    def __init__(self, input_dir: Path):
        self.input_dir = input_dir

    def validate(self) -> None:

        if not self.input_dir.exists():
            raise FileNotFoundError(
                f"Input directory does not exist: "
                f"{self.input_dir}"
            )

        if not self.input_dir.is_dir():
            raise NotADirectoryError(
                f"Input path is not a directory: "
                f"{self.input_dir}"
            )

    def scan(self) -> list[Path]:

        self.validate()

        if RECURSIVE_SCAN:
            candidates = self.input_dir.rglob("*")
        else:
            candidates = self.input_dir.glob("*")

        assemblies = [
            path
            for path in candidates
            if (
                path.is_file()
                and not path.name.startswith("~$")
                and path.suffix.lower()
                == ASSEMBLY_EXTENSION.lower()
            )
        ]

        return sorted(assemblies)

    def create_jobs(self) -> list[AssemblyJob]:

        assemblies = self.scan()

        jobs = []

        for assembly in assemblies:

            relative = assembly.relative_to(
                self.input_dir
            )

            output_dir = (
                OUTPUT_DIR
                / relative.parent
                / assembly.stem
            )

            jobs.append(
                AssemblyJob(
                    source=assembly,
                    output_dir=output_dir,
                )
            )

        return jobs