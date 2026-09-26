from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class AssemblyJob:
    source: Path
    output_dir: Path


@dataclass
class SheetMetalComponent:
    name: str
    path: Path
    configuration: str
    thickness: Optional[float] = None


@dataclass
class ExportResult:
    component: SheetMetalComponent
    output: Optional[Path]
    success: bool
    skipped: bool = False
    error: Optional[str] = None