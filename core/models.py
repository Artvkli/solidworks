from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class DrawingJob:

    source: Path
    output: Path

    status: str = "pending"
    error: Optional[str] = None


@dataclass
class SheetResult:

    name: str
    output: Optional[Path]

    success: bool
    message: str


@dataclass
class ConversionResult:

    source: Path
    output: Optional[Path]

    success: bool
    message: str

    sheets: list[SheetResult] = field(default_factory=list)