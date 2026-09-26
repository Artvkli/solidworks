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
    sheet_name: str
    output: Optional[Path]
    success: bool
    skipped: bool = False
    error: Optional[str] = None


@dataclass
class ConversionResult:
    source: Path
    success: bool
    sheets: list[SheetResult] = field(default_factory=list)
    error: Optional[str] = None