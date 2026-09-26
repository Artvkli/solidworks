from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class DrawingJob:
    source: Path
    output: Path
    status: str = "pending"
    error: Optional[str] = None


@dataclass
class ConversionResult:
    source: Path
    output: Optional[Path]
    success: bool
    message: str