from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"

DRAWING_EXTENSION = ".slddrw"
OUTPUT_EXTENSION = ".dwg"

SKIP_EXISTING = True
RECURSIVE_SCAN = True

CREATE_OUTPUT_DIRS = True