from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"


DRAWING_EXTENSION = ".slddrw"
OUTPUT_EXTENSION = ".dwg"


# Scanner
RECURSIVE_SCAN = True
SKIP_EXISTING = True


# Export
DWG_SAVE_VERSION = 0
DWG_SAVE_OPTIONS = 1


# Runtime
CREATE_OUTPUT_DIRS = True