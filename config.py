from pathlib import Path
    

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"

ASSEMBLY_EXTENSION = ".sldasm"
OUTPUT_EXTENSION = ".dxf"

RECURSIVE_SCAN = True
SKIP_EXISTING = True
CREATE_OUTPUT_DIRS = True

# Sheet metal DXF export options
#
# Bit 1 = flat-pattern geometry
# Bit 3 = bend lines
#
# 1  = geometry only
# 5  = geometry + bend lines
DXF_SHEET_METAL_OPTIONS = 1

# False = export one DXF per part
# True  = allow single-file export
EXPORT_TO_SINGLE_FILE = True

# Remove duplicate Part + Configuration combinations
DEDUPLICATE_COMPONENTS = True