"""Test setup: make the project importable and stub pywin32 when it is not installed (Linux / CI).

The tests use fake SolidWorks objects (tests/fakesw.py). They check the Python logic only;
they do NOT prove that a real SolidWorks drawing is correct.
"""
import sys
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
TESTS = Path(__file__).resolve().parent
if str(TESTS) not in sys.path:
    sys.path.insert(0, str(TESTS))

try:  # real pywin32 (Windows) works with the fakes as well
    import pythoncom  # noqa: F401
    from win32com.client import VARIANT  # noqa: F401
except ImportError:
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.VT_BYREF, pythoncom.VT_I4, pythoncom.VT_R8 = 0x4000, 3, 5
    pythoncom.VT_BOOL, pythoncom.VT_ARRAY, pythoncom.VT_DISPATCH = 11, 0x2000, 9
    pythoncom.CoInitialize = pythoncom.CoUninitialize = pythoncom.CoFreeUnusedLibraries = lambda: None

    class VARIANT:  # minimal stand-in for win32com.client.VARIANT
        def __init__(self, vt, value):
            self.varianttype, self.value = vt, value

    win32com = types.ModuleType("win32com")
    client = types.ModuleType("win32com.client")
    client.VARIANT = VARIANT
    client.Dispatch = client.GetActiveObject = lambda name: None
    win32com.client = client
    sys.modules.update({"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client})
