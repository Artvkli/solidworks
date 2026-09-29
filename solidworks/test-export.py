from pathlib import Path
import time
import pythoncom
import win32com.client


# ============================================================
# تنظیمات
# ============================================================

PART_PATH = Path(
    r"F:\solidworks\file\AHU 3500\AHU3500-08.SLDPRT"
)

OUTPUT_DIR = Path(
    r"F:\solidworks\solidworks\output"
)

# SolidWorks constants
SW_DOC_PART = 1

# ExportToDWG2
SW_EXPORT_SHEET_METAL = 1

# Sheet metal options:
# 1 = geometry
# 2 = hidden edges
# 4 = bend lines
# 8 = sketches
#
# 1 + 4 = 5
SHEET_METAL_OPTIONS = 5


# ============================================================
# Helper
# ============================================================

def safe_call(obj, method_name, *args):
    try:
        method = getattr(obj, method_name)
        return method(*args)
    except Exception as e:
        print(f"[ERROR] {method_name}: {e}")
        return None


def print_feature(feature, level=0):
    """
    چاپ Feature Tree به صورت recursive
    """

    if feature is None:
        return

    try:
        name = feature.Name
    except Exception:
        name = "?"

    try:
        type_name = feature.GetTypeName2
    except Exception:
        type_name = "?"

    indent = "  " * level

    print(
        f"{indent}- Name: {name}"
    )
    print(
        f"{indent}  Type: {type_name}"
    )

    # Sub features
    try:
        sub = feature.GetFirstSubFeature()

        while sub is not None:
            print_feature(sub, level + 1)

            try:
                sub = sub.GetNextSubFeature()
            except Exception:
                break

    except Exception:
        pass


def scan_features(model):
    """
    چاپ کل Feature Tree
    """

    print()
    print("=" * 70)
    print("FEATURE TREE")
    print("=" * 70)

    try:
        feature = model.FirstFeature
    except Exception as e:
        print("Cannot get FirstFeature:", e)
        return []

    features = []

    while feature is not None:

        try:
            name = feature.Name
        except Exception:
            name = "?"

        try:
            type_name = feature.GetTypeName2
        except Exception:
            type_name = "?"

        print()
        print(f"FEATURE: {name}")
        print(f"TYPE   : {type_name}")

        features.append(
            {
                "name": str(name),
                "type": str(type_name),
                "feature": feature,
            }
        )

        # چاپ SubFeature ها
        try:
            sub = feature.GetFirstSubFeature()

            while sub is not None:

                print_feature(sub, 1)

                try:
                    sub = sub.GetNextSubFeature()
                except Exception:
                    break

        except Exception:
            pass

        try:
            feature = feature.GetNextFeature()
        except Exception:
            break

    return features


def find_flat_patterns(model):
    """
    جستجوی FlatPattern در کل Feature Tree
    """

    print()
    print("=" * 70)
    print("SEARCHING FOR FLAT PATTERN")
    print("=" * 70)

    result = []

    def recursive(feature, level=0):

        if feature is None:
            return

        try:
            name = str(feature.Name)
        except Exception:
            name = ""

        try:
            type_name = str(feature.GetTypeName2)
        except Exception:
            type_name = ""

        print(
            f"{'  ' * level}"
            f"{name}  -->  {type_name}"
        )

        if (
            "flatpattern" in name.lower()
            or "flat-pattern" in name.lower()
            or "flatpattern" in type_name.lower()
            or "flat-pattern" in type_name.lower()
        ):
            print(
                f"{'  ' * level}"
                f">>> POSSIBLE FLAT PATTERN FOUND"
            )

            result.append(feature)

        # SubFeatures
        try:
            sub = feature.GetFirstSubFeature()

            while sub is not None:

                recursive(sub, level + 1)

                try:
                    sub = sub.GetNextSubFeature()
                except Exception:
                    break

        except Exception:
            pass

    try:
        feature = model.FirstFeature
    except Exception as e:
        print("Cannot access FirstFeature:", e)
        return result

    while feature is not None:

        recursive(feature)

        try:
            feature = feature.GetNextFeature()
        except Exception:
            break

    return result


def get_sheet_metal_bodies(model):

    print()
    print("=" * 70)
    print("CHECKING BODIES")
    print("=" * 70)

    result = []

    try:
        bodies = model.GetBodies2(0, True)
    except Exception as e:
        print("GetBodies2 failed:", e)
        return result

    if bodies is None:
        print("No bodies found.")
        return result

    for index, body in enumerate(bodies, start=1):

        try:
            name = body.Name
        except Exception:
            name = "?"

        try:
            is_sheet_metal = body.IsSheetMetal()
        except Exception as e:
            print(f"Body {index}: IsSheetMetal failed:", e)
            is_sheet_metal = False

        print(
            f"Body {index}: "
            f"{name} | "
            f"SheetMetal = {is_sheet_metal}"
        )

        if is_sheet_metal:
            result.append(body)

    print()
    print(f"Sheet Metal bodies: {len(result)}")

    return result


def rebuild(model):

    print()
    print("Rebuilding model...")

    try:
        result = model.EditRebuild3()

        print(
            f"EditRebuild3 result: {result}"
        )

    except Exception as e:
        print(
            f"Rebuild failed: {e}"
        )


def try_export(model, output_path):

    print()
    print("=" * 70)
    print("TRYING DWG EXPORT")
    print("=" * 70)

    try:
        model_path = model.GetPathName()
    except Exception:
        model_path = str(PART_PATH)

    print("Model:")
    print(model_path)

    print()
    print("Output:")
    print(output_path)

    # 12 عدد برای Alignment
    alignment = (
        0.0, 0.0, 0.0,
        1.0, 0.0, 0.0,
        0.0, 1.0, 0.0,
        0.0, 0.0, 1.0
    )

    print()
    print("Calling ExportToDWG2...")

    try:

        result = model.ExportToDWG2(
            str(output_path),
            str(model_path),
            SW_EXPORT_SHEET_METAL,
            True,
            alignment,
            False,
            False,
            SHEET_METAL_OPTIONS,
            None
        )

        print()
        print("ExportToDWG2 result:")
        print(result)

    except Exception as e:

        print()
        print("ExportToDWG2 EXCEPTION:")
        print(e)

        return False

    time.sleep(2)

    if output_path.exists():

        print()
        print("=" * 70)
        print("SUCCESS")
        print("=" * 70)

        print("DWG CREATED:")
        print(output_path)

        print(
            f"Size: {output_path.stat().st_size} bytes"
        )

        return True

    print()
    print("=" * 70)
    print("NO DWG FILE")
    print("=" * 70)

    print(
        "SolidWorks did not create the output file."
    )

    return False


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 70)
    print("SOLIDWORKS SINGLE PART DWG TEST")
    print("=" * 70)

    print()
    print("PART:")
    print(PART_PATH)

    print()
    print("OUTPUT:")
    print(OUTPUT_DIR)

    # --------------------------------------------------------
    # Check files
    # --------------------------------------------------------

    if not PART_PATH.exists():

        print()
        print("ERROR:")
        print("Part file does not exist.")

        return

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    output_path = (
        OUTPUT_DIR /
        f"{PART_PATH.stem}_TEST.dwg"
    )

    # حذف خروجی قبلی
    if output_path.exists():

        try:
            output_path.unlink()
            print()
            print("Old DWG deleted.")
        except Exception as e:
            print(
                "Could not delete old DWG:",
                e
            )

    # --------------------------------------------------------
    # Connect SolidWorks
    # --------------------------------------------------------

    print()
    print("Connecting to SolidWorks...")

    try:

        sw = win32com.client.Dispatch(
            "SldWorks.Application"
        )

        sw.Visible = True

        print("Connected.")

    except Exception as e:

        print()
        print("Could not connect to SolidWorks:")
        print(e)

        return

    # --------------------------------------------------------
    # Open part
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("OPENING PART")
    print("=" * 70)

    model = None

    try:

        # اول بررسی می‌کنیم شاید قبلاً باز باشد
        try:
            model = sw.GetOpenDocumentByName(
                str(PART_PATH)
            )
        except Exception:
            model = None

        # اگر باز نبود، بازش می‌کنیم
        if model is None:

            errors = 0
            warnings = 0

            model = sw.OpenDoc6(
                str(PART_PATH),
                SW_DOC_PART,
                0,
                "",
                errors,
                warnings
            )

        if model is None:

            print()
            print("FAILED TO OPEN PART.")

            return

        print()
        print("PART OPENED SUCCESSFULLY.")

    except Exception as e:

        print()
        print("OPEN ERROR:")
        print(e)

        return

    # --------------------------------------------------------
    # Model information
    # --------------------------------------------------------

    try:
        print()
        print("Actual model path:")
        print(model.GetPathName())
    except Exception:
        pass

    try:
        print()
        print("Model title:")
        print(model.GetTitle())
    except Exception:
        pass

    try:
        print()
        print("Model type:")
        print(model.GetType())
    except Exception:
        pass

    # --------------------------------------------------------
    # Rebuild
    # --------------------------------------------------------

    rebuild(model)

    # --------------------------------------------------------
    # Bodies
    # --------------------------------------------------------

    sheet_metal_bodies = get_sheet_metal_bodies(
        model
    )

    if not sheet_metal_bodies:

        print()
        print("THIS PART IS NOT SHEET METAL.")
        print()
        print("Stopping test.")

        return

    # --------------------------------------------------------
    # Feature Tree
    # --------------------------------------------------------

    features = scan_features(model)

    # --------------------------------------------------------
    # Flat Pattern
    # --------------------------------------------------------

    flat_patterns = find_flat_patterns(
        model
    )

    print()
    print("=" * 70)
    print("FLAT PATTERN RESULT")
    print("=" * 70)

    print(
        f"Flat Pattern candidates: "
        f"{len(flat_patterns)}"
    )

    for index, feature in enumerate(
        flat_patterns,
        start=1
    ):

        try:
            name = feature.Name
        except Exception:
            name = "?"

        try:
            type_name = feature.GetTypeName2
        except Exception:
            type_name = "?"

        print(
            f"{index}. "
            f"{name} | {type_name}"
        )

    # --------------------------------------------------------
    # Try export
    # --------------------------------------------------------

    success = try_export(
        model,
        output_path
    )

    # --------------------------------------------------------
    # Final
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("TEST FINISHED")
    print("=" * 70)

    if success:

        print()
        print("RESULT: SUCCESS")
        print()
        print("DWG:")
        print(output_path)

    else:

        print()
        print("RESULT: FAILED")
        print()
        print(
            "Feature Tree output above is important."
        )
        print(
            "Send me the complete console output."
        )


if __name__ == "__main__":

    pythoncom.CoInitialize()

    try:

        main()

    except KeyboardInterrupt:

        print()
        print("Interrupted by user.")

    except Exception as e:

        print()
        print("=" * 70)
        print("UNEXPECTED ERROR")
        print("=" * 70)
        print(e)

    finally:

        pythoncom.CoUninitialize()