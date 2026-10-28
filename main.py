from pathlib import Path
import gc
import json
import sys
import traceback

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is required. Install: pip install pywin32")
    raise

from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.sw_drawings import SwDrawingExporter
from solidworks.dwg_exporter import DwgExporter
from solidworks.thickness_dwg import ThicknessDwgBuilder
from solidworks.sheet_metal import SheetMetalDetector

GROUP_BY_THICKNESS = True       # DXF/DWG: one file per thickness | SLDDRW/PDF: one folder per thickness
REPEAT_BY_QUANTITY = False
ODA_CONVERTER_PATH = None
PART_GAP = 20.0
MAX_ROW_WIDTH = 6000.0
USE_SCAN_CACHE = True

# ---- output mode (chosen in the terminal at start; can also be given as: python main.py 1|2|3) ----
#   label, make SLDDRW + PDF (bending), make DXF / DWG (laser)
MODES = {
    "1": ("DXF / DWG for laser cutting", False, True),
    "2": ("SLDDRW + PDF for bending", True, False),
    "3": ("Both (DXF / DWG + SLDDRW / PDF)", True, True),
}

# ---- flat pattern SLDDRW ----
EXPORT_PDF = True               # a PDF next to every SLDDRW
MASTER_DRAWING_NAME = "master.slddrw"
SHEET_FORMAT_NAME = None
SHEET_TEXT_REPLACEMENTS = {
    "{PART_NAME}": "PartName",
    "{QTY}": "Quantity",
    "{THICKNESS}": "Thickness",
    "{MATERIAL}": "Material",
}

ADD_FOLDED_VIEWS = True  # folded model views above the flat pattern
# Folded views and their arrangement: Top above Front, Right to the right of Front.
# Use ("*Top", "*Front") for the earlier two-view layout.
FOLDED_VIEWS = ("*Top", "*Front", "*Right")
LAYOUT = {
    "margin_mm": 10,
    "side_pad_mm": 0,
    "top_pad_mm": 5,
    "table_gap_mm": 0,
    "table_clearance_mm": 9,
    "folded_band_ratio": 0.38,
    "band_gap_mm": 2,
    "view_gap_mm": 12,
    "dim_pad_mm": 12,
    "common_scale": False,
    "bend_dim_pad_mm": 18,
    "auto_rotate_flat": True,
    "scale_mode": "fine",
    "max_scale": 2.0,
    "folded_weight": 1.5,
    "folded_ratio": 0.6,         # بیشتر = Top/Front/Right بزرگ‌تر، فلت کوچک‌تر
    "folded_arrange": "auto",    # "standard" = همیشه همان چیدمان قدیمی
    "folded_rotate": True,       # False = هیچ نمایی چرخانده نشود
}

# Bend dimensions: bend line -> nearest parallel edge of the OUTER sheet boundary
# (holes / cut-outs are never used as reference; bends without a reliable reference are reported).
ADD_BEND_DIMENSIONS = True
# Stacking order when dimensions overlap:
#   "largest_outside"  = smallest values nearest the part, largest farthest out (standard; same as old "largest_first")
#   "smallest_outside" = the reverse (same as old "smallest_first")
BEND_DIM_ORDER = "largest_outside"
BEND_DIM_OFFSET_MM = 8          # gap between the part and the first row of bend dimensions
BEND_DIM_STEP_MM = 7            # distance between stacked rows
BEND_DIM_TEXT_WIDTH_MM = 14     # space reserved for one dimension text
BEND_DIM_TEXT_HEIGHT_MM = 3.5
BEND_DIM_RIGHT_PAD_MM = 18      # initial room right of the flat pattern (below: LAYOUT["bend_dim_pad_mm"]);
                                # both adapt automatically to what the dimensions need
BEND_DIM_MERGE_COLLINEAR = True  # collinear bend lines measured to the same edge share one dimension
BEND_DIM_VERIFY = "flag"        # each dimension is read back from SolidWorks; "flag" = report a mismatch,
                                # "delete" = also remove the mismatching dimension
BEND_DEBUG_JSON = True          # write <part>_bend_probe.json next to each drawing (bends, edges, frames, tags)
ADD_QTY_NOTE = True
QTY_NOTE_X_MM = None
QTY_NOTE_Y_MM = None

# Overall dimensions for Flat, Top and Front views.
ADD_DIMENSIONS = True
DIMENSION_OFFSET_MM = 8

INSERT_BEND_TABLE = True
BEND_TABLE_START_TAG = "A"
BEND_TABLE_TEMPLATE_NAME = None
BEND_TABLE_RIGHT_MARGIN_MM = 10
BEND_TABLE_BOTTOM_MM = 66
BEND_TABLE_REMOVE_WORDS = ("factor", "allowance", "deduction", "length", "thickness")
# Controlled test: only the first N sheet-metal parts (scan order) and/or the parts named here.
# While either is set, drawings go to output/flat_drawings_test so production drawings are not overwritten.
# Production run: SLDDRW_LIMIT = None and SLDDRW_ONLY_PARTS = None.
SLDDRW_LIMIT = None  # e.g. 5 - only the first N sheet-metal parts (scan order)
SLDDRW_ONLY_PARTS = None  # e.g. ("AHU3500-08",) - part names as printed in the scan
SLDDRW_TEST_FOLDER = "flat_drawings_test"


def choose_mode():
    """Terminal menu. Returns (label, make_slddrw, make_dxf) or None to quit."""
    arg = sys.argv[1].strip() if len(sys.argv) > 1 else None
    if arg in MODES:
        print(f"\nMode {arg}: {MODES[arg][0]} (from the command line)")
        return MODES[arg]

    print("\nWhat do you want to export?")
    for key, (label, _slddrw, _dxf) in MODES.items():
        print(f"  {key}. {label}")
    print("  0. Exit")
    while True:
        try:
            choice = input("Select 1, 2 or 3: ").strip()
        except EOFError:
            return None
        if choice in MODES:
            print(f"Selected: {MODES[choice][0]}")
            return MODES[choice]
        if choice.lower() in ("0", "q", "exit"):
            return None
        print("Invalid choice - type 1, 2, 3 (or 0 to exit).")


def connect_to_solidworks():
    print("\nConnecting to SolidWorks...")
    try:
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
        print("Connected successfully.")
        return sw_app
    except Exception as exc:
        print(f"Could not connect to SolidWorks: {exc}")
        return None


def select_assembly(input_folder):
    assemblies = find_assemblies(input_folder)
    if not assemblies:
        print(f"\nNo .SLDASM files found in:\n{input_folder}")
        return None
    print("\nAssemblies found:")
    for i, assembly in enumerate(assemblies, 1):
        print(f"  {i}. {assembly.name}")
    selected = assemblies[0]
    print(f"Selected: {selected}")
    return selected


def print_summary(components, unique_parts):
    parts = sum(1 for c in components if c.is_part)
    assemblies = sum(1 for c in components if c.is_assembly)
    print("\n" + "=" * 60)
    print("ASSEMBLY SUMMARY")
    print("=" * 60)
    print(f"Total components: {len(components)}")
    print(f"Part instances:   {parts}")
    print(f"Sub-assemblies:   {assemblies}")
    print(f"Unique part/config: {len(unique_parts)}")


def save_run_summary(output_folder, assembly_path, mode_label, scan_result, drawing_result, slddrw_result):
    data = {
        "assembly": str(assembly_path),
        "mode": mode_label,
        "scan": {
            "unique_parts": scan_result.get("unique_parts", 0),
            "sheet_metal": scan_result.get("sheet_metal", 0),
            "not_sheet_metal": scan_result.get("not_sheet_metal", 0),
            "failed": scan_result.get("failed", 0),
            "total_quantity": scan_result.get("total_quantity", 0),
            "failed_parts": scan_result.get("failed_parts", []),
        },
        "slddrw": {
            "created": len(slddrw_result["created"]) if slddrw_result else 0,
            "pdf_failed": slddrw_result.get("pdf_failed", []) if slddrw_result else [],
            "failed": slddrw_result["failed"] if slddrw_result else [],
        },
        "drawings": {
            "created": len(drawing_result.get("files", [])) if drawing_result else 0,
            "failed_parts": drawing_result.get("failed_parts", []) if drawing_result else [],
            "error": drawing_result.get("error") if drawing_result else None,
            "files": [
                {
                    "thickness": item.get("thickness"),
                    "unique_parts": item.get("unique_parts"),
                    "quantity": item.get("quantity"),
                    "path": str(item.get("path")),
                    "is_dwg": item.get("is_dwg", False),
                }
                for item in (drawing_result.get("files", []) if drawing_result else [])
            ],
        },
    }
    path = Path(output_folder) / "run_summary.json"
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Run summary saved: {path}")
    except Exception as exc:
        print(f"Could not save run summary: {exc}")


SCAN_CACHE_NAME = "scan_cache.json"
SCAN_CACHE_FIELDS = (
    "unique_parts", "sheet_metal", "not_sheet_metal", "failed", "failed_parts",
    "total_quantity", "sheet_metal_parts", "mirror_links",
)


def _scan_cache_key(assembly_path):
    try:
        mtime = Path(assembly_path).stat().st_mtime
    except OSError:
        mtime = None
    return {"assembly": str(assembly_path), "mtime": mtime, "group_by_thickness": GROUP_BY_THICKNESS}


def save_scan_cache(output_folder, assembly_path, scan_result):
    data = {
        "key": _scan_cache_key(assembly_path),
        "result": {k: scan_result.get(k) for k in SCAN_CACHE_FIELDS},
    }
    path = Path(output_folder) / SCAN_CACHE_NAME
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
        print(f"Scan saved for the next run: {path}")
    except Exception as exc:
        print(f"Could not save the scan cache: {exc}")


def load_scan_cache(output_folder, assembly_path):
    path = Path(output_folder) / SCAN_CACHE_NAME
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("key") != _scan_cache_key(assembly_path):
            print("The saved scan belongs to another assembly or other settings - scanning again.")
            return None
        result = data["result"]
        if not result.get("sheet_metal_parts"):
            return None
        result["thickness_groups"] = {}
        result["csv_path"] = None
        return result
    except Exception as exc:
        print(f"Could not read the scan cache ({exc}) - scanning again.")
        return None


def export_slddrw_files(sw_app, assembly_path, scan_result, output_folder, input_folder):
    """Copy master.slddrw for each part, add views/dimensions/table, save SLDDRW + PDF (one folder per thickness)."""
    master_path = input_folder / MASTER_DRAWING_NAME
    if not master_path.is_file():
        print(f"\nMASTER DRAWING NOT FOUND: {master_path}")
        print("Put your master.slddrw (sheet format + title block) in the input folder and run again.")
        return None

    exporter = SwDrawingExporter(
        sw_app,
        assembly_path=assembly_path,
        reconnect=connect_to_solidworks,
        master_drawing=master_path,
        sheet_format_path=(input_folder / SHEET_FORMAT_NAME) if SHEET_FORMAT_NAME else None,
        text_replacements=SHEET_TEXT_REPLACEMENTS,
        layout=LAYOUT,
        add_folded_views=ADD_FOLDED_VIEWS,
        add_qty_note=ADD_QTY_NOTE,
        qty_note_x_mm=QTY_NOTE_X_MM,
        qty_note_y_mm=QTY_NOTE_Y_MM,
        insert_bend_table=INSERT_BEND_TABLE,
        bend_table_template=(input_folder / BEND_TABLE_TEMPLATE_NAME) if BEND_TABLE_TEMPLATE_NAME else None,
        bend_table_right_margin_mm=BEND_TABLE_RIGHT_MARGIN_MM,
        bend_table_bottom_mm=BEND_TABLE_BOTTOM_MM,
        bend_table_remove_words=BEND_TABLE_REMOVE_WORDS,
        bend_table_start_value=BEND_TABLE_START_TAG,
        add_bend_dimensions=ADD_BEND_DIMENSIONS,
        bend_dim_offset_mm=BEND_DIM_OFFSET_MM,
        bend_dim_step_mm=BEND_DIM_STEP_MM,
        bend_dim_text_width_mm=BEND_DIM_TEXT_WIDTH_MM,
        bend_dim_order=BEND_DIM_ORDER,
        add_dimensions=ADD_DIMENSIONS,
        dimension_offset_mm=DIMENSION_OFFSET_MM,
        folded_view_names=FOLDED_VIEWS,
        bend_dim_right_pad_mm=BEND_DIM_RIGHT_PAD_MM,
        bend_dim_text_height_mm=BEND_DIM_TEXT_HEIGHT_MM,
        bend_dim_merge_collinear=BEND_DIM_MERGE_COLLINEAR,
        bend_dim_verify=BEND_DIM_VERIFY,
        bend_debug_json=BEND_DEBUG_JSON,
        group_by_thickness=GROUP_BY_THICKNESS,
        export_pdf=EXPORT_PDF,
    )
    items = scan_result["sheet_metal_parts"]
    test_mode = bool(SLDDRW_LIMIT or SLDDRW_ONLY_PARTS)
    if SLDDRW_ONLY_PARTS:
        wanted = {str(n).strip().lower() for n in SLDDRW_ONLY_PARTS}
        items = [i for i in items if str(i.get("name", "")).strip().lower() in wanted]
        if not items:
            print(f"\nNone of SLDDRW_ONLY_PARTS {list(SLDDRW_ONLY_PARTS)} is a scanned sheet-metal part.")
    if SLDDRW_LIMIT:
        items = items[:int(SLDDRW_LIMIT)]

    folder = output_folder / (SLDDRW_TEST_FOLDER if test_mode else "خروجی خمکاری")
    folder.mkdir(parents=True, exist_ok=True)
    results = {"created": [], "pdf_failed": [], "failed": []}

    print("\n" + "=" * 60)
    print(f"BENDING DRAWINGS: SLDDRW{' + PDF' if EXPORT_PDF else ''} ({len(items)} parts)"
          + (f" - TEST MODE -> {folder.name}" if test_mode else ""))
    print("=" * 60)

    used_names = set()
    for index, item in enumerate(items, 1):
        stem = exporter._safe_name(item.get("name") or Path(item["path"]).stem)
        base, n = stem, 2
        while stem.lower() in used_names:
            stem = f"{base}_{n}"
            n += 1
        used_names.add(stem.lower())

        slddrw_path = folder / f"{stem}.SLDDRW"
        ok, message, steps = exporter.export_part(item, slddrw_path)
        record = {"part": item.get("name"), "message": message, "steps": steps}
        if ok:
            record.update({"slddrw": item.get("slddrw_path") or str(slddrw_path),
                           "pdf": item.get("pdf_path"),
                           "thickness": item.get("thickness"),
                           "quantity": item.get("quantity", 1)})
            results["created"].append(record)
            pdf_ok = message == "ok"
            if EXPORT_PDF and not pdf_ok:
                results["pdf_failed"].append(item.get("name"))
            where = Path(record["slddrw"]).parent.name
            print(f"[{index}/{len(items)}] OK: {item.get('name')} -> {where}"
                  + ("" if pdf_ok else "  (PDF NOT saved)"))
        else:
            results["failed"].append(record)
            print(f"[{index}/{len(items)}] FAILED: {item.get('name')} -> {message}")
        for line in steps:
            print(f"      {line}")

    summary = folder / "slddrw_summary.json"
    summary.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nSLDDRW: {len(results['created'])} created, {len(results['failed'])} failed"
          + (f" | PDF missing for {len(results['pdf_failed'])}" if results["pdf_failed"] else ""))
    print(f"Summary: {summary}")
    return results


def export_laser_files(sw_app, assembly_path, scan_result, output_folder):
    """DXF export, merge per thickness and DWG conversion (laser cutting)."""
    print("\n" + "=" * 60)
    print("LASER FILES: DXF / DWG")
    print("=" * 60)
    exporter = DwgExporter(sw_app, assembly_path=assembly_path, reconnect=connect_to_solidworks)
    try:
        builder = ThicknessDwgBuilder(
            exporter,
            gap=PART_GAP,
            max_row_width=MAX_ROW_WIDTH,
            repeat_by_quantity=REPEAT_BY_QUANTITY,
            oda_path=ODA_CONVERTER_PATH,
            group_by_thickness=GROUP_BY_THICKNESS,
        )
        return builder.build(
            scan_result["sheet_metal_parts"],
            output_folder / ("خروجی برش لیزر" if GROUP_BY_THICKNESS else "خروجی برش لیزر"),
            base_name=assembly_path.stem,
        )
    except Exception as exc:
        print("\nDXF / DWG EXPORT FAILED")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()
        return {"files": [], "failed_parts": [], "error": str(exc)}
    finally:
        try:
            exporter.activate_assembly()
        except Exception:
            pass


def main():
    project_folder = Path(__file__).resolve().parent
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"
    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)

    print("\n" + "=" * 60)
    print("SOLIDWORKS SHEET METAL -> LASER (DXF / DWG) / BENDING (SLDDRW / PDF)")
    print("=" * 60)

    mode = choose_mode()
    if mode is None:
        print("Nothing selected. Exit.")
        return 0
    mode_label, make_slddrw, make_dxf = mode

    sw_app = connect_to_solidworks()
    if sw_app is None:
        return 1

    assembly_path = select_assembly(input_folder)
    if assembly_path is None:
        return 1

    assembly = SolidWorksAssembly(sw_app)
    if not assembly.open(assembly_path):
        print("Could not open assembly. Stopping.")
        return 1

    scan_result = load_scan_cache(output_folder, assembly_path) if USE_SCAN_CACHE else None
    if scan_result is not None:
        print("\n" + "=" * 60)
        print("USING THE SAVED SCAN (no parts are scanned again)")
        print("=" * 60)
        print(f"Sheet metal parts: {scan_result['sheet_metal']}  |  total quantity: {scan_result['total_quantity']}")
        print(f"Delete {Path(output_folder) / SCAN_CACHE_NAME} (or set USE_SCAN_CACHE = False) to scan again.")
    else:
        components = assembly.get_components()
        if not components:
            print("No components found. Stopping.")
            return 1
        unique_parts = assembly.get_unique_part_quantities(components)
        print_summary(components, unique_parts)
        detector = SheetMetalDetector(sw_app, assembly_path=assembly_path, read_thickness=GROUP_BY_THICKNESS)
        scan_result = detector.scan_and_export(unique_parts, output_folder)
        if USE_SCAN_CACHE:
            save_scan_cache(output_folder, assembly_path, scan_result)

    slddrw_result = None
    drawing_result = None
    try:
        if make_dxf:
            drawing_result = export_laser_files(sw_app, assembly_path, scan_result, output_folder)
        if make_slddrw:
            try:
                slddrw_result = export_slddrw_files(sw_app, assembly_path, scan_result, output_folder, input_folder)
            except Exception as exc:
                print(f"\nSLDDRW EXPORT FAILED (continuing): {type(exc).__name__}: {exc}")
                traceback.print_exc()
    finally:
        if not make_dxf:  # the laser export re-activates the assembly itself
            try:
                DwgExporter(sw_app, assembly_path=assembly_path, reconnect=connect_to_solidworks).activate_assembly()
            except Exception:
                pass
        gc.collect()
        try:
            pythoncom.CoFreeUnusedLibraries()
        except Exception:
            pass

    save_run_summary(output_folder, assembly_path, mode_label, scan_result, drawing_result, slddrw_result)
    print("\n" + "=" * 60)
    print(f"PROCESS FINISHED - {mode_label}")
    print("=" * 60)
    print(f"Sheet metal (unique): {scan_result['sheet_metal']}")
    print(f"Sheet metal (quantity): {scan_result['total_quantity']}")
    print(f"Not sheet metal: {scan_result['not_sheet_metal']}")
    print(f"Scan failures: {scan_result['failed']}")
    if make_dxf and drawing_result:
        print(f"Laser files (DXF/DWG) created: {len(drawing_result['files'])}  |  "
              f"export failures: {len(drawing_result['failed_parts'])}")
    if make_slddrw and slddrw_result:
        print(f"Bending drawings (SLDDRW) created: {len(slddrw_result['created'])}  |  "
              f"failed: {len(slddrw_result['failed'])}"
              + (f"  |  PDF missing: {len(slddrw_result['pdf_failed'])}" if slddrw_result.get("pdf_failed") else ""))
    print(f"Output: {output_folder}")
    return 0


if __name__ == "__main__":
    pythoncom.CoInitialize()
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(130)
    except Exception as exc:
        print("\nUNEXPECTED ERROR")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()
        sys.exit(1)
    finally:
        pythoncom.CoUninitialize()