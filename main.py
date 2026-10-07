from pathlib import Path
import csv
import gc
import json
import re
import sys
import traceback

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is required. Install: pip install pywin32")
    raise

from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.dwg_exporter import DwgExporter
from solidworks.pdf_builder import PdfDrawingExporter
from solidworks.sw_drawings import SwDrawingPdfExporter
from solidworks.thickness_dwg import ThicknessDwgBuilder
from solidworks.sheet_metal import SheetMetalDetector

GROUP_BY_THICKNESS = True
REPEAT_BY_QUANTITY = False
ODA_CONVERTER_PATH = None
PART_GAP = 20.0
MAX_ROW_WIDTH = 6000.0

# ---- PDF drawings (one PDF per sheet metal part, one folder per thickness) ----
EXPORT_PDF = True
# "solidworks": real SolidWorks drawings (default template) with folded views + flat pattern + bend table.
#               If one fails, the simple Python drawing below is used for that part.
# "python":     only the simple Python drawing (flat pattern + overall size), no SolidWorks drawing.
PDF_MODE = "solidworks"
SW_DRAWING_TEMPLATE = (
    None  # None = SolidWorks default drawing template; or r"C:\...\Drawing.drwdot"
)
PDF_FOLDED_VIEWS = True  # views next to the flat pattern
# Views after the flat pattern, in this order: "Front", "Top", "Right", "Isometric" (the bent part in 3D)
PDF_VIEWS = ("Front", "Isometric")
PDF_BEND_TABLE = True
PDF_DIMENSIONS = True  # model dimensions imported into the views
# FIRST RUN: only the first N parts get a PDF, so problems show up quickly.
# Set to None to export all parts.
PDF_TEST_LIMIT = 3
# ---- TEST SHORTCUT ----------------------------------------------------------
# True  -> skip the scan and the DXF export. Use the DXF files that already exist in
#          output/.../_parts_dxf (plus the scan CSV) and only make the PDFs.
# False -> run the whole pipeline (scan -> DXF -> merged drawings -> PDFs).
PDF_TEST_FROM_EXISTING_DXF = True
# Only these parts get a PDF (names as in the log, e.g. ["AHU3500-09", "AHU3500-D01-01"]).
# Empty list = use PDF_TEST_LIMIT below instead.
PDF_TEST_ONLY = []
# -----------------------------------------------------------------------------
# The settings below are used by the simple Python drawing (and as its fallback):
PDF_PAPER = "A4"  # "A4" or "A3"
PDF_ORIENTATION = "portrait"  # "portrait", "landscape" or "auto" (auto = whichever gives the bigger scale)
PDF_COMPANY = "TECHNICAL DEPARTMENT"
PDF_DRAWN_BY = "HE.A"
PDF_CHECKED_BY = "R.GH"
PDF_APPROVED_BY = "R.GH"
PDF_MATERIAL = ""


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


def save_run_summary(
    output_folder, assembly_path, scan_result, drawing_result, pdf_result=None
):
    data = {
        "assembly": str(assembly_path),
        "scan": {
            "unique_parts": scan_result.get("unique_parts", 0),
            "sheet_metal": scan_result.get("sheet_metal", 0),
            "not_sheet_metal": scan_result.get("not_sheet_metal", 0),
            "failed": scan_result.get("failed", 0),
            "total_quantity": scan_result.get("total_quantity", 0),
            "failed_parts": scan_result.get("failed_parts", []),
        },
        "drawings": {
            "created": len(drawing_result.get("files", [])) if drawing_result else 0,
            "failed_parts": drawing_result.get("failed_parts", [])
            if drawing_result
            else [],
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
        "pdf": {
            "created": pdf_result.get("created", 0) if pdf_result else 0,
            "failed": pdf_result.get("failed", 0) if pdf_result else 0,
            "failed_parts": pdf_result.get("failed_parts", []) if pdf_result else [],
        },
    }
    path = Path(output_folder) / "run_summary.json"
    try:
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"Run summary saved: {path}")
    except Exception as exc:
        print(f"Could not save run summary: {exc}")


def _find_existing_dxf_folder(output_folder):
    """The _parts_dxf folder written by an earlier full run."""
    preferred = "by_thickness" if GROUP_BY_THICKNESS else "combined"
    for sub in (preferred, "by_thickness", "combined"):
        folder = Path(output_folder) / sub / "_parts_dxf"
        if folder.exists() and any(folder.glob("*.dxf")):
            return folder
    return None


def _read_scan_csv(csv_path):
    """Rows of the scan CSV by safe part name (thickness, quantities, mirrors)."""
    rows = {}
    if not Path(csv_path).exists():
        return rows
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            name = (row.get("Part") or "").strip()
            if name:
                rows[ThicknessDwgBuilder._safe(name)] = row
    return rows


def _number(value, cast=float):
    try:
        text = str(value).strip()
        return cast(text) if text else None
    except Exception:
        return None


def load_test_items(output_folder, assembly_path, unique_parts):
    """
    Rebuilds the items of an earlier run from files on disk:
      - flat pattern DXFs  (output/.../_parts_dxf/*.dxf)
      - scan CSV           (thickness, quantity, mirrors)
      - the open assembly  (part path + configuration, needed for SolidWorks drawings)
    Returns [{"item": {...}, "dxf": Path}, ...]
    """
    dxf_folder = _find_existing_dxf_folder(output_folder)
    if dxf_folder is None:
        print(
            "\nNo DXF files found in output/.../_parts_dxf. Run once with PDF_TEST_FROM_EXISTING_DXF = False."
        )
        return []

    csv_path = Path(output_folder) / f"{Path(assembly_path).stem}_sheet_metal.csv"
    rows = _read_scan_csv(csv_path)
    print(f"\nDXF folder: {dxf_folder}")
    print(
        f"Scan CSV:   {csv_path} ({'found, ' + str(len(rows)) + ' parts' if rows else 'NOT FOUND - thickness unknown'})"
    )

    parts_by_name = {}
    for part in unique_parts:
        parts_by_name.setdefault(ThicknessDwgBuilder._safe(part["name"]), []).append(
            part
        )

    entries = []
    for dxf in sorted(dxf_folder.glob("*.dxf"), key=lambda p: p.name.lower()):
        key = dxf.stem
        if key not in rows and key not in parts_by_name:
            key = re.sub(
                r"_\d+$", "", key
            )  # duplicate names get a _2 suffix when exported

        row = rows.get(key)
        candidates = parts_by_name.get(key, [])
        part = None
        if candidates:
            wanted_config = (row or {}).get("Configuration") or ""
            part = next(
                (
                    c
                    for c in candidates
                    if (c.get("configuration") or "") == wanted_config
                ),
                candidates[0],
            )

        if row is None and part is None:
            print(f"  skipped (no data for it): {dxf.name}")
            continue

        quantity = _number((row or {}).get("Total quantity"), int)
        if quantity is None:
            quantity = int(part["quantity"]) if part else 1

        item = {
            "name": (row or {}).get("Part") or part["name"],
            "path": str(part["path"]) if part else "",
            "configuration": part["configuration"]
            if part
            else ((row or {}).get("Configuration") or None),
            "quantity": quantity,
            "own_quantity": _number((row or {}).get("Own quantity"), int) or quantity,
            "mirror_quantity": _number((row or {}).get("Mirror quantity"), int) or 0,
            "mirror_parts": [
                x for x in ((row or {}).get("Mirror parts") or "").split("; ") if x
            ],
            "thickness": _number((row or {}).get("Thickness (mm)")),
            "thickness_source": (row or {}).get("Thickness source") or "",
            "thickness_raw_mm": _number((row or {}).get("Raw value (mm)")),
        }
        entries.append({"item": item, "dxf": dxf})

    if PDF_TEST_ONLY:
        wanted = {ThicknessDwgBuilder._safe(n).lower() for n in PDF_TEST_ONLY}
        entries = [
            e
            for e in entries
            if ThicknessDwgBuilder._safe(e["item"]["name"]).lower() in wanted
        ]
        print(f"PDF_TEST_ONLY: {len(entries)} of {len(wanted)} requested parts found")

    print(f"Parts ready for PDF: {len(entries)}")
    return entries


def run_pdf_test(sw_app, assembly_path, unique_parts, output_folder):
    """Test shortcut: PDFs only, from the DXFs of an earlier run."""
    print("\n" + "=" * 60)
    print("PDF TEST MODE (existing DXFs, no scan, no DXF export)")
    print("=" * 60)

    entries = load_test_items(output_folder, assembly_path, unique_parts)
    if not entries:
        return 1

    python_pdf = PdfDrawingExporter(
        paper=PDF_PAPER,
        orientation=PDF_ORIENTATION,
        company=PDF_COMPANY,
        drawn_by=PDF_DRAWN_BY,
        checked_by=PDF_CHECKED_BY,
        approved_by=PDF_APPROVED_BY,
        material=PDF_MATERIAL,
    )

    limit = None if PDF_TEST_ONLY else PDF_TEST_LIMIT

    if PDF_MODE == "solidworks" and sw_app is not None:
        sw_pdf = SwDrawingPdfExporter(
            sw_app,
            assembly_path=assembly_path,
            reconnect=connect_to_solidworks,
            template=SW_DRAWING_TEMPLATE,
            folded_views=PDF_FOLDED_VIEWS,
            view_names=PDF_VIEWS,
            bend_table=PDF_BEND_TABLE,
            dimensions=PDF_DIMENSIONS,
        )
        try:
            result = sw_pdf.export_all(
                entries, Path(output_folder) / "pdf", fallback=python_pdf, limit=limit
            )
        finally:
            sw_pdf.activate_assembly()
    else:
        if limit:
            entries = sorted(
                entries,
                key=lambda e: (e["item"].get("thickness") or 0, e["item"]["name"]),
            )[: int(limit)]
        result = python_pdf.export_all(entries, Path(output_folder) / "pdf")

    print("\n" + "=" * 60)
    print("PDF TEST FINISHED")
    print("=" * 60)
    print(
        f"PDFs created: {result.get('created', 0)}   failed: {result.get('failed', 0)}"
    )
    print(f"Folder: {Path(output_folder) / 'pdf'}")
    return 0 if result.get("failed", 0) == 0 else 1


def main():
    project_folder = Path(__file__).resolve().parent
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"
    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)

    print("\n" + "=" * 60)
    print("SOLIDWORKS SHEET METAL -> DXF/DWG/PDF EXPORTER")
    print("=" * 60)

    if PDF_TEST_FROM_EXISTING_DXF and PDF_MODE != "solidworks":
        # simple Python drawings need neither SolidWorks nor the assembly
        assembly_path = select_assembly(input_folder)
        if assembly_path is None:
            return 1
        return run_pdf_test(None, assembly_path, [], output_folder)

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

    components = assembly.get_components()
    if not components:
        print("No components found. Stopping.")
        return 1

    unique_parts = assembly.get_unique_part_quantities(components)
    print_summary(components, unique_parts)

    # ---- test shortcut: everything below this block (scan, DXF export, merged drawings) is skipped ----
    if PDF_TEST_FROM_EXISTING_DXF:
        return run_pdf_test(sw_app, assembly_path, unique_parts, output_folder)

    detector = SheetMetalDetector(
        sw_app, assembly_path=assembly_path, read_thickness=GROUP_BY_THICKNESS
    )
    scan_result = detector.scan_and_export(unique_parts, output_folder)

    drawing_result = None
    exporter = DwgExporter(
        sw_app,
        assembly_path=assembly_path,
        reconnect=connect_to_solidworks,
    )
    builder = ThicknessDwgBuilder(
        exporter,
        gap=PART_GAP,
        max_row_width=MAX_ROW_WIDTH,
        repeat_by_quantity=REPEAT_BY_QUANTITY,
        oda_path=ODA_CONVERTER_PATH,
        group_by_thickness=GROUP_BY_THICKNESS,
    )

    try:
        drawing_result = builder.build(
            scan_result["sheet_metal_parts"],
            output_folder / ("by_thickness" if GROUP_BY_THICKNESS else "combined"),
            base_name=assembly_path.stem,
        )
    except Exception as exc:
        print("\nDRAWING EXPORT FAILED")
        print(f"{type(exc).__name__}: {exc}")
        traceback.print_exc()
        drawing_result = {"files": [], "failed_parts": [], "error": str(exc)}
    finally:
        exporter.activate_assembly()
        gc.collect()
        try:
            pythoncom.CoFreeUnusedLibraries()
        except Exception:
            pass

    # ---- PDF drawings: built from the flat pattern DXFs exported above ----
    pdf_result = None
    if EXPORT_PDF:
        exported_parts = (drawing_result or {}).get("exported_parts", [])
        if not exported_parts:
            print("\nNo flat pattern DXFs were exported, so no PDF drawings were made.")
        else:
            python_pdf = PdfDrawingExporter(
                paper=PDF_PAPER,
                orientation=PDF_ORIENTATION,
                company=PDF_COMPANY,
                drawn_by=PDF_DRAWN_BY,
                checked_by=PDF_CHECKED_BY,
                approved_by=PDF_APPROVED_BY,
                material=PDF_MATERIAL,
            )
            try:
                if PDF_MODE == "solidworks":
                    sw_pdf = SwDrawingPdfExporter(
                        sw_app,
                        assembly_path=assembly_path,
                        reconnect=connect_to_solidworks,
                        template=SW_DRAWING_TEMPLATE,
                        folded_views=PDF_FOLDED_VIEWS,
                        view_names=PDF_VIEWS,
                        bend_table=PDF_BEND_TABLE,
                        dimensions=PDF_DIMENSIONS,
                    )
                    try:
                        pdf_result = sw_pdf.export_all(
                            exported_parts,
                            output_folder / "pdf",
                            fallback=python_pdf,
                            limit=PDF_TEST_LIMIT,
                        )
                    finally:
                        sw_pdf.activate_assembly()
                else:
                    pdf_result = python_pdf.export_all(
                        exported_parts, output_folder / "pdf"
                    )
            except Exception as exc:
                print("\nPDF EXPORT FAILED")
                print(f"{type(exc).__name__}: {exc}")
                traceback.print_exc()

    save_run_summary(
        output_folder, assembly_path, scan_result, drawing_result, pdf_result
    )

    print("\n" + "=" * 60)
    print("PROCESS FINISHED")
    print("=" * 60)
    print(f"Sheet metal (unique): {scan_result['sheet_metal']}")
    print(f"Sheet metal (quantity): {scan_result['total_quantity']}")
    print(f"Not sheet metal: {scan_result['not_sheet_metal']}")
    print(f"Scan failures: {scan_result['failed']}")
    if drawing_result:
        print(f"Drawings created: {len(drawing_result['files'])}")
        print(f"Export failures: {len(drawing_result['failed_parts'])}")
    if pdf_result:
        print(
            f"PDF drawings created: {pdf_result['created']} (failures: {pdf_result['failed']})"
        )
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
