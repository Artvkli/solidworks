SOLIDWORKS SHEET METAL -> DXF/DWG EXPORTER

1) Install Python packages:
   pip install -r req.txt
   (pywin32 is required and is Windows-only; it is listed in req.txt)

2) Put the .SLDASM and its referenced SolidWorks files in:
   input\

3) Close unnecessary SolidWorks documents before running.

4) Run:
   python main.py

Outputs are written to:
   output\

Important:
- The scanner does NOT keep Part COM objects alive between scan and export.
- Each export opens one Part, activates its configuration, rebuilds it, exports with ExportToDWG2, then closes it if the script opened it.
- If ExportToDWG2 fails, the log contains the actual COM/API error and a Flat-Pattern diagnostic instead of silently replacing the error with "Flat-Pattern feature not found".
- A disconnected SolidWorks COM session is retried once after reconnecting.
- run_summary.json contains the final scan/export summary.
- ODA File Converter is optional; without it, the grouped outputs remain DXF.

UPDATE - FLAT PATTERN DIMENSIONS
--------------------------------
The native SolidWorks drawing exporter now attempts to create editable drawing
dimensions on the Flat Pattern view: overall extents plus each bend line to the
nearest outer straight edge. These are created through the SolidWorks drawing
API (selected bend sketch segments + visible outline edges), before the .SLDDRW
is saved and the PDF is exported from that same drawing.

Important: this code must be run on Windows with SolidWorks installed. The current
execution environment cannot run SolidWorks COM, so the new dimensions have been
syntax-checked but not verified against a live part/drawing. Review the console log
for the line starting with "flat dimensions:"; it reports how many native dimensions
SolidWorks actually created. If the count is zero, the selected visible-edge entity
type differs in your SolidWorks version and the API call needs a version-specific fix.

UPDATE - BEND DIMENSIONS (PHASE 2)
----------------------------------
Bend-line-to-outer-edge dimensions were rewritten (see MAINTENANCE_NOTES.md):
values are measured in the bend sketch's own frame, holes/cut-outs are never used as reference,
equal values on different bends are all kept, every dimension is read back and verified, and the
layout reserves room below and to the right of the flat pattern. Folded views: Top, Front, Right.

First run: main.py has SLDDRW_LIMIT = 1 -> one part, written to output\flat_drawings_test\
(production drawings are not touched). Check the log and <part>_bend_probe.json, open the drawing,
then set SLDDRW_LIMIT = None for all parts.

Offline tests (no SolidWorks needed): pip install pytest, then  python -m pytest tests
