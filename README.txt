SOLIDWORKS SHEET METAL -> DXF/DWG EXPORTER

1) Install Python packages:
   pip install -r requirements.txt

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
