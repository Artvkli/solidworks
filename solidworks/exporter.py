from pathlib import Path

from config import DWG_SAVE_OPTIONS, DWG_SAVE_VERSION
from core.models import ConversionResult


class DrawingExporter:

    def export(self, document, output_path: Path) -> ConversionResult:

        output_path = output_path.resolve()

        try:
            output_path.parent.mkdir(
                parents=True,
                exist_ok=True
            )

            extension = document.Extension

            success = extension.SaveAs(
                str(output_path),
                DWG_SAVE_VERSION,
                DWG_SAVE_OPTIONS,
                None,
                None,
                0,
                0
            )

            if not success:
                return ConversionResult(
                    source=Path(document.GetPathName()),
                    output=output_path,
                    success=False,
                    message="SOLIDWORKS SaveAs returned False."
                )

            if not output_path.exists():
                return ConversionResult(
                    source=Path(document.GetPathName()),
                    output=output_path,
                    success=False,
                    message=(
                        "SOLIDWORKS reported success, "
                        "but the DWG file was not created."
                    )
                )

            return ConversionResult(
                source=Path(document.GetPathName()),
                output=output_path,
                success=True,
                message="DWG export completed successfully."
            )

        except Exception as exc:
            return ConversionResult(
                source=Path(document.GetPathName()),
                output=output_path,
                success=False,
                message=f"DWG export failed: {exc}"
            )