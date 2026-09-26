from pathlib import Path

import pythoncom
import win32com.client


class SolidWorksAssembly:

    SW_DOC_ASSEMBLY = 2
    SW_OPEN_DOC_OPTIONS_SILENT = 1

    def __init__(self, connection):
        self.connection = connection
        self.document = None
        self.source_path: Path | None = None

    def open(self, assembly_path: Path):

        assembly_path = assembly_path.resolve()

        if not assembly_path.exists():
            raise FileNotFoundError(
                f"Assembly does not exist: {assembly_path}"
            )

        if assembly_path.suffix.lower() != ".sldasm":
            raise ValueError(
                f"Not a SOLIDWORKS assembly: {assembly_path}"
            )

        sw = self.connection.app

        if sw is None:
            raise RuntimeError(
                "SOLIDWORKS is not connected."
            )

        # Make the assembly directory the current working directory.
        try:
            sw.SetCurrentWorkingDirectory(
                str(assembly_path.parent)
            )
        except Exception:
            pass

        # OpenDoc6 uses ByRef error/warning outputs.
        errors = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0,
        )

        warnings = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0,
        )

        self.document = sw.OpenDoc6(
            str(assembly_path),
            self.SW_DOC_ASSEMBLY,
            self.SW_OPEN_DOC_OPTIONS_SILENT,
            "",
            errors,
            warnings,
        )

        if self.document is None:
            raise RuntimeError(
                "SOLIDWORKS failed to open assembly: "
                f"{assembly_path}\n"
                f"Open errors: {errors.value}\n"
                f"Open warnings: {warnings.value}"
            )

        self.source_path = assembly_path

        return self.document

    def resolve_components(self):

        if self.document is None:
            raise RuntimeError(
                "No assembly is open."
            )

        try:
            self.document.ResolveAllLightWeightComponents(
                False
            )
        except Exception:
            pass

    def get_components(self):

        if self.document is None:
            raise RuntimeError(
                "No assembly is open."
            )

        # False = include components inside sub-assemblies.
        components = self.document.GetComponents(False)

        if components is None:
            return []

        return list(components)

    def close(self):

        if self.document is None:
            return

        sw = self.connection.app

        if sw is not None:
            try:
                title = self.document.GetTitle

                if callable(title):
                    title = title()

                sw.CloseDoc(str(title))

            except Exception:
                pass

        self.document = None
        self.source_path = None