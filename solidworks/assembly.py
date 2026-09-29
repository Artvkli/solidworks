from pathlib import Path

import pythoncom
import win32com.client

from .component import Component

SW_OPEN_SILENT = 1
SW_DOC_ASSEMBLY = 2


class SolidWorksAssembly:
    def __init__(self, sw_app):
        self.sw_app = sw_app
        self.model = None
        self.opened_by_us = False

    @staticmethod
    def _get(obj, name, default=None):
        if obj is None:
            return default
        try:
            value = getattr(obj, name)
            return value() if callable(value) else value
        except Exception:
            return default

    @staticmethod
    def _flag(obj, name):
        value = SolidWorksAssembly._get(obj, name, False)
        return bool(value) if isinstance(value, (bool, int)) else False

    def open(self, file_path):
        file_path = Path(file_path).resolve()
        if not file_path.exists() or file_path.suffix.lower() != ".sldasm":
            return False

        try:
            existing = self.sw_app.GetOpenDocumentByName(str(file_path))
            if existing is not None:
                self.model = existing
                self.opened_by_us = False
                return self._prepare()
        except Exception:
            pass

        try:
            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            self.model = self.sw_app.OpenDoc6(
                str(file_path), SW_DOC_ASSEMBLY, SW_OPEN_SILENT, "", errors, warnings
            )
            self.opened_by_us = self.model is not None
            print(f"OpenDoc6 assembly errors={getattr(errors, 'value', errors)} warnings={getattr(warnings, 'value', warnings)}")
            return self.model is not None and self._prepare()
        except Exception as exc:
            print(f"Assembly open failed: {exc}")
            return False

    def _prepare(self):
        try:
            self.model.ResolveAllLightweightComponents(False)
        except Exception as exc:
            print(f"Warning: could not resolve lightweight components: {exc}")
        return True

    def get_components(self):
        if self.model is None:
            return []
        try:
            raw = self.model.GetComponents(False)
        except Exception as exc:
            print(f"GetComponents failed: {exc}")
            return []
        components = []
        for sw_component in raw or []:
            try:
                if self._flag(sw_component, "IsSuppressed") or self._flag(sw_component, "IsEnvelope"):
                    continue
                path = self._get(sw_component, "GetPathName")
                if not path:
                    continue
                path = Path(str(path))
                if path.suffix.lower() == ".sldprt":
                    kind = "PART"
                elif path.suffix.lower() == ".sldasm":
                    kind = "ASSEMBLY"
                else:
                    continue
                components.append(
                    Component(
                        name=str(self._get(sw_component, "Name2", path.stem)),
                        path=path,
                        component_type=kind,
                        sw_component=sw_component,
                        lightweight=self._flag(sw_component, "IsLightWeight"),
                        configuration=self._get(sw_component, "ReferencedConfiguration"),
                    )
                )
            except Exception as exc:
                print(f"Warning: component skipped: {exc}")
        return components

    def get_unique_part_quantities(self, components=None):
        components = self.get_components() if components is None else components
        grouped = {}
        for component in components:
            if not component.is_part:
                continue
            path = Path(component.path).resolve()
            config = component.configuration or ""
            key = (str(path).lower(), str(config).lower())
            if key not in grouped:
                grouped[key] = {
                    "name": path.stem,
                    "path": path,
                    "quantity": 0,
                    "configuration": component.configuration,
                    "components": [],
                }
            grouped[key]["quantity"] += 1
            grouped[key]["components"].append(component)
        return list(grouped.values())


def find_assemblies(input_folder):
    folder = Path(input_folder)
    if not folder.exists():
        return []
    return sorted(
        [p for p in folder.iterdir() if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() == ".sldasm"],
        key=lambda p: p.name.lower(),
    )
