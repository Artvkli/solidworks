from pathlib import Path

import pythoncom
import win32com.client

from .component import Component

SW_OPEN_SILENT = 1
SW_DOC_ASSEMBLY = 2
SW_COMPONENT_HIDDEN = 0  # swComponentHidden (swComponentVisibilityState_e)


class SolidWorksAssembly:
    def __init__(self, sw_app, skip_hidden=True):
        self.sw_app = sw_app
        self.model = None
        self.opened_by_us = False
        self.skip_hidden = skip_hidden  

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
    def _obj(obj, name):
        """Read a property/method that returns a COM object (or None)."""
        if obj is None:
            return None
        try:
            value = getattr(obj, name)
            if hasattr(value, "_oleobj_"):  # already evaluated COM object
                return value
            return value() if callable(value) else value
        except Exception:
            return None

    @staticmethod
    def _flag(obj, name):
        value = SolidWorksAssembly._get(obj, name, False)
        return bool(value) if isinstance(value, (bool, int)) else False

    def _own_hidden_state(self, sw_component):
        """
        True  -> this component is hidden
        False -> this component is visible
        None  -> SolidWorks did not let us read its visibility
        Two independent signals are used: the Visible property and IsHidden().
        """
        signals = []

        visible = self._get(sw_component, "Visible")
        if isinstance(visible, int) and not isinstance(visible, bool):
            signals.append(visible == SW_COMPONENT_HIDDEN)

        try:
            hidden = sw_component.IsHidden(False)
            if isinstance(hidden, (bool, int)):
                signals.append(bool(hidden))
        except Exception:
            pass

        if not signals:
            return None
        return any(signals)

    def _hidden_state(self, sw_component):
        """
        Returns (is_hidden, readable).
        A component is hidden if it is hidden itself OR sits inside a hidden parent assembly.
        """
        own = self._own_hidden_state(sw_component)
        readable = own is not None
        if own:
            return True, True

        current = self._obj(sw_component, "GetParent")
        for _ in range(100):  # guard against endless parent chains
            if current is None:
                break
            if self._own_hidden_state(current):
                return True, readable
            current = self._obj(current, "GetParent")

        return False, readable

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
            print(
                f"OpenDoc6 assembly errors={getattr(errors, 'value', errors)} warnings={getattr(warnings, 'value', warnings)}"
            )
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
        skipped_hidden = []
        unreadable_visibility = 0
        for sw_component in raw or []:
            try:
                if self._flag(sw_component, "IsSuppressed") or self._flag(
                    sw_component, "IsEnvelope"
                ):
                    continue
                if self.skip_hidden:
                    hidden, readable = self._hidden_state(sw_component)
                    if not readable:
                        unreadable_visibility += 1
                    if hidden:
                        skipped_hidden.append(
                            str(self._get(sw_component, "Name2", "?"))
                        )
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
                        configuration=self._get(
                            sw_component, "ReferencedConfiguration"
                        ),
                    )
                )
            except Exception as exc:
                print(f"Warning: component skipped: {exc}")
        if self.skip_hidden:
            print(f"Hidden components skipped: {len(skipped_hidden)}")
            for name in skipped_hidden[:60]:
                print(f"  - {name}")
            if len(skipped_hidden) > 60:
                print(f"  ... and {len(skipped_hidden) - 60} more")
            if unreadable_visibility:
                print(
                    f"WARNING: visibility could not be read for {unreadable_visibility} "
                    f"component(s) - they were treated as VISIBLE. Check the list above."
                )
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
        [
            p
            for p in folder.iterdir()
            if p.is_file()
            and not p.name.startswith("~$")
            and p.suffix.lower() == ".sldasm"
        ],
        key=lambda p: p.name.lower(),
    )
