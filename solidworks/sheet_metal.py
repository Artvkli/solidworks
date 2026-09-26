from pathlib import Path

import win32com.client


class SheetMetalProcessor:

    SW_EXPORT_TO_DWG_SHEET_METAL = 1

    def __init__(self, connection):

        self.connection = connection

    @staticmethod
    def safe_filename(name: str) -> str:

        invalid = '<>:"/\\|?*'

        for char in invalid:
            name = name.replace(char, "_")

        name = name.strip().rstrip(".")

        if not name:
            return "Part"

        return name

    def get_component_model(self, component):

        try:
            model = component.GetModelDoc2
        except Exception:
            model = None

        if model is None:
            try:
                model = component.GetModelDoc2()
            except Exception:
                model = None

        return model

    def get_component_path(self, component) -> Path | None:

        try:
            path = component.GetPathName()
        except Exception:
            return None

        if not path:
            return None

        return Path(path)

    def get_configuration(self, component) -> str:

        try:
            configuration = (
                component.ReferencedConfiguration
            )

            if configuration:
                return str(configuration)

        except Exception:
            pass

        return "Default"

    def get_feature_type(self, feature) -> str:

        for method_name in (
            "GetTypeName2",
            "GetTypeName",
        ):

            try:

                method = getattr(
                    feature,
                    method_name,
                )

                if callable(method):
                    value = method()
                else:
                    value = method

                if value:
                    return str(value)

            except Exception:
                pass

        return ""

    def find_sheet_metal_feature(self, model):

        try:
            feature = model.FirstFeature()
        except Exception:
            return None

        while feature is not None:

            feature_type = (
                self.get_feature_type(feature)
                .lower()
            )

            if feature_type in (
                "sheetmetal",
                "sm3d",
                "sm-base-flange",
                "smbaseflange",
                "baseflange",
            ):
                return feature

            try:
                feature = feature.GetNextFeature()
            except Exception:
                break

        return None

    def get_thickness(self, model):

        feature = self.find_sheet_metal_feature(
            model
        )

        if feature is None:
            return None

        try:
            definition = feature.GetDefinition()

            thickness = definition.Thickness

            if thickness is not None:
                return float(thickness)

        except Exception:
            pass

        return None

    def is_sheet_metal(self, model) -> bool:

        feature = self.find_sheet_metal_feature(
            model
        )

        return feature is not None

    def activate_configuration(
        self,
        model,
        configuration: str,
    ) -> None:

        try:
            result = model.ShowConfiguration2(
                configuration
            )

            if result is False:
                raise RuntimeError(
                    f"Could not activate configuration: "
                    f"{configuration}"
                )

        except Exception as exc:

            raise RuntimeError(
                f"Could not activate configuration "
                f"'{configuration}': {exc}"
            ) from exc

        try:
            model.EditRebuild3()
        except Exception:
            pass

    def export_dxf(
        self,
        model,
        model_path: Path,
        output_path: Path,
        options: int,
    ) -> Path:

        output_path = output_path.resolve()

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        # Default alignment.
        alignment = (
            0.0, 0.0, 0.0,
            1.0, 0.0, 0.0,
            0.0, 1.0, 0.0,
            0.0, 0.0, 1.0,
        )

        part = model

        # Newer API signature:
        #
        # ExportToDWG2(
        #   FilePath,
        #   ModelName,
        #   Action,
        #   ExportToSingleFile,
        #   Alignment,
        #   IsXDirFlipped,
        #   IsYDirFlipped,
        #   SheetMetalOptions,
        #   ViewsCount,
        #   Views
        # )
        #
        # Older COM signatures omit ViewsCount.
        try:

            success = part.ExportToDWG2(
                str(output_path),
                str(model_path),
                self.SW_EXPORT_TO_DWG_SHEET_METAL,
                True,
                alignment,
                False,
                False,
                options,
                0,
                None,
            )

        except Exception as first_error:

            try:

                success = part.ExportToDWG2(
                    str(output_path),
                    str(model_path),
                    self.SW_EXPORT_TO_DWG_SHEET_METAL,
                    True,
                    alignment,
                    False,
                    False,
                    options,
                    None,
                )

            except Exception as second_error:

                raise RuntimeError(
                    "SOLIDWORKS ExportToDWG2 failed. "
                    f"New signature: {first_error}. "
                    f"Legacy signature: {second_error}."
                ) from second_error

        if not success:
            raise RuntimeError(
                "SOLIDWORKS ExportToDWG2 returned False."
            )

        if not output_path.exists():

            raise RuntimeError(
                "SOLIDWORKS reported a successful DXF "
                "export, but the file was not created: "
                f"{output_path}"
            )

        return output_path

    def process_component(
        self,
        component,
        output_root: Path,
        options: int,
    ):

        path = self.get_component_path(
            component
        )

        if path is None:
            return None

        if path.suffix.lower() != ".sldprt":
            return None

        model = self.get_component_model(
            component
        )

        if model is None:
            return None

        configuration = (
            self.get_configuration(component)
        )

        self.activate_configuration(
            model,
            configuration,
        )

        if not self.is_sheet_metal(model):
            return None

        thickness = self.get_thickness(
            model
        )

        name = self.safe_filename(
            path.stem
        )

        if thickness is None:

            thickness_folder = (
                "unknown_thickness"
            )

        else:

            thickness_folder = (
                f"{thickness:g}mm"
            )

        output_dir = (
            output_root / thickness_folder
        )

        output_path = (
            output_dir / f"{name}.dxf"
        )

        return {
            "name": name,
            "path": path,
            "configuration": configuration,
            "thickness": thickness,
            "model": model,
            "output": output_path,
        }