from pathlib import Path


class SheetMetalProcessor:

    SW_EXPORT_TO_DWG_SHEET_METAL = 1

    SHEET_METAL_FEATURE_TYPES = {
        "sheetmetal",
        "smbaseflange",
        "sm-base-flange",
        "sm3d",
        "baseflange",
    }

    def __init__(self, connection):
        self.connection = connection

    @staticmethod
    def safe_filename(name: str) -> str:

        invalid = '<>:"/\\|?*'

        for char in invalid:
            name = name.replace(char, "_")

        name = name.strip().rstrip(".")

        return name or "Part"

    def get_component_model(self, component):

        model = None

        try:
            value = component.GetModelDoc2

            if callable(value):
                model = value()
            else:
                model = value

        except Exception:
            pass

        return model

    def get_component_path(self, component):

        try:
            value = component.GetPathName

            if callable(value):
                value = value()

            if value:
                return Path(str(value))

        except Exception:
            pass

        return None

    def get_configuration(self, component) -> str:

        try:
            value = component.ReferencedConfiguration

            if callable(value):
                value = value()

            if value:
                return str(value)

        except Exception:
            pass

        return "Default"

    def get_feature_type(self, feature) -> str:

        try:
            value = feature.GetTypeName2

            if callable(value):
                value = value()

            if value:
                return str(value)

        except Exception:
            pass

        try:
            value = feature.GetTypeName

            if callable(value):
                value = value()

            if value:
                return str(value)

        except Exception:
            pass

        return ""

    def get_all_features(self, model):

        features = []

        # FirstFeature
        try:

            value = model.FirstFeature

            if callable(value):
                feature = value()
            else:
                feature = value

            while feature is not None:

                features.append(feature)

                try:
                    value = feature.GetNextFeature

                    if callable(value):
                        feature = value()
                    else:
                        feature = value

                except Exception:
                    break

            if features:
                return features

        except Exception:
            pass

        # Fallback: FeatureManager.GetFeatures
        try:

            feature_manager = model.FeatureManager

            value = feature_manager.GetFeatures

            if callable(value):
                result = value(True)
            else:
                result = value

            if result:
                return list(result)

        except Exception:
            pass

        return features

    def find_sheet_metal_feature(self, model):

        features = self.get_all_features(model)

        for feature in features:

            feature_type = (
                self.get_feature_type(feature)
                .strip()
                .lower()
            )

            if feature_type in self.SHEET_METAL_FEATURE_TYPES:
                return feature

        return None

    def is_sheet_metal(self, model) -> bool:

        return (
            self.find_sheet_metal_feature(model)
            is not None
        )

    def get_thickness(self, model):

        feature = self.find_sheet_metal_feature(model)

        if feature is None:
            return None

        # SheetMetal feature definition
        try:

            value = feature.GetDefinition

            if callable(value):
                definition = value()
            else:
                definition = value

            if definition is not None:

                value = definition.Thickness

                if callable(value):
                    value = value()

                if value is not None:
                    return float(value)

        except Exception:
            pass

        # Fallback: scan other SheetMetal-like
        # feature definitions.
        for candidate in self.get_all_features(model):

            feature_type = (
                self.get_feature_type(candidate)
                .strip()
                .lower()
            )

            if feature_type not in self.SHEET_METAL_FEATURE_TYPES:
                continue

            try:

                value = candidate.GetDefinition

                if callable(value):
                    definition = value()
                else:
                    definition = value

                if definition is None:
                    continue

                thickness = definition.Thickness

                if callable(thickness):
                    thickness = thickness()

                if thickness is not None:
                    return float(thickness)

            except Exception:
                continue

        return None

    def activate_configuration(
        self,
        model,
        configuration: str,
    ) -> None:

        try:

            value = model.ShowConfiguration2

            if callable(value):
                result = value(configuration)
            else:
                result = value

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

            value = model.EditRebuild3

            if callable(value):
                value()

        except Exception:
            pass

    def process_component(
        self,
        component,
        output_root: Path,
        options: int,
    ):

        path = self.get_component_path(component)

        if path is None:
            return None

        if path.suffix.lower() != ".sldprt":
            return None

        model = self.get_component_model(component)

        if model is None:
            return None

        configuration = self.get_configuration(
            component
        )

        self.activate_configuration(
            model,
            configuration,
        )

        if not self.is_sheet_metal(model):
            return None

        thickness = self.get_thickness(model)

        name = self.safe_filename(path.stem)

        if thickness is None:
            thickness_folder = "unknown_thickness"
        else:
            thickness_folder = f"{thickness:g}mm"

        output_dir = output_root / thickness_folder

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

        # Alignment matrix
        alignment = (
            0.0, 0.0, 0.0,
            1.0, 0.0, 0.0,
            0.0, 1.0, 0.0,
            0.0, 0.0, 1.0,
        )

        export_to_dwg = model.ExportToDWG2

        success = export_to_dwg(
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

        if not success:
            raise RuntimeError(
                "ExportToDWG2 returned False."
            )

        if not output_path.exists():
            raise RuntimeError(
                "SOLIDWORKS reported successful export "
                "but the DXF file was not created: "
                f"{output_path}"
            )

        return output_path