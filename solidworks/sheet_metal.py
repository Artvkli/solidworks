from pathlib import Path
import logging


logger = logging.getLogger(__name__)


class SheetMetalProcessor:

    def get_component_model(self, component):
        try:
            value = component.GetModelDoc2
            return value() if callable(value) else value
        except Exception as exc:
            logger.exception(
                "GetModelDoc2 failed for component: %s",
                self.get_component_name(component),
            )
            return None

    def get_component_path(self, component):
        try:
            value = component.GetPathName
            value = value() if callable(value) else value

            if value:
                return Path(str(value))

        except Exception:
            pass

        return None

    def get_component_name(self, component):
        try:
            value = component.Name2
            return value() if callable(value) else str(value)
        except Exception:
            return "Unknown"

    def get_configuration(self, component):
        try:
            value = component.ReferencedConfiguration
            return value() if callable(value) else str(value)
        except Exception:
            return ""

    def get_all_features(self, model):
        features = []

        try:
            first = model.FirstFeature
            first = first() if callable(first) else first
        except Exception as exc:
            logger.exception("FirstFeature failed")
            return features

        feature = first

        while feature is not None:
            features.append(feature)

            try:
                next_feature = feature.GetNextFeature
                next_feature = (
                    next_feature()
                    if callable(next_feature)
                    else next_feature
                )
            except Exception:
                break

            feature = next_feature

        return features

    def get_feature_type(self, feature):
        try:
            value = feature.GetTypeName2
            return value() if callable(value) else str(value)
        except Exception:
            pass

        try:
            value = feature.GetTypeName
            return value() if callable(value) else str(value)
        except Exception:
            return ""

    def get_feature_name(self, feature):
        try:
            value = feature.Name
            return value() if callable(value) else str(value)
        except Exception:
            return "Unknown"

    def inspect_model(self, model, model_path):
        logger.info("==============================================")
        logger.info("MODEL INSPECTION")
        logger.info("Path: %s", model_path)
        logger.info("Model type: %s", type(model))

        features = self.get_all_features(model)

        logger.info("Feature count: %d", len(features))

        for index, feature in enumerate(features[:50], start=1):
            feature_type = self.get_feature_type(feature)
            feature_name = self.get_feature_name(feature)

            logger.info(
                "FEATURE %02d | name=%s | type=%s",
                index,
                feature_name,
                feature_type,
            )

        logger.info("==============================================")

        return features

    def find_sheet_metal_feature(self, model):
        features = self.get_all_features(model)

        for feature in features:
            feature_type = self.get_feature_type(feature)

            normalized = feature_type.strip().lower()

            if normalized in {
                "sheetmetal",
                "smbaseflange",
                "sm-base-flange",
                "sm3d",
                "baseflange",
            }:
                return feature

        return None

    def get_thickness(self, feature):
        try:
            definition = feature.GetDefinition

            if callable(definition):
                definition = definition()

            if definition is None:
                return None

            value = definition.Thickness

            if callable(value):
                value = value()

            if value is None:
                return None

            return float(value)

        except Exception as exc:
            logger.warning(
                "Could not read thickness: %s",
                exc,
            )
            return None

    def is_sheet_metal(self, model):
        feature = self.find_sheet_metal_feature(model)
        return feature is not None

    def process_component(self, component, output_root):
        name = self.get_component_name(component)

        logger.info("----------------------------------------------")
        logger.info("COMPONENT: %s", name)

        path = self.get_component_path(component)

        if path is None:
            logger.warning(
                "No component path: %s",
                name,
            )
            return None

        logger.info("Component path: %s", path)

        if path.suffix.lower() != ".sldprt":
            logger.info(
                "Skipping non-part component: %s",
                path,
            )
            return None

        model = self.get_component_model(component)

        if model is None:
            logger.warning(
                "GetModelDoc2 returned None: %s",
                path,
            )
            return None

        logger.info(
            "Model object: %s",
            type(model),
        )

        configuration = self.get_configuration(component)

        logger.info(
            "Configuration: %s",
            configuration,
        )

        if configuration:
            try:
                model.ShowConfiguration2(configuration)
            except Exception as exc:
                logger.warning(
                    "ShowConfiguration2 failed: %s",
                    exc,
                )

        try:
            model.EditRebuild3()
        except Exception:
            pass

        # IMPORTANT:
        # Print the actual feature tree before deciding
        # whether this is sheet metal.
        self.inspect_model(model, path)

        sheet_feature = self.find_sheet_metal_feature(model)

        if sheet_feature is None:
            logger.info(
                "NOT SHEET METAL: %s",
                path.name,
            )
            return None

        logger.info(
            "SHEET METAL DETECTED: %s",
            path.name,
        )

        thickness = self.get_thickness(sheet_feature)

        if thickness is None:
            logger.warning(
                "Could not determine thickness: %s",
                path.name,
            )
            return None

        logger.info(
            "Thickness: %.4f mm",
            thickness * 1000.0,
        )

        return {
            "name": path.stem,
            "path": path,
            "configuration": configuration,
            "model": model,
            "thickness_m": thickness,
        }