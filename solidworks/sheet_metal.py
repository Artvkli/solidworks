from pathlib import Path
import logging

import pythoncom
import win32com.client


logger = logging.getLogger(__name__)


class SheetMetalProcessor:

    SW_DOC_PART = 1
    SW_OPEN_DOC_OPTIONS_SILENT = 1

    SHEET_METAL_FEATURE_TYPES = {
        "sheetmetal",
        "smbaseflange",
        "sm-base-flange",
        "sm3d",
        "baseflange",
    }

    def __init__(self, connection):
        self.connection = connection

    # =========================================================
    # Generic COM helper
    # =========================================================

    @staticmethod
    def _get_value(obj, name, default=None):

        try:
            value = getattr(obj, name)

            if callable(value):
                value = value()

            return value

        except Exception:
            return default

    # =========================================================
    # Component information
    # =========================================================

    def get_component_name(self, component):

        value = self._get_value(
            component,
            "Name2",
            "Unknown",
        )

        return str(value)

    def get_component_path(self, component):

        try:
            value = getattr(
                component,
                "GetPathName",
            )

            if callable(value):
                value = value()

            if value:
                return Path(str(value))

        except Exception:
            pass

        return None

    def get_configuration(self, component):

        value = self._get_value(
            component,
            "ReferencedConfiguration",
            "",
        )

        if value is None:
            return ""

        return str(value)

    # =========================================================
    # Get model from component
    # =========================================================

    def get_component_model(self, component):

        name = self.get_component_name(
            component
        )

        # -----------------------------------------------------
        # First attempt:
        # IComponent2.GetModelDoc2
        # -----------------------------------------------------

        try:

            value = getattr(
                component,
                "GetModelDoc2",
            )

            if callable(value):
                value = value()

            if value is not None:

                logger.debug(
                    "GetModelDoc2 succeeded: %s",
                    name,
                )

                return value, False

        except Exception as exc:

            logger.warning(
                "GetModelDoc2 failed for %s: %s",
                name,
                exc,
            )

        # -----------------------------------------------------
        # Fallback:
        # Open the SLDPRT directly from its path
        # -----------------------------------------------------

        path = self.get_component_path(
            component
        )

        if path is None:

            logger.warning(
                "Cannot open component because path is unavailable: %s",
                name,
            )

            return None, False

        if not path.exists():

            logger.warning(
                "Component file does not exist: %s",
                path,
            )

            return None, False

        logger.info(
            "Opening part directly: %s",
            path,
        )

        model = self.open_part(
            path
        )

        if model is None:

            logger.warning(
                "Could not open part: %s",
                path,
            )

            return None, False

        return model, True

    # =========================================================
    # Open Part
    # =========================================================

    def open_part(self, path: Path):

        sw = self.connection.app

        if sw is None:
            raise RuntimeError(
                "SOLIDWORKS is not connected."
            )

        path = path.resolve()

        # -----------------------------------------------------
        # Check if the document is already open
        # -----------------------------------------------------

        try:

            get_open = getattr(
                sw,
                "GetOpenDocumentByName",
            )

            if callable(get_open):

                existing = get_open(
                    str(path)
                )

            else:

                existing = None

            if existing is not None:

                logger.info(
                    "Part already open: %s",
                    path,
                )

                return existing

        except Exception:

            pass

        # -----------------------------------------------------
        # Open document
        # -----------------------------------------------------

        errors = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0,
        )

        warnings = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4,
            0,
        )

        model = sw.OpenDoc6(
            str(path),
            self.SW_DOC_PART,
            self.SW_OPEN_DOC_OPTIONS_SILENT,
            "",
            errors,
            warnings,
        )

        if model is None:

            logger.error(
                "OpenDoc6 failed for part: %s | errors=%s | warnings=%s",
                path,
                errors.value,
                warnings.value,
            )

            return None

        logger.info(
            "Part opened successfully: %s",
            path,
        )

        return model

    # =========================================================
    # Close directly opened part
    # =========================================================

    def close_part(self, model):

        if model is None:
            return

        try:

            title = getattr(
                model,
                "GetTitle",
            )

            if callable(title):
                title = title()

            if not title:
                return

            sw = self.connection.app

            if sw is not None:

                sw.CloseDoc(
                    str(title)
                )

                logger.debug(
                    "Part closed: %s",
                    title,
                )

        except Exception as exc:

            logger.warning(
                "Could not close part: %s",
                exc,
            )

    # =========================================================
    # Configuration
    # =========================================================

    def activate_configuration(
        self,
        model,
        configuration,
    ):

        if not configuration:
            return

        try:

            result = model.ShowConfiguration2(
                configuration
            )

            logger.debug(
                "ShowConfiguration2('%s') => %s",
                configuration,
                result,
            )

        except Exception as exc:

            logger.warning(
                "ShowConfiguration2 failed for '%s': %s",
                configuration,
                exc,
            )

        try:

            model.EditRebuild3()

        except Exception:

            pass

    # =========================================================
    # Feature tree
    # =========================================================

    def get_all_features(self, model):

        features = []

        try:

            first = getattr(
                model,
                "FirstFeature",
            )

            if callable(first):
                first = first()

        except Exception as exc:

            logger.error(
                "FirstFeature failed: %s",
                exc,
            )

            return features

        feature = first

        counter = 0

        while feature is not None:

            counter += 1

            if counter > 10000:

                logger.warning(
                    "Feature traversal safety limit reached."
                )

                break

            features.append(
                feature
            )

            try:

                next_feature = getattr(
                    feature,
                    "GetNextFeature",
                )

                if callable(next_feature):
                    next_feature = next_feature()

            except Exception:

                break

            feature = next_feature

        return features

    # =========================================================
    # Feature type
    # =========================================================

    def get_feature_type(self, feature):

        # GetTypeName2
        try:

            value = getattr(
                feature,
                "GetTypeName2",
            )

            if callable(value):
                value = value()

            if value:
                return str(value)

        except Exception:

            pass

        # GetTypeName
        try:

            value = getattr(
                feature,
                "GetTypeName",
            )

            if callable(value):
                value = value()

            if value:
                return str(value)

        except Exception:

            pass

        return ""

    # =========================================================
    # Feature name
    # =========================================================

    def get_feature_name(self, feature):

        value = self._get_value(
            feature,
            "Name",
            "Unknown",
        )

        return str(value)

    # =========================================================
    # Feature inspection
    # =========================================================

    def inspect_model(
        self,
        model,
        model_path,
    ):

        logger.info(
            "=============================================="
        )

        logger.info(
            "MODEL INSPECTION"
        )

        logger.info(
            "Path: %s",
            model_path,
        )

        logger.info(
            "Model type: %s",
            type(model),
        )

        features = self.get_all_features(
            model
        )

        logger.info(
            "Feature count: %d",
            len(features),
        )

        for index, feature in enumerate(
            features[:100],
            start=1,
        ):

            name = self.get_feature_name(
                feature
            )

            feature_type = self.get_feature_type(
                feature
            )

            logger.info(
                "FEATURE %03d | name=%s | type=%s",
                index,
                name,
                feature_type,
            )

        logger.info(
            "=============================================="
        )

        return features

    # =========================================================
    # Find Sheet Metal feature
    # =========================================================

    def find_sheet_metal_feature(
        self,
        model,
    ):

        features = self.get_all_features(
            model
        )

        for feature in features:

            feature_type = (
                self.get_feature_type(
                    feature
                )
            )

            normalized = (
                feature_type
                .strip()
                .lower()
            )

            if normalized in self.SHEET_METAL_FEATURE_TYPES:

                return feature

        return None

    # =========================================================
    # Thickness
    # =========================================================

    def get_thickness(
        self,
        feature,
    ):

        try:

            definition = getattr(
                feature,
                "GetDefinition",
            )

            if callable(definition):
                definition = definition()

            if definition is None:
                return None

            thickness = getattr(
                definition,
                "Thickness",
            )

            if callable(thickness):
                thickness = thickness()

            if thickness is None:
                return None

            return float(thickness)

        except Exception as exc:

            logger.warning(
                "Could not read thickness: %s",
                exc,
            )

            return None

    # =========================================================
    # Process one component
    # =========================================================

    def process_component(
        self,
        component,
        output_root,
    ):

        name = self.get_component_name(
            component
        )

        logger.info(
            "----------------------------------------------"
        )

        logger.info(
            "COMPONENT: %s",
            name,
        )

        # -----------------------------------------------------
        # Component path
        # -----------------------------------------------------

        path = self.get_component_path(
            component
        )

        if path is None:

            logger.warning(
                "No component path: %s",
                name,
            )

            return None

        logger.info(
            "Component path: %s",
            path,
        )

        # -----------------------------------------------------
        # Only SLDPRT
        # -----------------------------------------------------

        if path.suffix.lower() != ".sldprt":

            logger.info(
                "Skipping non-part component: %s",
                path,
            )

            return None

        # -----------------------------------------------------
        # Get model
        # -----------------------------------------------------

        model, opened_by_us = (
            self.get_component_model(
                component
            )
        )

        if model is None:

            logger.warning(
                "Could not get model: %s",
                path,
            )

            return None

        try:

            logger.info(
                "Model object: %s",
                type(model),
            )

            # -------------------------------------------------
            # Configuration
            # -------------------------------------------------

            configuration = (
                self.get_configuration(
                    component
                )
            )

            logger.info(
                "Configuration: %s",
                configuration,
            )

            self.activate_configuration(
                model,
                configuration,
            )

            # -------------------------------------------------
            # Inspect feature tree
            # -------------------------------------------------

            self.inspect_model(
                model,
                path,
            )

            # -------------------------------------------------
            # Detect Sheet Metal
            # -------------------------------------------------

            sheet_metal_feature = (
                self.find_sheet_metal_feature(
                    model
                )
            )

            if sheet_metal_feature is None:

                logger.info(
                    "NOT SHEET METAL: %s",
                    path.name,
                )

                return None

            logger.info(
                "SHEET METAL DETECTED: %s",
                path.name,
            )

            # -------------------------------------------------
            # Thickness
            # -------------------------------------------------

            thickness_m = (
                self.get_thickness(
                    sheet_metal_feature
                )
            )

            if thickness_m is None:

                logger.warning(
                    "Could not determine thickness: %s",
                    path.name,
                )

                return None

            thickness_mm = (
                thickness_m * 1000.0
            )

            logger.info(
                "Thickness: %.4f mm",
                thickness_mm,
            )

            return {
                "name": path.stem,
                "path": path,
                "configuration": configuration,
                "model": model,
                "thickness_m": thickness_m,
                "thickness_mm": thickness_mm,
                "feature": sheet_metal_feature,
                "opened_by_us": opened_by_us,
            }

        finally:

            # Only close documents that this processor opened.
            if opened_by_us:

                self.close_part(
                    model
                )