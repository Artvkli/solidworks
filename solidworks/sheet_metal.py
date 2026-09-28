# sheet_metal.py

from pathlib import Path
import win32com.client


SW_DOC_PART = 1
SW_SOLID_BODY = 0

SW_EXPORT_SHEET_METAL = 1

# Flat geometry = 1
# Bend lines   = 4
# Total        = 5
SW_SHEET_METAL_OPTIONS = 5


class SheetMetalDetector:
    def __init__(self, sw_app, assembly_path=None):

        self.sw_app = sw_app

        self.assembly_path = Path(assembly_path) if assembly_path else None

        self._opened_models = {}

    # =========================================================
    # SAFE COM READ
    # =========================================================

    @staticmethod
    def _read(obj, name, default=None):

        if obj is None:
            return default

        try:
            value = getattr(obj, name)
        except Exception:
            return default

        try:
            if callable(value):
                return value()

            return value

        except Exception:
            return default

    # =========================================================
    # GET MODEL FROM COMPONENT
    # =========================================================

    def get_model_from_component(self, component):

        sw_component = component.sw_component

        # -----------------------------------------------------
        # 1. Best option:
        #    GetModelDoc2()
        # -----------------------------------------------------

        if sw_component is not None:
            try:
                get_model = getattr(sw_component, "GetModelDoc2", None)

                if callable(get_model):
                    model = get_model()

                    if model is not None:
                        return model

            except Exception:
                pass

        # -----------------------------------------------------
        # 2. Try original path
        # -----------------------------------------------------

        original_path = Path(component.path)

        if original_path.exists():
            model = self.open_part(original_path)

            if model is not None:
                return model

        # -----------------------------------------------------
        # 3. Search next to assembly
        # -----------------------------------------------------

        if self.assembly_path:
            assembly_folder = self.assembly_path.parent

            filename = original_path.name

            candidate = assembly_folder / filename

            if candidate.exists():
                print(f"Using local part:\n  {candidate}")

                model = self.open_part(candidate)

                if model is not None:
                    return model

        # -----------------------------------------------------
        # 4. Recursive search inside assembly folder
        # -----------------------------------------------------

        if self.assembly_path:
            assembly_folder = self.assembly_path.parent

            filename = original_path.name

            print(f"Searching for:\n  {filename}")

            try:
                matches = list(assembly_folder.rglob(filename))

            except Exception:
                matches = []

            if matches:
                candidate = matches[0]

                print(f"Found:\n  {candidate}")

                model = self.open_part(candidate)

                if model is not None:
                    return model

        return None

    # =========================================================
    # OPEN PART
    # =========================================================

    def open_part(self, part_path):

        part_path = Path(part_path)

        if not part_path.exists():
            print(f"Part does not exist: {part_path}")

            return None

        key = str(part_path).lower()

        if key in self._opened_models:
            return self._opened_models[key]

        # Already open in SolidWorks
        try:
            model = self.sw_app.GetOpenDocumentByName(str(part_path))

            if model is not None:
                self._opened_models[key] = model

                return model

        except Exception:
            pass

        # Open part
        try:
            errors = 0
            warnings = 0

            model = self.sw_app.OpenDoc6(
                str(part_path), SW_DOC_PART, 1, "", errors, warnings
            )

            if model is not None:
                self._opened_models[key] = model

                return model

        except Exception as e:
            print(f"Could not open part:\n  {part_path}\n  Error: {e}")

        return None

    # =========================================================
    # GET BODIES
    # =========================================================

    def get_bodies(self, model):

        try:
            bodies = model.GetBodies2(SW_SOLID_BODY, False)

            if bodies is None:
                return []

            return list(bodies)

        except Exception as e:
            print(f"GetBodies2 error: {e}")

            return []

    # =========================================================
    # GET SHEET METAL BODIES
    # =========================================================

    def get_sheet_metal_bodies(self, model):

        result = []

        bodies = self.get_bodies(model)

        for body in bodies:
            try:
                is_sheet = body.IsSheetMetal()

                if is_sheet:
                    result.append(body)

            except Exception:
                pass

        return result

    # =========================================================
    # FEATURES
    # =========================================================

    def iter_features(self, model):

        try:
            feature = model.FirstFeature

            while feature is not None:
                yield feature

                try:
                    feature = feature.GetNextFeature()
                except Exception:
                    break

        except Exception:
            return

    # =========================================================
    # FIND FEATURE
    # =========================================================

    def find_feature(self, model, type_name):

        wanted = type_name.lower()

        for feature in self.iter_features(model):
            try:
                feature_type = feature.GetTypeName2()

                if feature_type and feature_type.lower() == wanted:
                    return feature

            except Exception:
                pass

        return None

    # =========================================================
    # SHEET METAL FEATURE
    # =========================================================

    def get_sheet_metal_feature(self, model):

        return self.find_feature(model, "SheetMetal")

    # =========================================================
    # IS SHEET METAL
    # =========================================================

    def is_sheet_metal(self, component):

        if not component.is_part:
            return False

        model = self.get_model_from_component(component)

        if model is None:
            print("Could not obtain ModelDoc2")

            return False

        # -----------------------------------------------------
        # Method 1: Body.IsSheetMetal()
        # -----------------------------------------------------

        sheet_bodies = self.get_sheet_metal_bodies(model)

        if sheet_bodies:
            return True

        # -----------------------------------------------------
        # Method 2: SheetMetal feature
        # -----------------------------------------------------

        feature = self.get_sheet_metal_feature(model)

        if feature is not None:
            return True

        return False

    # =========================================================
    # THICKNESS
    # =========================================================

    def get_thickness(self, component):

        model = self.get_model_from_component(component)

        if model is None:
            return None

        feature = self.get_sheet_metal_feature(model)

        if feature is None:
            return None

        try:
            definition = feature.GetDefinition()

            if definition is None:
                return None

            thickness = getattr(definition, "Thickness", None)

            if thickness is not None:
                try:
                    if callable(thickness):
                        thickness = thickness()

                except Exception:
                    pass

            return thickness

        except Exception:
            return None

    # =========================================================
    # FLAT PATTERN
    # =========================================================

    def get_flat_pattern(self, model):

        return self.find_feature(model, "FlatPattern")

    # =========================================================
    # EXPORT DWG
    # =========================================================

    def export_dwg(self, component, output_path):

        model = self.get_model_from_component(component)

        if model is None:
            print("Could not obtain model for export.")

            return False

        sheet_bodies = self.get_sheet_metal_bodies(model)

        if not sheet_bodies:
            print("No Sheet Metal body found.")

            return False

        # -----------------------------------------------------
        # Get actual model path
        # -----------------------------------------------------

        model_path = self._read(model, "GetPathName", "")

        if not model_path:
            model_path = component.path

        # -----------------------------------------------------
        # Export
        # -----------------------------------------------------

        try:
            output_path = Path(output_path)

            output_path.parent.mkdir(parents=True, exist_ok=True)

            result = model.ExportToDWG2(
                str(output_path),
                str(model_path),
                SW_EXPORT_SHEET_METAL,
                True,
                None,
                False,
                False,
                SW_SHEET_METAL_OPTIONS,
                None,
            )

            print(f"Export result: {result}")

            return bool(result)

        except Exception as e:
            print(f"Export DWG error:\n  {e}")

            return False

    # =========================================================
    # SAFE OUTPUT NAME
    # =========================================================

    @staticmethod
    def unique_output_path(output_folder, part_name):

        output_folder = Path(output_folder)

        output_folder.mkdir(parents=True, exist_ok=True)

        base = output_folder / f"{part_name}.dwg"

        if not base.exists():
            return base

        counter = 1

        while True:
            candidate = output_folder / f"{part_name}_{counter}.dwg"

            if not candidate.exists():
                return candidate

            counter += 1

    # =========================================================
    # SCAN AND EXPORT
    # =========================================================

    def scan_and_export(self, components, unique_parts, output_folder):

        output_folder = Path(output_folder)

        output_folder.mkdir(parents=True, exist_ok=True)

        exported = 0
        sheet_metal = 0
        non_sheet_metal = 0
        failed = 0

        print()
        print("=" * 60)
        print("SHEET METAL SCAN")
        print("=" * 60)

        total = len(unique_parts)

        for index, part in enumerate(unique_parts, start=1):
            print()
            print("-" * 60)
            print(f"[{index}/{total}]")
            print(f"Part: {part['name']}")
            print(f"Quantity: {part['quantity']}")

            component = part["components"][0]

            try:
                # -------------------------------------------------
                # Show original path
                # -------------------------------------------------

                print(f"Referenced path:\n  {component.path}")

                # -------------------------------------------------
                # Find model
                # -------------------------------------------------

                model = self.get_model_from_component(component)

                if model is None:
                    print("Result: FAILED - model not found")

                    failed += 1
                    continue

                # -------------------------------------------------
                # Check Sheet Metal
                # -------------------------------------------------

                is_sheet = self.is_sheet_metal(component)

                if not is_sheet:
                    print("Result: NOT SHEET METAL")

                    non_sheet_metal += 1
                    continue

                # -------------------------------------------------
                # Sheet Metal
                # -------------------------------------------------

                sheet_metal += 1

                print("Result: SHEET METAL")

                # -------------------------------------------------
                # Thickness
                # -------------------------------------------------

                thickness = self.get_thickness(component)

                if thickness is not None:
                    try:
                        print(f"Thickness: {float(thickness) * 1000:.3f} mm")

                    except Exception:
                        print(f"Thickness: {thickness}")

                # -------------------------------------------------
                # Output name
                # -------------------------------------------------

                part_name = Path(component.path).stem

                output_path = self.unique_output_path(output_folder, part_name)

                print(f"Exporting:\n  {output_path}")

                # -------------------------------------------------
                # Export
                # -------------------------------------------------

                success = self.export_dwg(component, output_path)

                if success:
                    exported += 1

                    print("DWG exported successfully.")

                else:
                    failed += 1

                    print("DWG export FAILED.")

            except Exception as e:
                failed += 1

                print(f"ERROR: {e}")

        # =====================================================
        # FINAL REPORT
        # =====================================================

        print()
        print("=" * 60)
        print("FINAL REPORT")
        print("=" * 60)

        print(f"Unique parts: {total}")

        print(f"Sheet Metal: {sheet_metal}")

        print(f"Exported: {exported}")

        print(f"Not Sheet Metal: {non_sheet_metal}")

        print(f"Failed: {failed}")

        print()
        print(f"Output: {output_folder}")

        return {
            "unique_parts": total,
            "sheet_metal": sheet_metal,
            "exported": exported,
            "not_sheet_metal": non_sheet_metal,
            "failed": failed,
        }
