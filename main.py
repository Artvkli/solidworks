from pathlib import Path

from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.sheet_metal import SheetMetalDetector


LINE = "=" * 60
THIN_LINE = "-" * 60


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------


def print_header(title):
    print()
    print(LINE)
    print(title)
    print(LINE)


def format_thickness(thickness_mm):
    """Format thickness for grouping/output filename."""

    if thickness_mm is None:
        return None

    if abs(thickness_mm - round(thickness_mm)) < 0.0001:
        return f"{int(round(thickness_mm))}mm"

    text = f"{thickness_mm:.3f}".rstrip("0").rstrip(".")
    return f"{text}mm"


def setup_folders():
    """Create and return (input, output, temp) folders."""

    project_folder = Path(__file__).parent

    input_folder = project_folder / "input"
    output_folder = project_folder / "output"
    temp_folder = output_folder / "_temp"

    for folder in (input_folder, output_folder, temp_folder):
        folder.mkdir(exist_ok=True)

    return input_folder, output_folder, temp_folder


# ---------------------------------------------------------
# Steps
# ---------------------------------------------------------


def connect_to_solidworks():
    print()
    print("Connecting to SolidWorks...")

    connection = SolidWorksConnection()

    if not connection.connect():
        print()
        print("Program stopped.")
        return None

    return connection.get_application()


def select_assembly(input_folder):
    assemblies = find_assemblies(input_folder)

    if not assemblies:
        print()
        print("No SLDASM files found.")
        print()
        print("Please put a SolidWorks Assembly inside:")
        print(input_folder)
        return None

    print()
    print("Assemblies found:")

    for index, assembly_file in enumerate(assemblies, start=1):
        print(f"{index}. {assembly_file.name}")

    # Select first assembly
    assembly_file = assemblies[0]

    print()
    print("Selected assembly:")
    print(assembly_file)

    return assembly_file


def print_component_summary(components):
    part_count = 0
    assembly_count = 0
    suppressed_count = 0
    lightweight_count = 0

    for component in components:
        if component.is_part:
            part_count += 1
        elif component.is_assembly:
            assembly_count += 1

        if getattr(component, "suppressed", False):
            suppressed_count += 1

        if getattr(component, "lightweight", False):
            lightweight_count += 1

    print_header("Component Summary")
    print()
    print("Total components :", len(components))
    print("PART             :", part_count)
    print("ASSEMBLY         :", assembly_count)
    print("SUPPRESSED       :", suppressed_count)
    print("LIGHTWEIGHT      :", lightweight_count)


def find_representative(components, part_path):
    """Find one actual component instance for the given part file."""

    target = str(part_path).lower()

    for component in components:
        if component.is_part and str(component.path).lower() == target:
            return component

    return None


def print_part_info(representative, checked, total, part_name, part_path, quantity):
    print()
    print(THIN_LINE)
    print(f"Checking {checked}/{total}")
    print("Part:", part_name)
    print("Path:", part_path)
    print("QTY :", quantity)
    print("Suppression:", getattr(representative, "suppression", None))
    print("Suppressed:", getattr(representative, "suppressed", False))
    print("Lightweight:", getattr(representative, "lightweight", False))
    print("Configuration:", getattr(representative, "configuration", None))


def collect_sheet_metal_parts(detector, components, unique_parts):
    print_header("Collecting Sheet Metal Parts")

    sheet_metal_parts = []

    for checked, item in enumerate(unique_parts, start=1):
        part_path = Path(item["path"])
        part_name = item["name"]
        quantity = item["quantity"]

        representative = find_representative(components, part_path)

        if representative is None:
            print()
            print("[SKIP] Could not find component instance:")
            print(part_name)
            continue

        print_part_info(
            representative,
            checked,
            len(unique_parts),
            part_name,
            part_path,
            quantity,
        )

        # Detect Sheet Metal
        try:
            is_sheet_metal = detector.is_sheet_metal(representative)
        except Exception as e:
            print()
            print("[ERROR] Sheet Metal detection failed:")
            print(e)
            continue

        if not is_sheet_metal:
            print("Result: NOT Sheet Metal")
            continue

        print("Result: >>> SHEET METAL <<<")

        # Read Thickness
        try:
            thickness = detector.get_thickness(representative)
        except Exception as e:
            print()
            print("[ERROR] Thickness detection failed:")
            print(e)
            thickness = None

        if thickness is None:
            print("Thickness: FAILED")
            continue

        thickness_mm = thickness * 1000.0
        thickness_name = format_thickness(thickness_mm)

        print("Thickness:", f"{thickness_mm:.3f} mm")
        print("Group:", thickness_name)

        sheet_metal_parts.append(
            {
                "name": part_name,
                "path": part_path,
                "quantity": quantity,
                "thickness": thickness,
                "thickness_mm": thickness_mm,
                "thickness_name": thickness_name,
                "component": representative,
            }
        )

    return sheet_metal_parts


def print_sheet_metal_summary(sheet_metal_parts):
    print()
    print_header("Sheet Metal Summary")

    print()
    print("Sheet Metal unique parts:", len(sheet_metal_parts))

    total_instances = 0

    for part in sheet_metal_parts:
        total_instances += part["quantity"]

        print()
        print(
            f"{part['name']} | QTY={part['quantity']} | {part['thickness_mm']:.3f} mm"
        )

    print()
    print("Total Sheet Metal instances:", total_instances)


def group_by_thickness(sheet_metal_parts):
    grouped = {}

    for part in sheet_metal_parts:
        grouped.setdefault(part["thickness_name"], []).append(part)

    print_header("Thickness Groups")

    for thickness_name, parts in grouped.items():
        print()
        print(thickness_name)

        for part in parts:
            print(f"  - {part['name']} | QTY={part['quantity']}")

    return grouped


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------


def main():
    print(LINE)
    print("SolidWorks Sheet Metal Automation")
    print(LINE)

    input_folder, output_folder, temp_folder = setup_folders()

    sw_app = connect_to_solidworks()
    if sw_app is None:
        return

    assembly_file = select_assembly(input_folder)
    if assembly_file is None:
        return

    # Open Assembly
    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_file):
        print()
        print("Could not open assembly.")
        return

    # Read Components
    components = assembly.get_components()

    if not components:
        print()
        print("No components found in assembly.")
        return

    print_component_summary(components)

    # Group unique PART files
    unique_parts = assembly.get_unique_part_quantities(components)

    print_header("Unique PART Files")
    print()
    print("Unique parts:", len(unique_parts))

    # Sheet Metal Detection
    detector = SheetMetalDetector(sw_app)
    sheet_metal_parts = collect_sheet_metal_parts(detector, components, unique_parts)

    if not sheet_metal_parts:
        print()
        print_header("Sheet Metal Summary")
        print()
        print("No Sheet Metal parts found.")
        print()
        print("Important:")
        print("The assembly was scanned successfully,")
        print("but no detected Sheet Metal feature was returned.")
        return

    print_sheet_metal_summary(sheet_metal_parts)
    grouped_by_thickness = group_by_thickness(sheet_metal_parts)

    # Current stage finished
    print_header("Sheet Metal Detection Finished")

    print()
    print("Detected parts:", len(sheet_metal_parts))
    print("Thickness groups:", len(grouped_by_thickness))

    print()
    print("Next stage:")
    print("Flat Pattern detection and DWG export.")


if __name__ == "__main__":
    main()
