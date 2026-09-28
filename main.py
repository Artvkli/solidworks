from pathlib import Path

from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.sheet_metal import SheetMetalDetector
from solidworks.autocad import AutoCADExporter


def format_thickness(thickness_mm):
    """1.500 -> 1.5mm | 2.000 -> 2mm | 15.000 -> 15mm"""

    value = f"{thickness_mm:.3f}".rstrip("0").rstrip(".")
    return f"{value}mm"


def header(title):
    print()
    print("=" * 60)
    print(title)
    print("=" * 60)


def main():

    header("SolidWorks Automation")

    # --------------------------------------------------
    # Project paths
    # --------------------------------------------------

    project_folder = Path(__file__).parent

    input_folder = project_folder / "input"
    output_folder = project_folder / "output"
    temp_folder = output_folder / "_temp"

    input_folder.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)
    temp_folder.mkdir(exist_ok=True)

    # --------------------------------------------------
    # Connect to SolidWorks
    # --------------------------------------------------

    connection = SolidWorksConnection()

    if not connection.connect():
        print()
        print("Program stopped.")
        return

    sw_app = connection.get_application()

    # --------------------------------------------------
    # Find assemblies
    # --------------------------------------------------

    assemblies = find_assemblies(input_folder)

    if not assemblies:
        print()
        print("No SLDASM files found.")
        print("Please put a SolidWorks Assembly inside:")
        print(input_folder)
        return

    print()
    print("Assemblies found:")

    for index, assembly_file in enumerate(assemblies, start=1):
        print(f"{index}. {assembly_file.name}")

    # --------------------------------------------------
    # Open first assembly
    # --------------------------------------------------

    assembly_file = assemblies[0]

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_file):
        return

    # --------------------------------------------------
    # Components
    # --------------------------------------------------

    components = assembly.get_components()

    header("Component Types")

    part_count = sum(1 for c in components if c.is_part)
    assembly_count = sum(1 for c in components if c.is_assembly)

    print("Total components:", len(components))
    print("PART:", part_count)
    print("ASSEMBLY:", assembly_count)

    # --------------------------------------------------
    # Collect unique Sheet Metal parts
    # --------------------------------------------------

    detector = SheetMetalDetector(sw_app)

    header("Collecting Sheet Metal Parts")

    components_by_path = {
        str(c.path).lower(): c for c in components if c.is_part and not c.suppressed
    }

    sheet_metal_data = []

    for item in assembly.get_unique_part_quantities(components):
        component = components_by_path.get(str(item["path"]).lower())

        if component is None:
            continue

        thickness = detector.get_thickness(component)

        if thickness is None:
            print(f"{item['name']} | not sheet metal (skipped)")
            continue

        thickness_mm = thickness * 1000

        sheet_metal_data.append(
            {
                "name": item["name"],
                "path": item["path"],
                "quantity": item["quantity"],
                "thickness": thickness_mm,
            }
        )

        print(f"{item['name']} | {thickness_mm:.3f} mm | QTY: {item['quantity']}")

    if not sheet_metal_data:
        print()
        print("No Sheet Metal parts found.")
        return

    # --------------------------------------------------
    # Group by thickness
    # --------------------------------------------------

    groups = {}

    for item in sheet_metal_data:
        key = round(item["thickness"], 3)
        groups.setdefault(key, []).append(item)

    header("Sheet Metal Groups")

    for thickness, parts in sorted(groups.items()):
        print()
        print(f"[{format_thickness(thickness)}]")

        for part in parts:
            print(f"  {part['name']} | QTY: {part['quantity']}")

    # --------------------------------------------------
    # Export temporary DWGs
    # --------------------------------------------------

    header("Exporting Flat Patterns")

    for thickness, parts in sorted(groups.items()):
        thickness_name = format_thickness(thickness)
        thickness_folder = temp_folder / thickness_name
        thickness_folder.mkdir(parents=True, exist_ok=True)

        print()
        print("-" * 60)
        print(f"Thickness Group: {thickness_name}")
        print("-" * 60)

        for part in parts:
            print()
            print("Part:", part["name"])
            print("QTY:", part["quantity"])

            model = detector.open_part(part["path"])

            if model is None:
                print("[FAILED] Could not open part.")
                continue

            try:
                if detector.find_flat_pattern_feature(model) is None:
                    print("[NO FLAT PATTERN]")
                    continue

                print("[FLAT PATTERN FOUND]")

                if not detector.activate_flat_pattern(model):
                    print("[FLAT PATTERN ACTIVATION FAILED]")
                    continue

                print("[FLAT PATTERN ACTIVE]")

                output_file = thickness_folder / f"{part['name']}.dwg"

                if detector.export_dxf(model, output_file):
                    print("[EXPORT SUCCESS]")
                    print(output_file)
                else:
                    print("[EXPORT FAILED]")

                detector.deactivate_flat_pattern(model)

            finally:
                detector.close_part(model)

    # --------------------------------------------------
    # Connect to AutoCAD
    # --------------------------------------------------

    header("Creating Final DWG Files")

    autocad = AutoCADExporter()

    if not autocad.connect():
        print()
        print("AutoCAD connection failed.")
        print("Temporary DWGs were created, but final DWGs were not created.")
        return

    for thickness, parts in sorted(groups.items()):
        thickness_name = format_thickness(thickness)
        thickness_folder = temp_folder / thickness_name
        final_dwg = output_folder / f"{thickness_name}.dwg"

        print()
        print("-" * 60)
        print(f"Creating: {final_dwg}")
        print("-" * 60)

        final_parts = []

        for part in parts:
            temp_dwg = thickness_folder / f"{part['name']}.dwg"

            if not temp_dwg.exists():
                print("[SKIP] Temporary DWG not found:")
                print(temp_dwg)
                continue

            final_parts.append(
                {
                    "name": part["name"],
                    "dwg_path": temp_dwg,
                    "quantity": part["quantity"],
                }
            )

        if not final_parts:
            print("[SKIP] No exported parts.")
            continue

        success = autocad.create_thickness_dwg(
            final_dwg,
            final_parts,
            spacing=100.0,
        )

        print()
        if success:
            print("[FINAL DWG SUCCESS]")
            print(final_dwg)
        else:
            print("[FINAL DWG FAILED]")
            print(thickness_name)

    header("Program finished.")


if __name__ == "__main__":
    main()
