from pathlib import Path
from solidworks.connection import SolidWorksConnection
from solidworks.assembly import SolidWorksAssembly, find_assemblies
from solidworks.sheet_metal import SheetMetalDetector
from solidworks.autocad import AutoCADExporter


def format_thickness(thickness_mm):
    """
    Convert thickness to a clean name.

    Examples:
        1.500 -> 1.5mm
        2.000 -> 2mm
        15.000 -> 15mm
    """

    value = f"{thickness_mm:.3f}".rstrip("0").rstrip(".")

    return f"{value}mm"


def main():

    print("=" * 60)
    print("SolidWorks Automation")
    print("=" * 60)

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
        print()
        print("Please put a SolidWorks Assembly inside:")
        print(input_folder)

        return

    print()
    print("Assemblies found:")

    for index, assembly_file in enumerate(
        assemblies,
        start=1
    ):
        print(
            f"{index}. {assembly_file.name}"
        )

    # --------------------------------------------------
    # Open first assembly
    # --------------------------------------------------

    assembly_file = assemblies[0]

    print()
    print("Opening assembly:")
    print(assembly_file.name)

    assembly = SolidWorksAssembly(sw_app)

    if not assembly.open(assembly_file):
        return

    # --------------------------------------------------
    # Get components
    # --------------------------------------------------

    components = assembly.get_components()


    print()
    print("=" * 60)
    print("Component Types")
    print("=" * 60)

    part_count = 0
    assembly_count = 0

    for component in components:

        if component.is_part:
            part_count += 1

        elif component.is_assembly:
            assembly_count += 1

    print("PART:", part_count)
    print("ASSEMBLY:", assembly_count)

    print()
    print("First 20 components:")

    for component in components[:20]:

        print(
            f"{component.name} | "
            f"{component.component_type} | "
            f"suppressed={component.suppressed}"
        )



#     print()
#     print("Total components:", len(components))

#     # --------------------------------------------------
#     # Sheet Metal Detector
#     # --------------------------------------------------

#     detector = SheetMetalDetector(sw_app)

#     print()
#     print("=" * 60)
#     print("Collecting Sheet Metal Parts")
#     print("=" * 60)

#     sheet_metal_data = []

#     # --------------------------------------------------
#     # Collect unique Sheet Metal parts
#     # --------------------------------------------------

#     for item in assembly.get_unique_part_quantities():

#         component = next(
#             (
#                 c
#                 for c in components
#                 if c.is_part
#                 and not c.suppressed
#                 and str(c.path).lower()
#                 == str(item["path"]).lower()
#             ),
#             None,
#         )

#         if component is None:
#             continue

#         # ----------------------------------------------
#         # Get thickness
#         # ----------------------------------------------

#         thickness = detector.get_thickness(component)

#         if thickness is None:
#             continue

#         thickness_mm = thickness * 1000

#         data = {
#             "name": item["name"],
#             "path": item["path"],
#             "quantity": item["quantity"],
#             "thickness": thickness_mm,
#         }

#         sheet_metal_data.append(data)

#         print(
#             f"{item['name']} | "
#             f"{thickness_mm:.3f} mm | "
#             f"QTY: {item['quantity']}"
#         )

#     # --------------------------------------------------
#     # Check Sheet Metal
#     # --------------------------------------------------

#     if not sheet_metal_data:

#         print()
#         print("No Sheet Metal parts found.")

#         return

#     # --------------------------------------------------
#     # Group by thickness
#     # --------------------------------------------------

#     groups = {}

#     for item in sheet_metal_data:

#         thickness_key = round(
#             item["thickness"],
#             3
#         )

#         if thickness_key not in groups:
#             groups[thickness_key] = []

#         groups[thickness_key].append(item)

#     # --------------------------------------------------
#     # Show groups
#     # --------------------------------------------------

#     print()
#     print("=" * 60)
#     print("Sheet Metal Groups")
#     print("=" * 60)

#     for thickness, parts in sorted(
#         groups.items()
#     ):

#         thickness_name = format_thickness(
#             thickness
#         )

#         print()
#         print(f"[{thickness_name}]")

#         for part in parts:

#             print(
#                 f"  {part['name']} | "
#                 f"QTY: {part['quantity']}"
#             )

#     # --------------------------------------------------
#     # Export temporary DWGs
#     # --------------------------------------------------

#     print()
#     print("=" * 60)
#     print("Exporting Flat Patterns")
#     print("=" * 60)

#     for thickness, parts in sorted(
#         groups.items()
#     ):

#         thickness_name = format_thickness(
#             thickness
#         )

#         thickness_folder = (
#             temp_folder / thickness_name
#         )

#         thickness_folder.mkdir(
#             parents=True,
#             exist_ok=True
#         )

#         print()
#         print("-" * 60)
#         print(
#             f"Thickness Group: "
#             f"{thickness_name}"
#         )
#         print("-" * 60)

#         for part in parts:

#             print()
#             print("Part:", part["name"])
#             print("QTY:", part["quantity"])

#             # ------------------------------------------
#             # Open part
#             # ------------------------------------------

#             model = detector.open_part(
#                 part["path"]
#             )

#             if model is None:

#                 print(
#                     "[FAILED] "
#                     "Could not open part."
#                 )

#                 continue

#             # ------------------------------------------
#             # Find Flat Pattern
#             # ------------------------------------------

#             flat_pattern = (
#                 detector.find_flat_pattern_feature(
#                     model
#                 )
#             )

#             if flat_pattern is None:

#                 print(
#                     "[NO FLAT PATTERN]"
#                 )

#                 continue

#             print(
#                 "[FLAT PATTERN FOUND]"
#             )

#             # ------------------------------------------
#             # Activate Flat Pattern
#             # ------------------------------------------

#             activated = (
#                 detector.activate_flat_pattern(
#                     model
#                 )
#             )

#             if not activated:

#                 print(
#                     "[FLAT PATTERN "
#                     "ACTIVATION FAILED]"
#                 )

#                 continue

#             print(
#                 "[FLAT PATTERN ACTIVE]"
#             )

#             # ------------------------------------------
#             # Temporary DWG
#             # ------------------------------------------

#             output_file = (
#                 thickness_folder
#                 / f"{part['name']}.dwg"
#             )

#             # ------------------------------------------
#             # Export
#             # ------------------------------------------

#             exported = detector.export_dxf(
#                 model,
#                 output_file
#             )

#             if exported:

#                 print(
#                     "[EXPORT SUCCESS]"
#                 )
#                 print(
#                     output_file
#                 )

#             else:

#                 print(
#                     "[EXPORT FAILED]"
#                 )

#     # --------------------------------------------------
#     # Connect to AutoCAD
#     # --------------------------------------------------

#     print()
#     print("=" * 60)
#     print("Creating Final DWG Files")
#     print("=" * 60)

#     autocad = AutoCADExporter()

#     if not autocad.connect():

#         print()
#         print(
#             "AutoCAD connection failed."
#         )

#         print(
#             "Temporary DWGs were created, "
#             "but final DWGs were not created."
#         )

#         return

#     # --------------------------------------------------
#     # Create final DWG for each thickness
#     # --------------------------------------------------

#     for thickness, parts in sorted(
#         groups.items()
#     ):

#         thickness_name = format_thickness(
#             thickness
#         )

#         thickness_folder = (
#             temp_folder / thickness_name
#         )

#         final_dwg = (
#             output_folder
#             / f"{thickness_name}.dwg"
#         )

#         print()
#         print("-" * 60)
#         print(
#             f"Creating: {final_dwg}"
#         )
#         print("-" * 60)

#         final_parts = []

#         for part in parts:

#             temp_dwg = (
#                 thickness_folder
#                 / f"{part['name']}.dwg"
#             )

#             if not temp_dwg.exists():

#                 print(
#                     "[SKIP] Temporary DWG "
#                     "not found:"
#                 )

#                 print(temp_dwg)

#                 continue

#             final_parts.append(
#                 {
#                     "name": part["name"],
#                     "dwg_path": temp_dwg,
#                     "quantity": part["quantity"],
#                 }
#             )

#         if not final_parts:

#             print(
#                 "[SKIP] No exported parts."
#             )

#             continue

#         success = (
#             autocad.create_thickness_dwg(
#                 final_dwg,
#                 final_parts,
#                 spacing=100.0
#             )
#         )

#         if success:

#             print()
#             print(
#                 "[FINAL DWG SUCCESS]"
#             )

#             print(final_dwg)

#         else:

#             print()
#             print(
#                 "[FINAL DWG FAILED]"
#             )

#             print(thickness_name)

#     # --------------------------------------------------
#     # Final
#     # --------------------------------------------------

#     print()
#     print("=" * 60)
#     print("Program finished.")
#     print("=" * 60)


if __name__ == "__main__":
    main()

