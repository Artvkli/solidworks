import win32com.client

sw = win32com.client.GetActiveObject("SldWorks.Application")

print("Connected!")
print("Version:", sw.RevisionNumber())



import os
import win32com.client


INPUT_FOLDER = r"C:\Drawings"
OUTPUT_FOLDER = r"C:\DWG_Output"


os.makedirs(OUTPUT_FOLDER, exist_ok=True)


sw = win32com.client.GetActiveObject("SldWorks.Application")

print("Connected to SOLIDWORKS")



files = os.listdir(INPUT_FOLDER)

drawing_files = [
    f for f in files
    if f.lower().endswith(".slddrw")
]

    
for filename in drawing_files:

    input_path = os.path.join(INPUT_FOLDER, filename)

    name = os.path.splitext(filename)[0]

    output_path = os.path.join(
        OUTPUT_FOLDER,
        name + ".dwg"
    )

    print(f"Converting: {filename}")

    errors = 0
    warnings = 0

    drawing = sw.OpenDoc6(
        input_path,
        3,
        0,
        "",
        errors,
        warnings
    )

    if drawing is None:
        print(f"ERROR: Could not open {filename}")
        continue

    success = drawing.SaveAs3(
        output_path,
        0,
        0
    )

    if success:
        print(f"OK: {output_path}")
    else:
        print(f"ERROR: Could not save {filename}")

    sw.CloseDoc(drawing.GetTitle)


print("Finished!")

