import os

# Check if Lines.xlsx exists
lines_path = r"C:\Users\Marco.Africani\Desktop\Month recap\Lines.xlsx"

if os.path.exists(lines_path):
    print(f"Lines.xlsx exists at: {lines_path}")
    print("Please close the file if it's open in Excel, then re-run the converter.")
else:
    print("Lines.xlsx does not exist yet.")
    print("The converter will create it on the next run.")
