# A shorter and easier version of https://github.com/matthewrenze/powerpoint-to-pdf/blob/master/ConvertAll.py

import os
import comtypes.client

path = input("Enter the path:\n")

if path:
    os.chdir(path)
files = os.listdir()

powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
powerpoint.Visible = 1

for file in files:
    if file.endswith((".ppt", ".pptx")):
        file_path = os.path.abspath(file)
        if file.endswith(".ppt"):
            output = file.replace(".ppt", ".pdf")
        else:
            output = file.replace(".pptx", ".pdf")
        output_path = os.path.abspath(output)
        slides = powerpoint.Presentations.Open(file_path)
        slides.SaveAs(output_path, 32)
        slides.Close()
        print(f"Converted {file} to pdf")
