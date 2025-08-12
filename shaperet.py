
from pptx import Presentation
import pandas as pd 
pptfile=r"Sample_Layout.pptx"

"""
Analyzes master slides in a PowerPoint presentation and saves the details to an Excel file.

Args:
  pptx_path (str): The path to the PowerPoint file.
  output_excel_path (str): The path to save the output Excel file.
"""
try:
  prs = Presentation(pptfile)
except Exception as e:
  print(f"Error opening PowerPoint file: {e}")


data = []

for i,layout in enumerate(prs.slide_layouts):
  for shape in layout.placeholders:
    text = ""
    if hasattr(shape, "text"):
      text = shape.text

    data.append({
        "Layout number": i,
        "Layout Name": layout.name,
        "Shape idx":shape.placeholder_format.idx,
        "Shape Type": f'{shape.placeholder_format.type}',
        "Shape Name": shape.name,
        "Text": text.strip()
    })

if data:
  df = pd.DataFrame(data)
  try:
    df.to_excel(r"layout_shapes.xlsx", index=False)
    print("Successfully saved master slide data to layouts_shape.xlsx")
  except Exception as e:
    print(f"Error saving data to Excel file: {e}")
else:
  print("No shapes found in master slides.")



