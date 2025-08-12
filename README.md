
# 🚀 PPTCreate

![Build](https://img.shields.io/badge/build-passing-brightgreen)
![Python](https://img.shields.io/badge/python-3.8+-blue)

> 📊 **Batch generate PowerPoint presentations from structured data and speaker notes effortlessly.**

---

## ✨ Features

- ⚡ **Automated PPT slide generation** using templates and structured data (Excel, notes)
- 🗣️ **Speaker notes**: Extract and add notes for each slide
- 🧩 **Smart layout detection** for optimal slide designs
- 🏷️ **Shape/layout analysis** for easy template creation
- 🌐 **Streamlit web app** for GUI batch processing and OneDrive integration

---

## 🗂️ Repository Structure

| File              | Description                                         |
|-------------------|-----------------------------------------------------|
| `main.py`         | 🖥️  Streamlit batch processor (GUI/auth/batch logic) |
| `pptcreator.py`   | 🏗️  Core PPT generation logic                        |
| `slidenum.py`     | 🔢  Layout/placeholder analysis utilities            |
| `dataextractor.py`| 📤  Extraction & preprocessing from PPTX             |
| `layoutdata.py`   | 📊  Layout analysis & Excel export                   |
| `notes.py`        | 📝  Speaker notes extraction/insertion               |
| `shaperet.py`     | 🔲  Shape-level metadata extraction                  |


---

## 🚦 Getting Started

### Prerequisites

- **Python 3.8+**
- Install dependencies:
  pip install python-pptx pandas streamlit requests msal openpyxl
- *Recommended*: Access to OneDrive if using the Streamlit web integration

---

## 1. Prepare Layout & Data
- Create a **PowerPoint** template with well-defined slide layouts and placeholders.

- Prepare an Excel sheet describing the layout, shape indices, and mapping (or use the provided scripts to extract from PPTX).

## 2. Running the Batch Processor
- The main web application (authentication & batch logic) is launched via Streamlit:
## bash
<pre> streamlit run main.py </pre>

1. Authenticate with your Microsoft (Azure) account if using OneDrive integration.

2. Select input folders and layout files.

3. Run the batch process to generate presentations.

## 3. Scripted Generation
You can use pptcreator.py directly in your own Python scripts or notebooks:
<pre>python
from pptcreator import pptcreate
pptcreate(
    excel="layout_shapes.xlsx",
    ppt="Layout_MUJ_V9.pptx",
    out="OUTPUT_MyPresentation.pptx",
    sninput="Unit_SampleNotes.pptx"
)</pre>

## Working
- How It Works:<br>
  * Shape Metadata Extraction: Utilities (shaperet.py, layoutdata.py, etc.) scans slide layouts for placeholders/shapes and exports the data to Excel for mapping.

  * PPT Creation: The main engine generates new slides, populates them with text and headers, and matches content to optimal layouts (based on number of paragraphs).

  * Speaker Notes: Speaker notes are extracted from a source text ppt and inserted into graphic ppt as required.

  * Web UI: Streamlit front-end for authentication, batch download, and processing with feedback/status reporting.

- Customization:<br>
  * Adjust layout matching logic in slidenum.py or your layout mapping Excel file to match your template designs.

  * Modify pptcreator.py to tweak slide content, title handling, or bullet formatting as needed.


## Acknowledgments
- Built using python-pptx

- Streamlit for the web interface

- MSAL for OneDrive authentication
