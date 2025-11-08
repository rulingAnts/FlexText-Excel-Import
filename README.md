# Native-Speaker Excel Transcription/Back-Translation Tool

> 📖 **[Visit the user-friendly website](https://rulingants.github.io/FlexText-Excel-Import/)** for complete documentation, downloads, and tutorials!

## Overview

This repository hosts a utility designed to streamline the process of transcription and back-translation for minority language texts. It pairs a structured [**Excel Template**](https://raw.githubusercontent.com/rulingAnts/FlexText-Excel-Import/refs/heads/main/Interlinear%20Text%20Excel%20Template.xltx) with a Python-based converter that generates **FlexText** files suitable for direct import into **Fieldworks Language Explorer (FLEx)**.

The primary goal is to provide a user-friendly format (Excel spreadsheet) for native speakers who may not be familiar with complex linguistic software, allowing them to easily contribute interlinearized text data.

## Easy Version

Download the app and the Excel templates [here](https://github.com/rulingAnts/FlexText-Excel-Import/releases/latest). Be sure to download both the portable app (exe file) and the Excel templates (zip file). Use one of the excel templates to enter a text back translation (or have a mother-tongue-speaking colleague use it for that) and then double click the app to convert the Excel sheet into a file type that FLEx can import. Then go to File > Import... > FlexText Interlinear...

You will need to make sure that your writing system codes match what you enter at the top of the excel sheet (to the right of the title, names, etc), and also make sure that while you're working on that text, you order your writing systems in writing system settings and in interlinear settings to match what is in the top of the excel sheet. Otherwise you'll see nothing. But if your writing system codes are correct in the sheet, and your writing system settings and interlinear settings match (each tab has its own interlinear settings and the text chart only matches the project-wide settings), then it will display correctly. This is a limitation in the way FLEx works, but once you learn how to adjust writing system and interlinear settings, it's mildly annoying, but still usable. If you're using the same glossing language and vernacular writing system as your mother-tongue-speaking colleague, then you will not have that problem.

> **Critical Warning:** When entering writing system identifiers, use the value labeled **Code** in FLEx's *Tools → Configure → Writing Systems…* dialog (General tab). **Do NOT use** the *Abbreviation* or the language name. If the codes in your Excel sheet do not exactly match existing codes, FLEx will silently create **new writing systems** and your imported data will be placed there, causing duplicates you must later clean up.

See: **[Writing System Codes (detailed guide)](docs/writing-system-codes.html)**

## For advanced, programming geeks like me: CLI Version

The command-line interface (CLI) is still supported, through **Python scripts** which perform a two-stage conversion process:

1. **Stage 1: `excel_to_xml.py`**: Converts the structured data from the Excel template (`.xlsx`) into an intermediate custom XML format.
  
2. **Stage 2: `xml_to_flextext.py`**: Converts the intermediate XML file into the final FLEx-compatible FlexText format (`.flextext`).
  

## Setup and GUI Usage

This guide details how to set up the Python environment and use the GUI to convert your transcribed Excel data into a FLEx-compatible FlexText file. **These instructions assume you are running a Windows operating system.**

### Prerequisites

1. **Python:** You must have Python 3.8 or newer installed on your system. You can download the latest version for Windows from the [official Python website](https://www.python.org/downloads/windows/).
  
2. **Libraries:** The scripts rely on the following Python libraries:
  
  - `openpyxl`: For reading data from the `.xlsx` Excel file.
    
  - `lxml` or similar: For efficient XML creation and manipulation.
    

You can install the necessary libraries using `pip`:

```
pip install openpyxl lxml
```

### Step 1: Prepare Your Data

1. Download and fill out the **Excel Template** ([`Interlinear Text Excel Template.xltx`](https://raw.githubusercontent.com/rulingAnts/FlexText-Excel-Import/refs/heads/main/Interlinear%20Text%20Excel%20Template.xltx)).

2. Ensure the writing system codes in the header are filled correctly using the **Code** values from your FLEx project (not Abbreviation or language name). See [docs/writing-system-codes.html](docs/writing-system-codes.html).
  
3. Save your completed transcription spreadsheet as an `.xlsx` file (e.g., `MyStory.xlsx`).
  

### Step 2: Run conversion GUI

Run `convert_interlinear_gui.py` in Python. The GUI window should be self-explanatory.

### Step 3: Import into Fieldworks Language Explorer (FLEx)

1. Open your FLEx project.
  
2. Go to **File** -> **Import** -> **Interlinear Text...**
  
3. Select the generated **`.flextext`** file (e.g., `MyStory.flextext`).
  
4. Follow the prompts in FLEx to complete the import. Your transcribed and back-translated text should now be available as a new interlinear text.
  

## Upcoming Features

A **portable standalone EXE file** is in development. This will simplify the process to a single step for end-users, requiring only a file selection and a button click.

Other input and output formats are also planned.

## Contributing

We welcome all contributions! Help us improve this tool by submitting bug reports or feature ideas via GitHub Issues, or by opening Pull Requests with code improvements. Your involvement benefits the entire language community.

## License

This project is released under the **GNU Affero General Public License v3.0 (AGPL-3.0)**. Please see the `LICENSE` file for full details.

## Acknowledgments and Copyright

### Copyright

Copyright (c) 2025, Seth Johnston and Brian Plimley. All rights reserved.

### Development Credit

This application's conceptual design and initial implementation were developed by **Gemini**, an advanced large language model built by Google, directed by Seth Johnston with coding contributions by Brian Plimley. The Excel template was designed and created by Seth Johnston.

### Technology Ecosystem

We gratefully acknowledge the Python ecosystem and the use of powerful open-source libraries (such as `openpyxl` for Excel handling and `lxml` for XML processing) which make this tool possible.
