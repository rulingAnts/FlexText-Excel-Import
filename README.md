# Native-Speaker Excel Transcrption/Back-Translation Tool

## Overview

This repository hosts a utility designed to streamline the process of transcription and back-translation for minority language texts. It pairs a structured [**Excel Template**](https://raw.githubusercontent.com/rulingAnts/FlexText-Excel-Import/refs/heads/main/Interlinear%20Text%20Excel%20Template.xltx) with a Python-based converter that generates **FlexText** files suitable for direct import into **Fieldworks Language Explorer (FLEx)**.

The primary goal is to provide a user-friendly format (Excel spreadsheet) for native speakers who may not be familiar with complex linguistic software, allowing them to easily contribute interlinearized text data.

## Current Status (CLI Version)

The current implementation is distributed as two command-line interface (CLI) **Python scripts**, which perform a two-stage conversion process:

1. **Stage 1: `excel_to_xml.py`**: Converts the structured data from the Excel template (`.xlsx`) into an intermediate custom XML format.
  
2. **Stage 2: `xml_to_flextext.py`**: Converts the intermediate XML file into the final FLEx-compatible FlexText format (`.flextext`).
  

## Setup and CLI Usage

This guide details how to set up the Python environment and use the two command-line interface (CLI) scripts to convert your transcribed Excel data into a FLEx-compatible FlexText file. **These instructions assume you are running a Windows operating system.**

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
  
2. Save your completed transcription spreadsheet as an `.xlsx` file (e.g., `MyStory.xlsx`).
  

### Step 2: Run Stage 1 (Excel to Intermediate XML)

The first script converts your Excel file (`.xlsx`) into an intermediate XML file (`.xml`).

**Usage:** Provide the full path to your Excel file as the first command-line argument. **Always enclose the path in quotes if it contains spaces.**

```
python excel_to_xml.py "C:\Path\To\MyStory.xlsx"
```

**Output:** The script will save the intermediate XML file in the *same directory* as your input file, using the same name with a `.xml` extension (e.g., `C:\Path\To\MyStory.xml`).

### Step 3: Run Stage 2 (Intermediate XML to FlexText)

The second script converts the intermediate XML file (`.xml`) into the final FlexText file (`.flextext`).

**Usage:** Provide the full path to the intermediate XML file as the first command-line argument.

```
python xml_to_flextext.py "C:\Path\To\MyStory.xml"
```

**Output:** The script will save the final FlexText file in the *same directory* as the XML input file, using the same name with a `.flextext` extension (e.g., `C:\Path\To\MyStory.flextext`).

### Step 4: Import into Fieldworks Language Explorer (FLEx)

1. Open your FLEx project.
  
2. Go to **File** -> **Import** -> **Interlinear Text...**
  
3. Select the generated **`.flextext`** file (e.g., `MyStory.flextext`).
  
4. Follow the prompts in FLEx to complete the import. Your transcribed and back-translated text should now be available as a new interlinear text.
  

## Upcoming Features

A **GUI (Graphical User Interface) version** is currently in development and will be distributed as a **portable standalone EXE file**. This will simplify the process to a single step for end-users, requiring only a file selection and a button click.

## License

This project is released under the **GNU Affero General Public License v3.0 (AGPL-3.0)**. Please see the `LICENSE` file for full details.

## Acknowledgments and Copyright

### Copyright

Copyright (c) 2025, Seth Johnston. All rights reserved.

### Development Credit

This application's conceptual design and initial implementation were developed by **Gemini**, an advanced large language model built by Google, directed by **Seth Johnston**.

### Technology Ecosystem

We gratefully acknowledge the Python ecosystem and the use of powerful open-source libraries (such as `openpyxl` for Excel handling and `lxml` for XML processing) which make this tool possible.
