# PPTX & PDF Utility Tool

A lightweight, Python-based CLI tool designed to simplify operations on PowerPoint presentations and PDF documents. With an interactive menu-driven interface, you can easily merge PowerPoint files, convert PowerPoint presentations to PDF, and merge multiple PDF documents into a single file.

## Features

- **Merge PowerPoint Files:** Combine multiple `.ppt` or `.pptx` files into a single presentation.
- **Convert PowerPoint to PDF:** Convert your `.ppt` or `.pptx` presentations into `.pdf` documents losslessly.
- **Merge PDF Files:** Combine multiple PDF files into a single, continuous PDF document.

## Directory Structure

The application operates using specific directories for input and output. Place your files in the respective folder before running the operation:

- `Merge_pptx/`: Place `.ppt` or `.pptx` files here to merge them.
- `Convert_pptx_pdf/`: Place `.ppt` or `.pptx` files here to convert them to PDFs.
- `Merge_pdf/`: Place `.pdf` files here to merge them.

## Requirements

- Python 3.8+
- Required libraries are listed in `requirements.txt`:
  - `Spire.Presentation`
  - `pypdf`

## Installation & Setup

1. **Navigate to the application directory** (or clone the repository):
   ```bash
   cd pptx-pdf-tools
   ```

2. **Set up a Python Virtual Environment (Recommended):**
   This keeps the application's dependencies isolated.
   ```bash
   python -m venv .venv
   source .venv/bin/activate    # On macOS/Linux
   # or
   .venv\Scripts\activate       # On Windows
   ```

3. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Add Files:** Add the files you want to work with into their corresponding operational folders (`Merge_pptx`, `Convert_pptx_pdf`, or `Merge_pdf`).

2. **Run the Script:**
   ```bash
   python file_tools.py
   ```

3. **Follow the Menu Instructions:**
   You will be greeted with an interactive prompt:
   ```text
   --- PPTX and PDF Utility Tool ---

   Please choose an option:
   1. Merge all PPT/PPTX files in 'Merge_pptx' folder
   2. Convert all PPT/PPTX files to PDF in 'Convert_pptx_pdf' folder
   3. Merge all PDF files in 'Merge_pdf' folder
   4. Exit
   ```
   Simply enter the number corresponding to the action you want to take and press **Enter**.

4. **Outputs:** The generated files (merged PPT, converted PDF, merged PDF) will be appended with a timestamp to prevent overwriting prior exports, and are saved directly into the folder you are working within.
