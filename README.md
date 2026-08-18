# Source Pull - Citation & Footnote Extractor

**Source Pull** is a Python tool designed to extract footnotes, law review citations, statutes, cases, and web links from Word documents (`.docx`) and parse them into structured Excel spreadsheets (`.xlsx`). 

It includes logic for identifying Bluebook citation formats, mapping *id.* and *supra* references back to their original source, and running via a simple GUI file picker.

---

## ✨ Features

- 📄 **Word Processing (`.docx`):** Automatically reads footnotes and XML elements from Microsoft Word documents.
- ⚖️ **Bluebook Citation Categorization:** Identifies and classifies sources into:
  - Cases
  - Law Review Articles
  - Statutes & Regulations
  - Internet Articles
  - Reports / Studies
  - Secondary Sources
- 🔗 **Smart Citation Tracking:** Resolves short-form citations, `id.`, and `supra note X` references to link subsequent citations back to the root source.
- 📊 **Excel Export:** Generates clean `.xlsx` spreadsheets showing the source, type, primary footnote, subsequent footnotes, and extracted links.
- 🖥️ **GUI File Picker:** Opens a native window dialog allowing you to select single or multiple files effortlessly.

---

## 🛠️ Prerequisites

Ensure you have **Python 3.9+** installed on your system.

### Install Required Dependencies

Run the following command in your terminal or command prompt:

```bash
pip install pandas python-docx lxml openpyxl pyinstaller
```
🚀 Step-by-Step Setup & Execution

Option A: Running via Python Script
Open your terminal or command prompt.

Run the script directly:

```Bash
python "Source Pull.py"
```
A native file selection dialog will appear. Select your .docx file(s).

The script will process the documents and export the generated .xlsx spreadsheet(s) into the same directory as the source file.

Option B: Building a Standalone Executable (.exe / .app)
Follow these steps to compile a standalone executable that runs without Python installed.

Step 2: Generate the .spec File
Run the setup script in your terminal:

```Bash
python setup.py
```
Step 3: Build the Application
Compile the program with PyInstaller using the generated spec file:

```Bash
python -m PyInstaller source_pull.spec
```
Step 4: Locate the Executable
Once the build process completes, locate your compiled file inside the dist/ directory:

Windows: dist/SourcePullApp.exe

macOS: dist/SourcePullApp

📁 Project File Structure
```Plaintext
source-pull/
├── Source Pull.py        # Main Python application
├── setup.py             # Script to generate source_pull.spec
├── source_pull.spec     # PyInstaller configuration file
└── README.md            # Documentation
```
⚠️ Notes & Troubleshooting
Platform Dependencies: PyInstaller builds native executables. Compiling on Windows generates a .exe file; compiling on macOS generates an executable binary for macOS.

Windows Defender / SmartScreen: Because the executable is unsigned, Windows may display a security prompt. Click "More Info" → "Run Anyway" to open it.
