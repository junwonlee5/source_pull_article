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
pip install pandas python-docx lxml openpyxl pyinstaller# source_pull_article
