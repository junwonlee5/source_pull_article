import os
import re
import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from docx import Document
from lxml import etree


def identify_source_type(text, url):
    text_clean = text.strip().replace("’", "'").replace("‘", "'")

    # 1. Law Review Pattern
    law_review_pattern = r"\d+\s+[A-Z][A-Za-z\.\s\&\'']+\s+(?:L\.\s+Rev\.|J\.|L\.\s+&|Rev\.|Notre\s+Dame\s+L\.\s+Rev\.)(?:\s+[A-Za-z]+)?\s+\d+"
    is_law_review = re.search(law_review_pattern, text_clean)

    # 2. Case Markers & Reporter Patterns
    is_case_marker = re.search(r"\s+[vV]\.\s+", text_clean)
    # Matches patterns like: 597 U.S. 1, 123 F.3d 456, 45 Me. 12, etc.
    reporter_pattern = (
        r"\b\d+\s+[A-Z][A-Za-z0-9\.\s\’\']{1,15}\s+\d+(?:,\s*\d+)?\b"
    )
    has_reporter = re.search(reporter_pattern, text_clean)

    # 3. Statute / Regulation Keywords
    statute_keywords = [
        "U.S.C.",
        "§",
        "Stat.",
        "Const.",
        "ADC",
        "C.F.R.",
        "Reg.",
        "Admin. Code",
        "¶",
    ]
    has_statute_keyword = any(m in text_clean for m in statute_keywords)
    admin_code_pattern = r"\b(?:\d+\s+)?[A-Z][A-Za-z\.\s]*\bADC\b"
    is_admin_code = re.search(admin_code_pattern, text_clean)
    is_statute_or_reg = has_statute_keyword or bool(is_admin_code)

    is_report = any(
        m in text_clean.lower() for m in ["dep't", "dept", "report", "review"]
    )

    # Priority Order
    if is_law_review:
        return "Law Review Article"
    elif is_statute_or_reg:
        return "Statute/Regulation"
    # FIXED: Catches cases with 'v.' OR citations that start directly with reporter/volume details
    elif (is_case_marker and has_reporter) or (
        is_case_marker and "v." in text_clean.lower()
    ):
        return "Case"
    elif (
        has_reporter and not is_statute_or_reg and not is_report
    ):  # Catches "597 U.S. 1, 26–27 (2022)"
        return "Case"
    elif url:
        return "Internet Article"
    elif is_report:
        return "Report / Study"
    else:
        return "Secondary Source / Other"


def extract_bluebook_data(word_file, excel_file):
    doc = Document(word_file)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    raw_notes = []
    for part in doc.part.package.parts:
        if "footnotes.xml" in part.partname:
            root = etree.fromstring(part.blob)
            notes = root.xpath("//w:footnote", namespaces=ns)
            for note in notes:
                fn_id_str = note.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id"
                )
                if not fn_id_str:
                    continue

                fn_id = int(fn_id_str)
                text = "".join(
                    [
                        t.text
                        for t in note.xpath(".//w:t", namespaces=ns)
                        if t.text
                    ]
                ).strip()

                if fn_id <= 0:
                    continue
                if text.startswith("*") or any(
                    kw in text.lower()
                    for kw in [
                        "professor of law",
                        "clerk",
                        "associate professor",
                    ]
                ):
                    continue

                if text:
                    raw_notes.append({"id": fn_id, "text": text})

    raw_notes.sort(key=lambda x: x["id"])

    source_history = {}
    short_name_map = {}
    reporter_map = {}
    last_source_key = None

    for visual_num, note in enumerate(raw_notes, start=1):
        fn_num = visual_num

        chunks = re.split(r";(?![^(]*\))", note["text"])

        for i, chunk in enumerate(chunks):
            chunk = chunk.strip()
            if not chunk:
                continue
            current_key = (fn_num, i)

            # --- ID. CHECK ---
            if re.match(r"^id\.?($|\s)", chunk, re.IGNORECASE):
                if last_source_key and last_source_key in source_history:
                    source_history[last_source_key][
                        "Subsequent Footnotes"
                    ].append(str(fn_num))
                continue

            # --- SUPRA CHECK ---
            supra_match = re.search(
                r"supra\s+note\s+(\d+)", chunk, re.IGNORECASE
            )
            if supra_match:
                original_fn = int(supra_match.group(1))

                parts = re.split(r"\bsupra\b", chunk, flags=re.IGNORECASE)
                prefix_text = (
                    parts[0].lower().strip().rstrip(",") if parts else ""
                )
                prefix_text = re.sub(
                    r"^[“\"'`\s]+|[”\"'`\s]+$", "", prefix_text
                )

                matching_keys = [
                    k for k in source_history.keys() if k[0] == original_fn
                ]
                found_supra_match = False

                for target_key in matching_keys:
                    orig_source_name = source_history[target_key][
                        "Source Name"
                    ].lower()

                    if prefix_text and any(
                        word in orig_source_name
                        for word in prefix_text.split()
                        if len(word) > 2
                    ):
                        source_history[target_key][
                            "Subsequent Footnotes"
                        ].append(str(fn_num))
                        last_source_key = target_key
                        found_supra_match = True
                        break

                if not found_supra_match and matching_keys:
                    fallback_key = matching_keys[0]
                    source_history[fallback_key]["Subsequent Footnotes"].append(
                        str(fn_num)
                    )
                    last_source_key = fallback_key

                continue

            # --- CASE SHORT CITE CHECK ---
            found_short = False
            for name, origin_key in short_name_map.items():
                reporter_str = reporter_map.get(origin_key, "___")
                if (
                    name.lower() in chunk.lower()
                    and reporter_str.lower() in chunk.lower()
                ):
                    source_history[origin_key]["Subsequent Footnotes"].append(
                        str(fn_num)
                    )
                    last_source_key = origin_key
                    found_short = True
                    break
            if found_short:
                continue

            # --- NEW SOURCE ANALYSIS ---
            url_match = re.search(r"https?://[^\s\]]+", chunk)
            url = url_match.group(0) if url_match else ""

            src_type = identify_source_type(chunk, url)

            if src_type == "Case":
                name_match = re.match(r"^([^,v]+)\s+v\.", chunk)
                reporter_match = re.search(
                    r"\d+\s+[A-Z][\w\.\s\’\']+\s+\d+", chunk
                )
                if name_match:
                    short_name_map[name_match.group(1).strip()] = current_key
                if reporter_match:
                    rep_part = re.sub(
                        r"^\d+\s+|\s+\d+$", "", reporter_match.group(0)
                    ).strip()
                    reporter_map[current_key] = rep_part

            source_history[current_key] = {
                "Source Name": chunk,
                "Source Type": src_type,
                "Footnote": fn_num,
                "Subsequent Footnotes": [],
                "Link/Location": url,
            }
            last_source_key = current_key

    final_data = [src for src in source_history.values()]
    for src in final_data:
        src["Subsequent Footnotes"] = ", ".join(
            sorted(set(src["Subsequent Footnotes"]), key=int)
        )

    df = pd.DataFrame(final_data)
    df[
        [
            "Source Name",
            "Source Type",
            "Footnote",
            "Subsequent Footnotes",
            "Link/Location",
        ]
    ].to_excel(excel_file, index=False)


if __name__ == "__main__":
    # Hide the main Tkinter root window
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("Opening file browser window...")

    # Open standard Windows file selection prompt (ALLOWS MULTIPLE SELECTION)
    word_paths = filedialog.askopenfilenames(
        title="Select Word Document(s)",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
    )

    if not word_paths:
        print("No files selected. Operation canceled.")
        sys.exit()

    total_files = len(word_paths)
    print(f"\nSelected {total_files} file(s) for processing.\n" + "-" * 40)

    for index, word_path in enumerate(word_paths, start=1):
        filename = os.path.basename(word_path)
        print(f"[{index}/{total_files}] Processing: {filename}...")

        try:
            base_name = os.path.splitext(word_path)[0]
            excel_path = f"{base_name}_extracted.xlsx"

            extract_bluebook_data(word_path, excel_path)
            print(f"    ✓ Saved: {os.path.basename(excel_path)}")

        except Exception as e:
            print(f"    ✗ Error processing {filename}: {e}")

    print("\n" + "=" * 40)
    print("All tasks completed!")
    input("\nPress Enter to exit...")
