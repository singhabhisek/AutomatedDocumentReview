import datetime
import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
import xml.etree.ElementTree as ET
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import re
from datetime import datetime  # Correct import

# ✅ Set Streamlit to Full-Width Mode
# st.set_page_config(layout="wide", page_title="PPT Validation App", page_icon="📊")


st.markdown(
    """
    <style>
        .block-container { padding-top: 3.0rem; } /* Reduce top padding */
    </style>
    """,
    unsafe_allow_html=True
)


# Apply custom CSS for styling
st.markdown(
    """
    <style>
        .custom-subheader {
            font-size: 22px !important;
            font-weight: bold;
            color: #333;
        }
    </style>
    """, 
    unsafe_allow_html=True
)


# SAMPLE_RELEASES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SampleReleases.xlsx')
# Define the path for the config file (assumes it's in a "config" folder next to the script)
CONFIG_FOLDER = os.path.join(os.getcwd(), "config")
CONFIG_FILE = os.path.join(CONFIG_FOLDER, "config.xlsx")
SHEET_NAME = "performance_testing_strategy"  # Assuming a single sheet for all word documents
SAMPLE_RELEASES_FILE = os.path.join(CONFIG_FOLDER,'SampleReleases.xlsx')
temp_dir = os.path.join(os.getcwd(), "temp")  # Create 'temp' folder path



# Load existing sample releases
def load_sample_releases():
    if os.path.exists(SAMPLE_RELEASES_FILE):
        return pd.read_excel(SAMPLE_RELEASES_FILE)
    else:
        st.error("SampleReleases.xlsx not found. Please place the file in the correct location.")
        return pd.DataFrame()

# Extract text from named shapes in a slide
def extract_named_shapes(zip_path, slide_number):
    shape_texts = {}
    slide_file = f"ppt/slides/slide{slide_number}.xml"

    with zipfile.ZipFile(zip_path, "r") as pptx_zip:
        if slide_file in pptx_zip.namelist():
            with pptx_zip.open(slide_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

                for sp in root.findall(".//p:sp", namespaces=ns):
                    name_elem = sp.find(".//p:nvSpPr/p:cNvPr", namespaces=ns)
                    if name_elem is not None and "name" in name_elem.attrib:
                        shape_name = name_elem.attrib["name"]
                        text_elem = sp.findall(".//a:t", namespaces=ns)
                        text_content = " ".join([t.text for t in text_elem if t.text])
                        shape_texts[shape_name] = text_content

    return shape_texts

# Check if embedded Excel files exist
def check_embedded_excel(zip_path):
    with zipfile.ZipFile(zip_path, "r") as pptx_zip:
        return any(f.startswith("ppt/embeddings/") and f.endswith(".xlsx") for f in pptx_zip.namelist())


def extract_tables_from_slide(zip_path, slide_number):
    """
    Extracts tables from a given slide in the PowerPoint (.pptx) file.

    Args:
        zip_path (str): Path to the PPTX file (as a zip archive).
        slide_number (int): The slide number to extract tables from.

    Returns:
        list: A list of tables, where each table is a list of rows, and each row is a list of cell values.
    """
    tables = []
    
    slide_path = f"ppt/slides/slide{slide_number}.xml"  # Locate the slide XML file
    
    with zipfile.ZipFile(zip_path, 'r') as pptx:
        if slide_path not in pptx.namelist():
            return tables  # If slide XML is missing, return an empty list
        
        slide_xml = pptx.read(slide_path)
        root = ET.fromstring(slide_xml)

        # Define namespaces to search for table elements
        ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
              'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

        # Find all tables in the slide
        for table in root.findall(".//a:tbl", ns):
            extracted_table = []
            
            # Find all rows in the table
            for row in table.findall(".//a:tr", ns):
                extracted_row = []
                
                # Find all cells in the row
                for cell in row.findall(".//a:tc", ns):
                    # Extract text from each cell
                    text_elem = cell.find(".//a:t", ns)
                    extracted_row.append(text_elem.text.strip() if text_elem is not None else "")
                
                extracted_table.append(extracted_row)  # Add row to table
            
            tables.append(extracted_table)  # Add table to list of tables
    
    return tables

def extract_embedded_files(zip_path, slide_number, output_dir="embedded_files"):
    """
    Extracts embedded files (Excel, CSV, etc.) from a specific slide in a PowerPoint file.

    :param zip_path: Path to the PPTX zip archive.
    :param slide_number: The slide number to check for embedded files.
    :param output_dir: Directory to store extracted files.
    :return: List of extracted file paths.
    """
    extracted_files = []
    os.makedirs(output_dir, exist_ok=True)  # Ensure directory exists

    with zipfile.ZipFile(zip_path, 'r') as pptx_zip:
        # Extract ALL embedded files from ppt/embeddings/
        for file_name in pptx_zip.namelist():
            if file_name.startswith("ppt/embeddings/"):  # Could be .xlsx, .csv, .bin
                extracted_path = os.path.join(output_dir, os.path.basename(file_name))
                with pptx_zip.open(file_name) as source, open(extracted_path, "wb") as target:
                    target.write(source.read())
                extracted_files.append(extracted_path.lower().strip())

        # Check slide-specific relationships for embedded files
        slide_rels_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        slide_embedded_files = []

        if slide_rels_path in pptx_zip.namelist():
            with pptx_zip.open(slide_rels_path) as rels_file:
                rels_content = rels_file.read().decode("utf-8")

                # Find all embedded references (may be .xlsx, .bin, .csv)
                embedded_refs = re.findall(r'Target="(../embeddings/[^"]+)"', rels_content)
                for ref in embedded_refs:
                    embedded_filename = os.path.basename(ref)
                    matched_file = os.path.normpath(os.path.join(output_dir, embedded_filename)).lower().strip()  # ✅ Normalize path

                    # print(f"🔍 Checking Embedded File: {embedded_filename}")  
                    # print(f"➡ Matched Path: {matched_file}")  
                    # print(f"✅ Extracted Files: {extracted_files}")  

                    # Compare after ensuring lowercase + consistent path format
                    if matched_file in extracted_files:
                        # print("✅ Match Found! Adding to results.")  
                        slide_embedded_files.append(matched_file)

    # print(slide_embedded_files)
    return slide_embedded_files if slide_embedded_files else extracted_files

def get_total_slides(pptx_path):
    """Extracts the total number of slides from a PowerPoint file."""
    with zipfile.ZipFile(pptx_path, 'r') as pptx_zip:
        slide_files = [f for f in pptx_zip.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")]
        return len(slide_files)
    

# Function to get a slide's display name based on its extracted title
def get_slide_display_name(slide_number, slide_shapes):
    """Extract slide title and format slide name dynamically"""
    default_names = {1: "Title Page", 2: "Observations Slide"}  # Custom names for Slide 1 & 2
    extracted_title = slide_shapes.get("Title", "").strip()  # Extract the title text

    if slide_number in default_names:
        return f"Slide {slide_number} - {default_names[slide_number]}"
    elif extracted_title:
        return f"Slide {slide_number} - {extracted_title}"  # Use extracted title
    else:
        return f"Slide {slide_number}"  # Default fallback if no title


def extract_text_from_slide(zip_path, slide_number):
    with zipfile.ZipFile(zip_path, 'r') as pptx:
        slide_path = f"ppt/slides/slide{slide_number}.xml"

        if slide_path not in pptx.namelist():
            return ""

        with pptx.open(slide_path) as slide_file:
            tree = ET.parse(slide_file)
            root = tree.getroot()

            # PowerPoint uses the following namespace for drawing text
            namespace = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
            }

            # Find all text elements
            text_elements = root.findall('.//a:t', namespace)
            all_text = " ".join([elem.text for elem in text_elements if elem.text])

            return all_text.strip()
           

def normalize_text(text):
    if text is None:
        return ""
    text = str(text).strip()  # Convert to lowercase & strip spaces
    text = re.sub(r"\s*[\-–—]\s*", "-", text)  # Replace different dashes with a standard hyphen
    text = re.sub(r"\s+", " ", text)  # Normalize spaces
    return text

# Main validation function
def validate_ppt(zip_path, checklist_row):
    total_slides = get_total_slides(zip_path)
    results = {}

    # Extract all shape text from Slide 1 (unnamed)
    slide1_shapes = extract_named_shapes(zip_path, 1)
    project_details_text = " ".join(
        shape_text.strip()
        for shape_text in slide1_shapes.values()
        if isinstance(shape_text, str)
    ).strip()

    # Required fields to validate
    required_fields = ["Enterprise Release ID", "Project Name", "Release", "Application ID", "Business Application", "Project ID"]

    # Patterns with fallback support
    patterns = {
        "Project Name": r"(?:project name\s*[:\-–]?\s*)?([^\n\r]+?(?=\s*(performance test report|enterprise release id|rlse|rlsea|application name|app id|appid)))",
        "Enterprise Release ID": r"enterprise\s+release\s+id\s*[:\-–]?\s*([^\s\n\r]+)",
        "Project ID": r"prj[-\s]?(\w+)",
        "Release": r"(rlse[a-z]*\d+)",
        "Business Application": r"(?:application|app)\s+names?\s*[:\-–]?\s*(.*?)(?=\s*\(?appid\s*[-–]?\s*[a-zA-Z]?\s*\d+\)?)",
        "Application ID": r"appid\s*[-–]?\s*[a-zA-Z]?\s*(\d+)"
    }

    # Extract values using regex
    extracted_values = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, project_details_text, re.IGNORECASE)
        if match:
            extracted_values[key] = normalize_text(match.group(1).strip())

    # Slide 1 validation comparison
    slide1_results = {}
    for key, expected_value in checklist_row.items():
        if key not in required_fields:
            continue
        expected_value = normalize_text(str(expected_value).strip())
        extracted_value = extracted_values.get(key, None)

        if extracted_value is None:
            slide1_results[key] = f"🚫 Missing (Expected: {expected_value})"
        elif key == "Application ID":
            if extracted_value == expected_value.replace("APP-", "").lower():
                slide1_results[key] = f"✅ Matched (Expected: {expected_value}, Found: APP-{extracted_value})"
            else:
                slide1_results[key] = f"❌ Not Matched (Expected: {expected_value}, Found: APP-{extracted_value})"
        elif key == "Business Application":
            def normalize_app_name_for_comparison(expected, found):
                expected_clean = expected.lower().strip()
                found_clean = found.lower().strip()

                # Allow skipping variations of (DEMO) only in expected (if user put extra info, but not in PPT)
                demo_pattern = r"\(\s*demo\s*\)"
                if re.search(demo_pattern, expected_clean) and not re.search(demo_pattern, found_clean):
                    expected_clean = re.sub(demo_pattern, "", expected_clean).strip()

                return expected_clean, found_clean

            exp_clean, found_clean = normalize_app_name_for_comparison(expected_value, extracted_value)

            if exp_clean == found_clean:
                slide1_results[key] = f"✅ Matched (Expected: {expected_value}, Found: {extracted_value})"
            else:
                slide1_results[key] = f"❌ Not Matched (Expected: {expected_value}, Found: {extracted_value})"
        elif extracted_value.lower() == expected_value.lower():
            slide1_results[key] = f"✅ Matched (Expected: {expected_value}, Found: {extracted_value})"
        else:
            slide1_results[key] = f"❌ Not Matched (Expected: {expected_value}, Found: {extracted_value})"

    results["Slide 1"] = slide1_results

    # # === Slide 2 ===
    # slide2_shapes = extract_named_shapes(zip_path, 2)
    # slide2_tables = extract_tables_from_slide(zip_path, 2)
    # embedded_files = extract_embedded_files(zip_path, 2)

    # project_name = checklist_row.get("Project Name", "").strip().lower()
    # release_id = checklist_row.get("Enterprise Release ID", "").strip().lower()

    # slide2_title_text = normalize_text(slide2_shapes.get("Slide2Header", "").strip().lower())
    # project_name_lower = normalize_text(project_name.lower().strip())

    # match = re.search(rf"\b{re.escape(project_name_lower)}\b", slide2_title_text, re.IGNORECASE)
    # title_missing = match is None

    # slide2_summary_text = normalize_text(slide2_shapes.get("Slide2Summary", "").strip().lower())
    # summary_missing = []

    # release_pattern = normalize_text(re.escape(release_id.lower()))
    # release_match = re.search(fr"\b{release_pattern}\b", slide2_summary_text)
    # if not release_match:
    #     summary_missing.append(f"Release ID '{release_id.upper()}' Not Found")

    # project_pattern = re.escape(normalize_text(project_name.lower()))
    # project_match = re.search(fr"\b{project_pattern}\b", slide2_summary_text)
    # if not project_match:
    #     summary_missing.append(f"Project Name '{project_name.title()}' Not Found")

    # # Table validation
    # table_valid = False
    # date_row_valid = False

    # for table in slide2_tables:
    #     for row_index, row in enumerate(table):
    #         if row_index == 0:
    #             continue

    #         first_column_text = row[0].strip().lower() if row and row[0] else ""
    #         second_column_text = str(row[1]).strip() if len(row) > 1 else ""
    #         third_column_text = str(row[2]).strip() if len(row) > 2 else ""

    #         if first_column_text in ["load test", "endurance test", "load", "endurance"]:
    #             table_valid = True

    #         if len(second_column_text) > 0 and len(third_column_text) > 0:
    #             date_row_valid = True

    #         if table_valid and date_row_valid:
    #             break
    #     if table_valid and date_row_valid:
    #         break

    # table_validation_result = (
    #     "✅ Valid" if table_valid and date_row_valid
    #     else "❌ Found the Test Type, however, dates are missing." if table_valid
    #     else "❌ Test Type is missing. Please validate and correct the Execution Details table."
    # )

    # has_embedded_excel = any(file.lower().endswith((".xlsm", ".xlsx", ".xls", ".csv")) for file in embedded_files)

    # results["Slide 2"] = {
    #     "Title Validation": "✅ Valid" if not title_missing else "❌ Missing or Incorrect Project Name",
    #     "Summary Validation": "✅ Valid" if not summary_missing else f"❌  {', '.join(summary_missing)}",
    #     "Table Validation": table_validation_result,
    #     "Embedded Excel": "✅ Found" if has_embedded_excel else "❌ No Excel file found",
    # }

    # === Slide 2 ===
    slide2_text = extract_text_from_slide(zip_path, 2)
    slide2_tables = extract_tables_from_slide(zip_path, 2)
    embedded_files = extract_embedded_files(zip_path, 2)

    project_name = checklist_row.get("Project Name", "").strip().lower()
    release_id = checklist_row.get("Enterprise Release ID", "").strip().lower()

    # 🔹 Normalize full text for search
    slide2_text_normalized = normalize_text(slide2_text.lower())

    # === Title Validation (Search for Project Name in entire text)
    project_name_lower = normalize_text(project_name)
    match = re.search(rf"\b{re.escape(project_name_lower)}\b", slide2_text_normalized, re.IGNORECASE)
    title_missing = match is None

    # === Summary Validation
    summary_missing = []

    # 🔹 Validate Release ID presence
    release_pattern = normalize_text(re.escape(release_id))
    release_match = re.search(fr"\b{release_pattern}\b", slide2_text_normalized)
    if not release_match:
        summary_missing.append(f"Release ID '{release_id.upper()}' Not Found")

    # 🔹 Validate Project Name presence
    project_pattern = re.escape(normalize_text(project_name))
    project_match = re.search(fr"\b{project_pattern}\b", slide2_text_normalized)
    if not project_match:
        summary_missing.append(f"Project Name '{project_name.title()}' Not Found")

    # === Table validation (same as before)
    table_valid = False
    date_row_valid = False

    for table in slide2_tables:
        for row_index, row in enumerate(table):
            if row_index == 0:
                continue

            first_column_text = row[0].strip().lower() if row and row[0] else ""
            second_column_text = str(row[1]).strip() if len(row) > 1 else ""
            third_column_text = str(row[2]).strip() if len(row) > 2 else ""

            if first_column_text in ["load test", "endurance test", "load", "endurance"]:
                table_valid = True

            if len(second_column_text) > 0 and len(third_column_text) > 0:
                date_row_valid = True

            if table_valid and date_row_valid:
                break
        if table_valid and date_row_valid:
            break

    table_validation_result = (
        "✅ Valid" if table_valid and date_row_valid
        else "❌ Found the Test Type, however, dates are missing." if table_valid
        else "❌ Test Type is missing. Please validate and correct the Execution Details table."
    )

    # === Embedded Excel Check
    has_embedded_excel = any(file.lower().endswith((".xlsm", ".xlsx", ".xls", ".csv")) for file in embedded_files)

    # === Final Slide 2 Validation Result
    results["Slide 2"] = {
        "Title Validation": "✅ Valid" if not title_missing else "❌ Missing or Incorrect Project Name",
        "Summary Validation": "✅ Valid" if not summary_missing else f"❌  {', '.join(summary_missing)}",
        "Table Validation": table_validation_result,
        "Embedded Excel": "✅ Found" if has_embedded_excel else "❌ No Excel file found",
    }

    # === Slide 3+ Validation ===
    # for slide_number in range(3, total_slides+1):
    #     slide_shapes = extract_named_shapes(zip_path, slide_number)
    #     extracted_title = slide_shapes.get("Title", "").strip()
    #     extracted_observations = slide_shapes.get("Observations", "").strip()

    #     results[f"Slide {slide_number}"] = {
    #         "Title Found": "✅ Yes" if extracted_title else "❌ No",
    #         "Observations Found": "✅ Yes" if extracted_observations else "❌ No",
    #     }

    # return results

    

    title_keywords = ["title", "chart", "graph", "metrics", "summary", "observations", "overview"]
    observation_keywords = ["observation", "issue", "finding", "remarks", "note", "conclusion", "summary"]

    for slide_number in range(3, total_slides + 1):
        slide_text = extract_text_from_slide(zip_path, slide_number).strip().lower()
        lines = [line.strip() for line in slide_text.splitlines() if line.strip()]
        
        extracted_title = ""
        for line in lines[:3]:
            if any(keyword in line for keyword in title_keywords):
                extracted_title = line
                break

        if not extracted_title and lines:
            extracted_title = lines[0]

        extracted_observations = any(keyword in slide_text for keyword in observation_keywords)
        
        results[f"Slide {slide_number}"] = {
            "Title Found": "✅ Yes" if extracted_title else "❌ No",
            "Observations Found": "✅ Yes" if extracted_observations else "❌ No",
            # "Extracted Title": extracted_title
        }
    return results

# # Validate PowerPoint against selected row
# def validate_ppt(zip_path, checklist_row):
#     total_slides = get_total_slides(zip_path)
#     results = {}

#     # Extract named shapes from Slide 1
#     slide1_shapes = extract_named_shapes(zip_path, 1)

#     # 🔹 Extract entire Project Details text block
#     project_details_text = slide1_shapes.get("Slide1ProjectDetails", "").strip()

#     # print("project_details_text: " + project_details_text)

#     # print("Normalized project_details_text:", repr(project_details_text))

#     required_fields = ["Enterprise Release ID", "Project Name", "Release", "Application ID", "Application Name", "Project ID"]  # Can be modified anytime

#     patterns = {
#     "Project Name": r"project\s*name\s*[:\-–]?\s*([\w\s\(\)\[\]\-–\.]+?)(?=\s*\b(release|project id|enterprise|application name|application id)\b|$)",
#     "Release": r"release\s*[:\-–]?\s*([\w\.\-]+)(?=\s*\b(project|application name|application id|enterprise release id|$)\b)",
#     "Project ID": r"project\s*id\s*[:\-–]?\s*([\w\-]+)(?=\s*\b(enterprise|application name|application id)\b|$)",
#     "Enterprise Release ID": r"enterprise\s+release\s+id\s*[:\-–]?\s*([\w\.\-\s]+)(?=\s*\b(application|application id)\b|$)",
#     "Application Name": r"application\s*name\s*[:\-–]?\s*([\w\d\s\(\)\[\]\-–]+?)(?=\s*\b(application id)\b|$)",
#     "Application ID": r"application id\s*[:\-–]?\s*(?:app-?id-?)?([\w\d\-]+)\b"
#     # "Application ID": r"application id\s*[:\-–]?\s*app-?([\w\d\-]+)"
# }   
#     extracted_values = {}
#     for key, pattern in patterns.items():
#         match = re.search(pattern, project_details_text, re.IGNORECASE)
#         if match:
#             extracted_values[key] = normalize_text(match.group(1).strip())
    
#     # print(extracted_values)
#     # 🔹 Compare extracted values with expected values from checklist
#     slide1_results = {}
#     for key, expected_value in checklist_row.items():
#         # print(key)
#         if key not in required_fields:
#             continue  # Skip fields that are not required
#         expected_value = normalize_text(str(expected_value).strip())
#         # print ("expected:" + expected_value)
#         # print (extracted_values)
#         extracted_value = extracted_values.get(key, None)

#         if extracted_value is None:
#             slide1_results[key] = f"🚫 Missing (Expected: {expected_value})"
#         elif key == "Application ID":  # Special handling for Application ID (removing "APP-")
#             if extracted_value == expected_value.replace("APP-", ""):
#                 slide1_results[key] = "✅ Matched"
#             else:
#                 slide1_results[key] = f"❌ Not Matched (Expected: {expected_value}, Found: APP-{extracted_value})"
#         elif extracted_value.lower() == expected_value.lower():
#             slide1_results[key] = "✅ Matched"
#         else:
#             slide1_results[key] = f"❌ Not Matched (Expected: {expected_value}, Found: {extracted_value})"

#     # ✅ Store Slide 1 validation results
#     results["Slide 1"] = slide1_results

    
#     # return results
#     # Slide 2 Validation
#     slide2_shapes = extract_named_shapes(zip_path, 2)
#     slide2_tables = extract_tables_from_slide(zip_path, 2)
#     embedded_files = extract_embedded_files(zip_path, 2)

#     # Fetch Project ID & Release ID from checklist
#     project_name = checklist_row.get("Project Name", "").strip().lower()
#     release_id = checklist_row.get("Enterprise Release ID", "").strip().lower()

#     # ✅ Validate Slide2Title (Check if Project ID is present)
#     # ✅ Extract Slide 2 Title & Convert to Lowercase
#     slide2_title_text = normalize_text(slide2_shapes.get("Slide2Header", "").strip().lower())
#     print (slide2_title_text)
#     project_name_lower = normalize_text(project_name.lower().strip())  # Normalize for comparison
#     # print(project_name_lower)

#     # ✅ Use Regex to Find "Project Y" Anywhere in the Title
#     match = re.search(rf"\b{re.escape(project_name_lower)}\b", slide2_title_text, re.IGNORECASE)

#     # ✅ If Project Name is Found in the Title, It’s Valid
#     title_missing = match is None  # If match is None, it means Project Name was NOT found

#     # print("Extracted Project Name Found:", match.group(0) if match else "Not Found")
#     # print("Expected Project Name:", project_name_lower)
#     # print("Title Validation Result:", "✅ Valid" if not title_missing else "❌ Missing Project Name")


#     # ✅ Validate Slide2Summary (Check for both Project ID & Release ID)
#     # ✅ Validate Slide2Summary (Check for Project Name & Release ID in any order)
#     slide2_summary_text = normalize_text(slide2_shapes.get("Slide2Summary", "").strip().lower())
#     # print ("==========" + slide2_summary_text)
#     summary_missing = []

    
#     # 🔹 Directly check for Release ID in text (from config)
#     release_pattern = normalize_text(re.escape(release_id.lower()))  # Escape special characters if any
#     # print ("93-------------------" + release_pattern)
#     release_match = re.search(fr"\b{release_pattern}\b", slide2_summary_text)
#     # print(slide2_summary_text)

#     # ✅ Validate Release ID presence
#     if release_match:
#         extracted_release_id = release_id  # Since it's an exact match
#     else:
#         summary_missing.append(f"Release ID '{release_id.upper()}' Not Found")

#     # print("*******"+ project_name)
#     # 🔹 Directly check for Project Name in text (from config)
#     project_pattern = re.escape(normalize_text(project_name.lower()))  # Escape special characters if any
#     project_match = re.search(fr"\b{project_pattern}\b", slide2_summary_text)

#     # ✅ Validate Project Name presence
#     if project_match:
#         extracted_project_name = project_name  # Since it's an exact match
#     else:
#         project_name = project_name.title();
#         summary_missing.append(f"Project Name '{project_name}' Not Found")

#     # 🔹 Print Debug Information (Optional)
#     print("Extracted Release ID:", release_id if release_match else "Not Found")
#     print("Extracted Project Name:", project_name if project_match else "Not Found")
#     print("Validation Summary:", summary_missing if summary_missing else "✅ Valid")

#     # ✅ Validate Table (Ensure at least one row contains "Load" or "Endurance" in first column)
#     table_valid = False
#     date_row_valid = False

#     for table in slide2_tables:
#         for row_index, row in enumerate(table):
#             if row_index == 0:
#                 continue  # Skip header row

#             first_column_text = row[0].strip().lower() if row and row[0] else ""
#             second_column_text = str(row[1]).strip() if len(row) > 1 else ""
#             third_column_text = str(row[2]).strip() if len(row) > 2 else ""

#             # ✅ Condition 1: Check if first column contains "Load" or "Endurance"
#             if first_column_text.lower() in ["load test", "endurance test", "load", "endurance"]:
#                 table_valid = True

#             date_row_valid = False  # 🔹 Reset before validation
#             # ✅ Condition 2: Ensure both second & third columns contain valid dates
#             if len(second_column_text)>0 and len(third_column_text)>0:
#                 date_row_valid = True
#                 # try:
#                 #     datetime.strptime(second_column_text, "%d/%m/%Y")  # Adjust format as needed
#                 #     datetime.strptime(third_column_text, "%d/%m/%Y")
#                 #     date_row_valid = True
#                 #     print("Dates are present")
#                 # except ValueError:
#                 #     date_row_valid = False  # If parsing fails, mark it invalid
#             # ✅ If both conditions met, exit loop early
#             if table_valid and date_row_valid:
#                 break

#         if table_valid and date_row_valid:
#             break

#     # ✅ Final Validation Result with Detailed Messages
#     if table_valid and date_row_valid:
#         table_validation_result = "✅ Valid"
#     elif table_valid and not date_row_valid:
#         table_validation_result = "❌ Found the Test Type, however, dates are missing."
#     else:
#         table_validation_result = "❌ Test Type is missing. Please validate and correct the Execution Details table."


#     # ✅ Validate Embedded Excel File Presence
#     has_embedded_excel = any(file.lower().endswith((".xlsm", ".xlsx", ".xls", ".csv")) for file in embedded_files)

#     # ✅ Store validation results
#     results["Slide 2"] = {
#         "Title Validation": "✅ Valid" if not title_missing else "❌ Missing or Incorrect Project Name",
#         "Summary Validation": "✅ Valid" if not summary_missing else f"❌  {', '.join(summary_missing)}",
#         "Table Validation": table_validation_result,
#         "Embedded Excel": "✅ Found" if has_embedded_excel else "❌ No Excel file found",
#         # "Extracted Shapes": slide2_shapes
#     }

    
#     # Validate Slide 3 onwards for "Observations" shape
#     for slide_number in range(3, total_slides+1):  # Check up to slide 10
#         slide_shapes = extract_named_shapes(zip_path, slide_number)

#         # Extract possible title and observation fields
#         extracted_title = slide_shapes.get("Title", "").strip()
#         extracted_observations = slide_shapes.get("Observations", "").strip()

#         results[f"Slide {slide_number}"] = {
#             "Title Found": "✅ Yes" if extracted_title else "❌ No",
#             "Observations Found": "✅ Yes" if extracted_observations else "❌ No",
#             # "Extracted Shapes": slide_shapes
#         }

#     return results

# Generate validation report in Excel
def generate_excel_report(validation_results):
    output = BytesIO()  # ✅ Create BytesIO buffer

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if validation_results and len(validation_results) > 0:  
            for slide, result in validation_results.items():
                df = pd.DataFrame.from_dict(result, orient='index', columns=["Validation Result"])
                df.to_excel(writer, sheet_name=slide)
        else:
            # ✅ Ensure at least one sheet is present
            df = pd.DataFrame([["No validation results found"]], columns=["Message"])
            df.to_excel(writer, sheet_name="Summary")

        writer.book.active = 0  # ✅ Ensure the first sheet is active

    writer.close()  # ✅ Explicitly close the writer

    output.seek(0)  # ✅ Reset buffer position
    return output

# Streamlit UI
# st.title("Test Report Validation Application - PPT Format")
st.title("📑 Test Report Validation Application - PPT Format")

# Load Sample Releases
# sample_releases_df = load_sample_releases()

# Load Sample Releases
sample_releases_df = pd.read_excel(SAMPLE_RELEASES_FILE)

st_col1, st_col2 = st.columns([0.8, 0.2])

# Add a search bar for filtering
with st_col2:
    search_text = st.text_input("", placeholder="🔍 Search...")
    escaped_search_text = re.escape(search_text)  # Escape special regex characters

# Filter the DataFrame dynamically based on search text
if search_text:
    # Convert all columns to string type for search and filter
    sample_releases_df_filtered = sample_releases_df[sample_releases_df.apply(lambda row: row.astype(str).str.contains(escaped_search_text, case=False, na=False).any(), axis=1)]
else:
    sample_releases_df_filtered = sample_releases_df

# Display the table for selection
st.subheader("📋 Select a Release for Validation")
gb = GridOptionsBuilder.from_dataframe(sample_releases_df)
gb.configure_selection('single', use_checkbox=True)
grid_options = gb.build()

grid_response = AgGrid(
    sample_releases_df_filtered,
    gridOptions=grid_options,
    update_mode=GridUpdateMode.VALUE_CHANGED | GridUpdateMode.SELECTION_CHANGED,
    height=300,
    fit_columns_on_grid_load=True
)

# File Upload Section
st.subheader("📂 Upload PowerPoint File")
uploaded_ppt = st.file_uploader("Upload PPTX File", type=["pptx"])
# print(uploaded_ppt)

# Button to trigger validation
# validation_results = None
if "validation_results" not in st.session_state:
    st.session_state.validation_results = None
selected_rows = grid_response.get('selected_rows', [])


# # ✅ Ensure row selection is handled correctly
# if isinstance(selected_rows, list) and selected_rows:  # Case 1: List of dictionaries
#     selected_row = selected_rows[0]  # Extract first row as dictionary
# elif isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:  # Case 2: DataFrame
#     selected_row = selected_rows.iloc[0].to_dict()  # Convert first row to dictionary
# else:
#     selected_row = None  # No selection

# # Process selected row
# if not selected_row:
#     st.warning("⚠️ No row selected. Please select a release.")

# 📌 Layout for Validate button & Export button side by side
col1, col2 = st.columns([0.8, 0.2])  # Adjust width ratio to align buttons properly


# if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
if uploaded_ppt is not None and isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
    selected_row_data = selected_rows.iloc[0]
    with col1:
        if st.button("✅ Validate PPT"):
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
                tmp_ppt.write(uploaded_ppt.read())
                tmp_ppt_path = tmp_ppt.name

            # Run validation
            # validation_results = validate_ppt(tmp_ppt_path, selected_row_data)
            st.session_state.validation_results = validate_ppt(tmp_ppt_path, selected_row_data)


            # Clean up temp file
            os.remove(tmp_ppt_path)

            # Display results
            st.subheader("✅ Validation Results")
            for slide, result in st.session_state.validation_results.items():
                # Extract the slide title from validation results
                extracted_title = result.get("Extracted Shapes", {}).get("Title", "").strip()

                # Assign a custom name for Slide 1 and Slide 2
                default_names = {
                    "Slide 1": "Title Page",
                    "Slide 2": "Observations Slide"
                }

                # Determine the final display name
                if slide in default_names:
                    slide_name = f"{slide} - {default_names[slide]}"
                elif extracted_title:
                    slide_name = f"{slide} - {extracted_title}"
                else:
                    slide_name = slide  # Fallback if no title is found

                # Display the updated slide name
                st.write(f"### {slide_name}")

                for key, value in result.items():
                    st.write(f"**{key}:** {value}")
            st.toast("✅ Validation Completed!")


# Generate & Download Excel Report
# if validation_results:
with col2:
    excel_data = generate_excel_report(st.session_state.validation_results) #validation_results)
    st.download_button(
        label="📥 Download Validation Report",
        data=excel_data,
        file_name="PPT_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
