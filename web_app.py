import streamlit as st
from pypdf import PdfReader, PdfWriter
import re
import io
import zipfile

# --- CONFIGURATION ---
st.set_page_config(page_title="PDF Splitter", layout="centered")

def clean_filename(s):
    return re.sub(r'[\\/*?:"<>|]', "", s)

def extract_info(text, page_num):
    doc_no = f"Unknown_{page_num}"
    ack_date = "Unknown_Date"
    
    # Extract Document No
    match_doc = re.search(r"(?:Document|Doc|Ref)\s*No\.?[:\s]*([A-Za-z0-9-/]+)", text, re.IGNORECASE)
    if match_doc: doc_no = match_doc.group(1).strip()

    # Extract Date
    match_date = re.search(r"(?:Ack|Date)[:\s]*(\d{2}[-./]\d{2}[-./]\d{4})", text, re.IGNORECASE)
    if match_date: ack_date = match_date.group(1).replace("/", "-").replace(".", "-")

    return clean_filename(doc_no), clean_filename(ack_date)

# --- UI ---
st.title("✂️ PDF Splitter & Renamer")
st.write("Upload a master PDF containing multiple bills. This tool will split them and rename them automatically.")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    if st.button("Start Processing"):
        # Create a progress bar
        my_bar = st.progress(0)
        status_text = st.empty()
        
        # Prepare a ZIP file in memory (servers can't save to C:/Users/...)
        zip_buffer = io.BytesIO()
        
        try:
            reader = PdfReader(uploaded_file)
            total_pages = len(reader.pages)
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for i, page in enumerate(reader.pages):
                    text = page.extract_text() or ""
                    doc, date = extract_info(text, i+1)
                    
                    # Target Filename: 123 _ 12-05-2024.pdf
                    pdf_filename = f"{doc} _ {date}.pdf"
                    
                    # Handle duplicates inside the zip
                    count = 1
                    base_name = pdf_filename
                    while pdf_filename in zip_file.namelist():
                        pdf_filename = f"{doc} _ {date} ({count}).pdf"
                        count += 1
                    
                    # Write page to temporary PDF buffer
                    page_buffer = io.BytesIO()
                    writer = PdfWriter()
                    writer.add_page(page)
                    writer.write(page_buffer)
                    
                    # Add to ZIP
                    zip_file.writestr(pdf_filename, page_buffer.getvalue())
                    
                    # Update UI
                    percent = int(((i + 1) / total_pages) * 100)
                    my_bar.progress(percent)
                    status_text.text(f"Processing page {i+1}/{total_pages}")

            st.success("✅ Processing Complete!")
            
            # Create Download Button
            st.download_button(
                label="📥 Download All Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Split_Documents.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"An error occurred: {e}")