import streamlit as st
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import os
import google.generativeai as genai
from openpyxl import load_workbook

# Set up the Google Generative AI client
os.environ["GEMINI_API_KEY"] = "AIzaSyDI2DelJZlGyXEPG3_b-Szo-ixRvaB0ydY"
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

# Define the prompt
prompt = ("the following is OCR extracted text from a single invoice PDF. "
          "Please use the OCR extracted text to give a structured summary. "
          "The structured summary should consider information such as PO Number, Invoice Number, Invoice Amount, Invoice Date, "
          "CGST Amount, SGST Amount, IGST Amount, Total Tax Amount, Taxable Amount, TCS Amount, IRN Number, Receiver GSTIN, "
          "Receiver Name, Vendor GSTIN, Vendor Name, Remarks and Vendor Code. If any of this information is not available or present, "
          "then NA must be denoted next to the value. Please do not give any additional information.")

# Streamlit UI elements
st.title("Invoice Data Extraction and Summarization")
st.write("Upload the Invoice PDFs")
uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

st.write("Upload the Local Master Excel File")
uploaded_excel = st.file_uploader("Choose an Excel file", type="xlsx")

# Processing if files are uploaded
if uploaded_files and uploaded_excel:
    # Save the uploaded Excel file
    excel_path = uploaded_excel.name
    with open(excel_path, "wb") as f:
        f.write(uploaded_excel.getbuffer())
    
    # Load the workbook and select the active sheet
    workbook = load_workbook(excel_path)
    worksheet = workbook.active

    def extract_text_from_pdf(pdf_path):
        doc = fitz.open(pdf_path)
        text_data = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text("text")
            text_data.append(text)
        return text_data

    def convert_pdf_to_images_and_ocr(pdf_path):
        images = convert_from_path(pdf_path)
        ocr_results = [pytesseract.image_to_string(image) for image in images]
        return ocr_results

    def combine_text_and_ocr_results(text_data, ocr_results):
        combined_results = []
        for text, ocr_text in zip(text_data, ocr_results):
            combined_results.append(text + "\n" + ocr_text)
        combined_text = "\n".join(combined_results)
        return combined_text

    def extract_parameters_from_response(response_text):
        def sanitize_value(value):
            return value.strip().replace('"', '').replace(',', '')

        parameters = {
            "PO Number": "NA",
            "Invoice Number": "NA",
            "Invoice Amount": "NA",
            "Invoice Date": "NA",
            "CGST Amount": "NA",
            "SGST Amount": "NA",
            "IGST Amount": "NA",
            "Total Tax Amount": "NA",
            "Taxable Amount": "NA",
            "TCS Amount": "NA",
            "IRN Number": "NA",
            "Receiver GSTIN": "NA",
            "Receiver Name": "NA",
            "Vendor GSTIN": "NA",
            "Vendor Name": "NA",
            "Remarks": "NA",
            "Vendor Code": "NA"
        }
        lines = response_text.splitlines()
        for line in lines:
            for key in parameters.keys():
                if key in line:
                    value = sanitize_value(line.split(":")[-1].strip())
                    parameters[key] = value
        return parameters

    for uploaded_file in uploaded_files:
        # Save the uploaded PDF file
        pdf_path = uploaded_file.name
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Process the PDF
        text_data = extract_text_from_pdf(pdf_path)
        ocr_results = convert_pdf_to_images_and_ocr(pdf_path)
        combined_text = combine_text_and_ocr_results(text_data, ocr_results)

        # Send the combined text to Google Generative AI
        input_text = f"{prompt}\n\n{combined_text}"
        generation_config = {
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 64,
            "max_output_tokens": 8192,
            "response_mime_type": "text/plain",
        }

        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=generation_config,
        )

        chat_session = model.start_chat(history=[])
        response = chat_session.send_message(input_text)
        parameters = extract_parameters_from_response(response.text)

        # Add data to the Excel file
        row_data = [parameters[key] for key in parameters.keys()]
        worksheet.append(row_data)

        # Print confirmation message
        st.write(f"Data from {pdf_path} has been successfully added to the Excel file")

    # Print the structured summaries after processing all PDFs
    for uploaded_file in uploaded_files:
        pdf_path = uploaded_file.name
        st.write(f"\n{pdf_path} Structured Summary:\n")
        for key, value in parameters.items():
            st.write(f"{key:20}: {value}")

    # Save the updated Excel file and provide a download link
    workbook.save(excel_path)
    st.write("Excel file has been updated.")
    st.download_button(label="Download Updated Excel", data=open(excel_path, "rb").read(), file_name=excel_path)

