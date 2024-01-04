from PIL import Image
import pytesseract
import streamlit as st
import fitz


# Path to the Tesseract OCR executable (change if necessary)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.title("OCR with Streamlit")

uploaded_file = st.file_uploader("Please upload a photo or a Pdf!", accept_multiple_files=False, type=["jpg", "jpeg", "png", "pdf"])

if uploaded_file:
    if uploaded_file.type == "application/pdf":
        # For PDF files, convert each page to an image and perform OCR
        with fitz.open(stream=uploaded_file.read(), filetype="pdf") as pdf_reader:
            num_pages = pdf_reader.page_count

            for page_num in range(num_pages):
                page = pdf_reader[page_num]

                # # # Convert the page to a Pixmap
                # pixmap = page.get_pixmap()

                matrix = fitz.Matrix(2, 2)  # You can adjust the scaling factor
                pixmap = page.get_pixmap(matrix=matrix)

                # Convert the Pixmap to a PIL Image
                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)

                st.image(image)

                # Perform OCR on the current page
                text = pytesseract.image_to_string(image)

                # Display the recognized text for each page
                st.write(f"Page {page_num + 1} OCR Result:")
                st.write(text)

    else:
        image = Image.open(uploaded_file)

        st.image(image)

        # Perform OCR
        text = pytesseract.image_to_string(image)

        # Print the recognized text
        st.write(text)