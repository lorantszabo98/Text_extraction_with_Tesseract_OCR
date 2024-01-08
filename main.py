from PIL import Image
import pytesseract
import streamlit as st
import fitz
from fuzzysearch import find_near_matches
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os


def find_regex_with_fuzzy(reference_list, text, max_l_dist=2):
    match = None
    for reference in reference_list:
        lower_case_name_regex = reference.casefold()
        matches = find_near_matches(lower_case_name_regex, text, max_l_dist=max_l_dist)
        if len(matches) > 0:
            match = matches
            break

    return match


def extract_data(start_regex, end_regex, text, max_l_dist=2):

    start_word = find_regex_with_fuzzy(start_regex, text, max_l_dist)

    if start_word is None:
        print("Start regex not found.")
        return None  # or return some default value or handle the case

    start_index = start_word[0].end

    end_word = find_regex_with_fuzzy(end_regex, text, max_l_dist)

    if end_word is None:
        print("End regex not found.")
        return None  # or return some default value or handle the case

    end_index = end_word[0].start

    print(start_index, end_index)

    return text[start_index + 1:end_index:1]


def read_consecutive_numbers_from_index(text, start_index):
    # Ensure the start index is within the bounds of the text
    if start_index < 0 or start_index >= len(text):
        return None

    # Find the index of the first numeric digit after the start index
    first_digit_index = -1
    for i in range(start_index, len(text)):
        if text[i].isdigit():
            first_digit_index = i
            break

    # If no numeric digit is found, return None
    if first_digit_index == -1:
        return None

    # Initialize an empty string to store the consecutive numbers
    consecutive_numbers = ""

    # Iterate through characters starting from the first numeric digit found
    for char in text[first_digit_index:]:
        # Check if the character is a numeric digit
        if char.isdigit():
            consecutive_numbers += char
        else:
            # Break the loop if a non-numeric character is encountered
            break

    return consecutive_numbers


pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.title("OCR with Streamlit")

# Create an empty DataFrame to store the extracted data
excel_data = pd.DataFrame(columns=["Dátum", "Rendszám", "Szállítólevél száma", "Súly"])

uploaded_file = st.file_uploader("Please upload a photo or a Pdf!", accept_multiple_files=False, type=["jpg", "jpeg", "png", "pdf"])

if uploaded_file:
    if uploaded_file.type == "application/pdf":
        # For PDF files, convert each page to an image and perform OCR
        with fitz.open(stream=uploaded_file.read(), filetype="pdf") as pdf_reader:
            num_pages = pdf_reader.page_count

            for page_num in range(num_pages):
                page = pdf_reader[page_num]

                matrix = fitz.Matrix(2, 2)  # You can adjust the scaling factor
                pixmap = page.get_pixmap(matrix=matrix)

                # Convert the Pixmap to a PIL Image
                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)

                st.image(image)

                # Perform OCR on the current page
                text = pytesseract.image_to_string(image, lang="hun")

                # Display the recognized text for each page
                st.write(f"Page {page_num + 1} OCR Result:")

                # lower casing the text for the proper string matching
                lower_case_text = text.casefold()

                st.write(lower_case_text)

                regexes_for_date_start = ['teljesítés kelte:']
                regexes_for_date_end = ['szállítólevélszám']

                date = extract_data(regexes_for_date_start, regexes_for_date_end, lower_case_text, max_l_dist=2)

                regexes_for_transfer_number_start = ['szállítólevélszám:']
                regexes_for_transfer_number_end = ['jármű']

                transfer_number = extract_data(regexes_for_transfer_number_start, regexes_for_transfer_number_end, lower_case_text, max_l_dist=1)

                regexes_for_license_plate_start = ['jármű:']
                regexes_for_license_plate_end = ['mérlegelt bruttó']

                license_plate = extract_data(regexes_for_license_plate_start, regexes_for_license_plate_end, lower_case_text, max_l_dist=1)

                if license_plate is not None:

                    license_plate = license_plate.upper()
                    license_plate = license_plate.split(',')

                    st.write(f"Rendszám: {license_plate[0]}")

                regexes_for_mass_start = ['nettó tömeg:']
                # regexes_for_mass_end = ['kg']

                mass_start = find_regex_with_fuzzy(regexes_for_mass_start, lower_case_text, max_l_dist=2)
                start_index = mass_start[0].end

                mass = read_consecutive_numbers_from_index(lower_case_text, start_index)

                st.write(f"Dátum: {date}")
                st.write(f"Szállítólevél: {transfer_number}")
                st.write(f"Nettó tömeg: {mass}")

                license_plate = license_plate[0].strip()
                transfer_number = transfer_number.strip()

                # Error handling section
                if len(transfer_number) != 8:
                    st.write(f"Hiba a {page_num}. oldalon. (Nem megfelelő szállítólevélszám)")

                if len(license_plate) != 7:
                    st.write(f"Hiba a {page_num}. oldalon. (Nem megfelelő rendszám)")

                # Append the extracted data to the DataFrame
                excel_data = excel_data.append({
                    "Dátum": date.strip(),
                    "Rendszám": license_plate,
                    "Szállítólevél száma": transfer_number,
                    "Súly": mass
                }, ignore_index=True)

            # Save the DataFrame to an Excel file
            excel_file = "output.xlsx"
            excel_data.to_excel(excel_file, index=False)

    else:
        image = Image.open(uploaded_file)

        st.image(image)

        # Perform OCR
        text = pytesseract.image_to_string(image)

        # Print the recognized text
        st.write(text)