import os
import tkinter as tk
from tkinter import filedialog
from fuzzysearch import find_near_matches
import fitz
from PIL import Image
import pytesseract
import pandas as pd

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
    if start_index < 0 or start_index >= len(text):
        return None
    first_digit_index = -1
    for i in range(start_index, len(text)):
        if text[i].isdigit():
            first_digit_index = i
            break
    if first_digit_index == -1:
        return None
    consecutive_numbers = ""
    for char in text[first_digit_index:]:
        if char.isdigit():
            consecutive_numbers += char
        else:
            break
    return consecutive_numbers


def extract_data_from_pdf(pdf_path):
    excel_data = pd.DataFrame(columns=["Dátum", "Rendszám", "Szállítólevél száma", "Súly", "Hiba"])
    with fitz.open(pdf_path, filetype="pdf") as pdf_reader:
        num_pages = pdf_reader.page_count
        for page_num in range(num_pages):
            page = pdf_reader[page_num]
            matrix = fitz.Matrix(2, 2)
            pixmap = page.get_pixmap(matrix=matrix)
            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)

            text = pytesseract.image_to_string(image, lang="hun")
            lower_case_text = text.casefold()

            regexes_for_date_start = ['teljesítés kelte:']
            regexes_for_date_end = ['szállítólevélszám']
            date = extract_data(regexes_for_date_start, regexes_for_date_end, lower_case_text, max_l_dist=2)

            regexes_for_transfer_number_start = ['szállítólevélszám:']
            regexes_for_transfer_number_end = ['jármű']
            transfer_number = extract_data(regexes_for_transfer_number_start, regexes_for_transfer_number_end, lower_case_text, max_l_dist=1)

            regexes_for_license_plate_start = ['jármű:']
            regexes_for_license_plate_end = ['mérlegelt bruttó']
            license_plate = extract_data(regexes_for_license_plate_start, regexes_for_license_plate_end, lower_case_text, max_l_dist=1)

            regexes_for_mass_start = ['nettó tömeg:']
            mass_start = find_regex_with_fuzzy(regexes_for_mass_start, lower_case_text, max_l_dist=2)
            start_index = mass_start[0].end
            mass = read_consecutive_numbers_from_index(lower_case_text, start_index)

            if license_plate is not None:
                license_plate = license_plate.upper()
                license_plate = license_plate.split(',')

            license_plate = license_plate[0].strip()
            transfer_number = transfer_number.strip()

            # Error handling section
            errors = []
            if len(transfer_number) != 8:
                errors.append("Nem megfelelő szállítólevélszám")
            if len(license_plate) != 7:
                errors.append("Nem megfelelő rendszám")

            # Append the extracted data and errors to the DataFrame
            excel_data = excel_data.append({
                "Dátum": date.strip() if date else "",
                "Rendszám": license_plate,
                "Szállítólevél száma": transfer_number,
                "Súly": mass.strip() if mass else "",
                "Hiba": ", ".join(errors) if errors else ""
            }, ignore_index=True)

    return excel_data


def browse_folder(folder_var):
    folder_selected = filedialog.askdirectory()
    folder_var.set(folder_selected)


def process_folder():
    folder_path = source_folder_var.get()
    if folder_path:
        excel_data = pd.DataFrame(columns=["Dátum", "Rendszám", "Szállítólevél száma", "Súly", "Hiba"])
        for file_name in os.listdir(folder_path):
            if file_name.lower().endswith(".pdf"):
                file_path = os.path.join(folder_path, file_name)
                file_data = extract_data_from_pdf(file_path)
                excel_data = pd.concat([excel_data, file_data], ignore_index=True)

        # Save the DataFrame to an Excel file
        destination_folder = destination_folder_var.get()
        excel_file = "output.xlsx"
        destination_folder_path = os.path.join(destination_folder, excel_file)
        excel_data.to_excel(destination_folder_path, index=False)
        result_label.config(text=f"Extraction completed. Data saved to {excel_file}")

# initialization of the OCR engine
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Create the main windowv with Tkinter
root = tk.Tk()
root.title("Szállítólevél olvasó")

# Folder selection
source_folder_var = tk.StringVar()
destination_folder_var = tk.StringVar()

source_folder_label = tk.Label(root, text="Kérem válassza ki a forrásmappát (ahol a PDF-ek vannak):")
source_folder_entry = tk.Entry(root, textvariable=source_folder_var, state="readonly", width=50)
source_folder_button = tk.Button(root, text="Kiválasztás", command=lambda: browse_folder(source_folder_var))

destination_folder_label = tk.Label(root, text="Kérem válassza ki a célmappát(ahova szeretné menteni az excelt):")
destination_folder_entry = tk.Entry(root, textvariable=destination_folder_var, state="readonly", width=50)
destination_folder_button = tk.Button(root, text="Kivlasztás", command=lambda: browse_folder(destination_folder_var))

extract_button = tk.Button(root, text="Futtatás", command=process_folder)

# A felső rész elrendezése
source_folder_label.grid(row=0, column=0)
source_folder_entry.grid(row=0, column=1)
source_folder_button.grid(row=0, column=2)
destination_folder_label.grid(row=1, column=0)
destination_folder_entry.grid(row=1, column=1)
destination_folder_button.grid(row=1, column=2)
extract_button.grid(row=2, column=0)


# # Az alsó rész létrehozása
# error_list = tk.Listbox(root)
#
# # Az alsó rész elrendezése
# error_list.grid(row=3, column=0, rowspan=3)

# Extraction button

# Result label
result_label = tk.Label(root, text="")


# Get the screen size and set the window size to half of it
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = screen_width // 2 - 50
window_height = screen_height // 2 - 50

# Set the window geometry
root.geometry(f"{window_width}x{window_height}+{50}+{50}")

# Start the main loop
root.mainloop()
