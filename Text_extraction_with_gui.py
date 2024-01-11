import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from fuzzysearch import find_near_matches
import fitz
from PIL import Image
import pytesseract
import pandas as pd
from queue import Queue
import threading


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


def regex_not_found_message(column):
    return f"A {column} nem meghatározható"


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

            errors = []

            regexes_for_date_start = ['teljesítés kelte:']
            regexes_for_date_end = ['szállítólevélszám']
            date = extract_data(regexes_for_date_start, regexes_for_date_end, lower_case_text, max_l_dist=2)

            if date is None:
                errors.append(regex_not_found_message("Dátum"))

            regexes_for_transfer_number_start = ['szállítólevélszám:']
            regexes_for_transfer_number_end = ['jármű']
            transfer_number = extract_data(regexes_for_transfer_number_start, regexes_for_transfer_number_end, lower_case_text, max_l_dist=1)

            if transfer_number is None:
                errors.append(regex_not_found_message("Szállítólevélszám"))

            else:
                transfer_number = transfer_number.strip()
                if len(transfer_number) != 8:
                    errors.append("Nem megfelelő szállítólevélszám")

            regexes_for_license_plate_start = ['jármű:']
            regexes_for_license_plate_end = ['mérlegelt bruttó']
            license_plate = extract_data(regexes_for_license_plate_start, regexes_for_license_plate_end, lower_case_text, max_l_dist=1)

            if license_plate is None:
                errors.append(regex_not_found_message("Rendszám"))
            else:
                license_plate = license_plate.upper()
                license_plate = license_plate.split(',')

                license_plate = license_plate[0].strip()
                if len(license_plate) != 7:
                    errors.append("Nem megfelelő rendszám")

            regexes_for_mass_start = ['nettó tömeg:']
            mass_start = find_regex_with_fuzzy(regexes_for_mass_start, lower_case_text, max_l_dist=2)

            if mass_start is None:
                errors.append(regex_not_found_message("Súly"))
                mass = None
            else:
                start_index = mass_start[0].end
                mass = read_consecutive_numbers_from_index(lower_case_text, start_index)

            # if there is an error, we add the page number to the file for easier inspection
            if len(errors) > 0:
                errors.insert(0, f"Az oldal száma: {page_num + 1}")

            # Append the extracted data and errors to the DataFrame
            excel_data = excel_data.append({
                "Dátum": date.strip() if date else "",
                "Rendszám": license_plate if license_plate else "",
                "Szállítólevél száma": transfer_number if transfer_number else "",
                "Súly": mass.strip() if mass else "",
                "Hiba": ", ".join(errors) if errors else ""
            }, ignore_index=True)

    return excel_data


def browse_folder(folder_var):
    folder_selected = filedialog.askdirectory()
    folder_var.set(folder_selected)


def process_folder(source_folder, destination_folder, progress_queue):

    excel_data = pd.DataFrame(columns=["Dátum", "Rendszám", "Szállítólevél száma", "Súly", "Hiba"])
    file_list = [file_name for file_name in os.listdir(source_folder) if file_name.lower().endswith(".pdf")]
    total_files = len(file_list)

    for i, file_name in enumerate(file_list):
        file_path = os.path.join(source_folder, file_name)
        file_data = extract_data_from_pdf(file_path)
        excel_data = pd.concat([excel_data, file_data], ignore_index=True)

        progress_queue.put((i + 1) * 100 // total_files)

    # Save the DataFrame to an Excel file
    excel_file = "output.xlsx"
    destination_folder_path = os.path.join(destination_folder, excel_file)
    try:
        excel_data.to_excel(destination_folder_path, index=False)

        tk.messagebox.showinfo("Kész", f"Az excel mentve lett a {destination_folder_path} helyre")

    except PermissionError:
        tk.messagebox.showerror('Hiba', 'Az Ecel fájl meg van nyitva! \nZárja be az Excelt és futtassa újra a programot!')
    except Exception as e:
        tk.messagebox.showerror("Hiba", f"Hiba történt a fájl mentése közben: {e}")

    progress_queue.put(100)


def update_progress_bar(progress_var, progress_queue, root):
    progress_text.grid(row=4, column=0, pady=5)
    progress_bar.grid(row=5, column=0, pady=5)

    def update():
        while not progress_queue.empty():
            progress_value = progress_queue.get()
            progress_var.set(progress_value)
            root.update_idletasks()

            if progress_value == 100:
                progress_text.grid_forget()
                progress_bar.grid_forget()
        root.after(1, update)  # Schedule the update after a short delay

    # Start the update loop
    update()


def start_processing():
    source_folder = source_folder_var.get()
    destination_folder = destination_folder_var.get()

    if not source_folder or not destination_folder:
        tk.messagebox.showerror("Hiba", "Kérem válasszon forrás és célmappát!")
        return

    # Start a new thread for processing
    progress_queue = Queue()
    process_thread = threading.Thread(target=process_folder, args=(source_folder, destination_folder, progress_queue), daemon=True)
    process_thread.start()

    # Start a new thread for updating progress
    update_thread = threading.Thread(target=update_progress_bar, args=(progress_var, progress_queue, root), daemon=True)
    update_thread.start()


def show_info():
    info_text = "- A forrásmappa kiválasztásához kattintson az első 'Kiválasztás' gombra, majd itt keresse meg azt a mappát " \
                ",amely a PDF-eket tartalmazza, kattintson a mappára egyszer, majd lent a 'Mappaválasztás gombra'.\n\n" \
                "- Ugyanezt végezze el a célmappa esetén is, de itt azt a mappát válassza ki, ahova menteni szeretné " \
                "az Excel fájlt.\n\n" \
                "- Ha mindez megvan, akkor kattinthat a 'Futtatás' gombra ezzel végrehajtva a szöveg kinyerését a PDF-ekből.\n\n" \
                "- Ez több percig is eltarthat a PDF-ek számának fügvényében, kérem addig ne zárja be a programot.\n\n" \
                "- A folyamat befejezését a program jelzi, majd a létrehozott Excel fájlt megtalálja a célmappába mentve 'output.xlsx' " \
                "néven.\n\n"
    info_window = tk.Toplevel(root)
    info_window.title("Információ")

    info_label = tk.Label(info_window, text=info_text, padx=10, pady=10, wraplength=600, justify="left")
    info_label.pack()

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

destination_folder_label = tk.Label(root, text="Kérem válassza ki a célmappát (ahova szeretné menteni az excelt):")
destination_folder_entry = tk.Entry(root, textvariable=destination_folder_var, state="readonly", width=50)
destination_folder_button = tk.Button(root, text="Kiválasztás", command=lambda: browse_folder(destination_folder_var))

extract_button = tk.Button(root, text="Futtatás", command=start_processing)

# A felső rész elrendezése
source_folder_label.grid(row=1, column=0, sticky="w")
source_folder_entry.grid(row=1, column=1)
source_folder_button.grid(row=1, column=2)
destination_folder_label.grid(row=2, column=0, sticky="w")
destination_folder_entry.grid(row=2, column=1)
destination_folder_button.grid(row=2, column=2)
extract_button.grid(row=3, column=0)

progress_text = tk.Label(root, text="Futtatás..., ez több percig is eltarthat")

# Progress bar
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, mode='determinate')


# Info Icon
info_icon = ttk.Label(root, text="ℹ", font=("Arial", 14), cursor="hand2")
info_icon.grid(row=0, column=2, padx=5)
info_icon.bind("<Button-1>", lambda event: show_info())

# Get the screen size and set the window size to half of it
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = screen_width // 2 - 40
window_height = screen_height // 2 - 40

# Set the window geometry
root.geometry(f"{window_width}x{window_height}+{50}+{50}")

# Start the main loop
root.mainloop()
