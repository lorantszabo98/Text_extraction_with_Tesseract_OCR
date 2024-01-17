# Text extraction with Tesseract OCR  

This program is a simple and user-friendly tool designed to quickly and efficiently extract information from Pdfs. The program processes PDF documents, converts them into images, and then performs optical character recognition (OCR) on their content.

## Installation

1. Install Tesseract OCR:
- Ensure that Tesseract OCR is installed on your computer. If it is not, you can download the latest version from [here](https://github.com/UB-Mannheim/tesseract/wiki) and follow the installation guide.

2. Add the path of the Tesseract executable to this line of the code:
```python
# initialization of the OCR engine
pytesseract.pytesseract.tesseract_cmd = r''
```
3. Intall the pytesseract library with the following command:
```bash
pip install pytesseract
```
4. Install the other requirements:
```bash
pip install fuzzysearch PyMuPDF pillow pandas
```
5. Change the language of the OCR if you need, in this line:
```python
text = pytesseract.image_to_string(image, lang="hun")
```
6. Define you own regexes and column names for the text extraction, for example:
```python
# Define your own column names
excel_data = pd.DataFrame(columns=["", "", "", "", ""])

# Define your own regexes 
regexes_for_transfer_number_start = ['']
regexes_for_transfer_number_end = ['']
```

## Execution

1. Run the program.

2. Select Folders:
- Click the first "Kiválasztás" button and choose the folder containing the PDFs.
- Click the second "Kiválasztás" button and choose the folder where the program will save the Excel file.
  
3.  Run the text extraction:
Click the "Futtatás" button to start the script.
The program processes the selected PDF files and displays the progress during execution.

4. Results:
The Excel file containing the processed data and errors (output.xlsx) will be saved in the selected destination folder.


## Notes

### Multi-page PDFs:

If the PDF contains multiple pages, the program processes each page and stores the results in a single Excel file.
In case of errors, the program indicates the type of error and appends the file name and page number for easier troubleshooting.

### Information:

Click on the "ℹ" icon to view information about the program, including its steps and requirements. (In hungarian)
