from tkinter import filedialog
from pypdf import PdfReader, PdfWriter
import os
import glob

# Step 1: Select a PDF file to split into pages of 3
input_file_path = filedialog.askopenfilename(title="Select Pdf File to split into pages of 3")

# Step 2: Create a PdfReader instance to read the PDF file
with open(input_file_path, 'rb') as file:
    pdf_reader = PdfReader(file)
    num_pages = len(pdf_reader.pages)

# Step 3: Get the current working directory
    current_directory = os.getcwd()

# Step 4: Check if a folder named "PDF" exists in the current working directory
    pdf_folder = os.path.join(current_directory, 'PDF')

    if not os.path.exists(pdf_folder):
        # If the folder doesn't exist, create it
        os.makedirs(pdf_folder)

# Step 5: Iterate over pages and split into groups of 3
    for start_page in range(0, num_pages, 3):
        end_page = min(start_page + 3, num_pages)
        pdf_writer = PdfWriter()
        output_pdf_path = os.path.join(pdf_folder, f'split_pages_{start_page + 1}-{end_page}.pdf')
        with open(output_pdf_path, 'wb') as output_pdf:
            for page_num in range(start_page, end_page):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)

            pdf_writer.write(output_pdf)

# After splitting, navigate to the 'PDF' folder
pdf_directory = os.path.join(os.getcwd(), 'PDF')
os.chdir(pdf_directory)

# Rename PDFs based on extracted information
pdf_files = glob.glob('*.pdf')

for pdf in pdf_files:
    reader = PdfReader(pdf)
    first_page = reader.pages[0]
    text_content = first_page.extract_text()
    lines = text_content.split('\n')

    if len(lines) >= 6:
        first_name = lines[4]
        last_name = lines[5]
        new_name = (last_name + first_name).lower()

        new_file_path = os.path.join(pdf_directory, new_name + '.pdf')

        if not os.path.exists(new_file_path):
            os.rename(pdf, new_name + '.pdf')
            print(f"{new_name} has been saved! to {new_file_path}")
        else:
            print(f'{new_name}.pdf already exists')
