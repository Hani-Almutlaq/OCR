import fitz  # From PyMuPDF library, to handling PDF files
from PIL import Image  # For handling image processing
import pytesseract  # Tesseract OCR library for text extraction
import pandas as pd  # Library for handling data in DataFrame format
import os  # Library for file and directory handling
import io  # Library for byte handling
import re  # Library for regular expressions

# Set the path to the OCR engine (tesseract.exe)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Initialize a DataFrame to store extracted data
# Columns: Full Name, Student ID, GPA
students = pd.DataFrame(columns=["Full Name", "Student ID", "GPA"])

# Specify the folder containing the PDF files
folder = "C:\\Users\\Hani\\Desktop\\OCR\\Samples" # The folder that contain Input
directory_list = os.listdir(folder)  # List all files in the specified folder

# Iterate over all files in the folder
for file_name in directory_list:
    if file_name.endswith(".pdf"): # Process only PDF files
        pdf_document = fitz.open(os.path.join(folder, file_name)) # Open the PDF file using PyMuPDF
        page = pdf_document[0] # Extract the first page of the PDF (assuming there's only one page)

        # Extract the image from the page
        image_reference = page.get_images(full=True)[0][0]  # Reference to the single image object
        base_image = pdf_document.extract_image(image_reference)  # Extract image data
        image_bytes = base_image["image"]  # Get image bytes

        # Convert image bytes to a PIL Image object
        image = Image.open(io.BytesIO(image_bytes))

        # Perform OCR (Optical Character Recognition) to extract text
        text = pytesseract.image_to_string(image)
        print(text)  # Print the extracted text for debugging purposes


        # Extract the candidate's name
        # Matches "Candidate name" followed by the actual name
        name_match = re.search(r'Candidate name\s*\n\s*(\S.+)', text, re.IGNORECASE)

        # Extract the student ID
        # Matches "Processing No." followed by a 9-digit number
        id_match = re.search(r'Processing No.\s*\n\s*(\d{9})', text, re.IGNORECASE)

        # Extract the GPA
        # Matches "Cum GPA" followed by optional underscores/spaces and the GPA value
        gpa_match = re.search(r'Cum GPA[_\s]*([\d\.]+)', text, re.IGNORECASE)

        # Retrieve the extracted data, or use "Not Found" if the match is unsuccessful
        name = name_match.group(1).strip() if name_match else "Not Found"
        student_id = id_match.group(1).strip() if id_match else "Not Found"
        gpa = gpa_match.group(1).strip() if gpa_match else "Not Found"

        # Append the extracted data to the DataFrame
        students = students._append({"Full Name": name, "Student ID": student_id, "GPA": gpa}, ignore_index=True)

output_file = "C:\\Users\\Hani\\Desktop\\OCR\\Students Reports.xlsx" # Specify the output file path and name
students.to_excel(output_file, index=False) # Export the DataFrame to an Excel file

print(f"Data has been successfully exported to {output_file}") # Print a success message