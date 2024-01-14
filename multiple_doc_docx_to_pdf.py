import os
import win32com.client

try:
    import pywin32
except ImportError:
    print("pywin32 is not installed. Please wait until the installation is complete...")
    os.system("pip install pywin32")

def prompt_for_folder(text_prompt):
    while True:
        directory = input(text_prompt).rstrip("\\")  
        if os.path.isdir(directory):
            return directory
        else:
            print("Invalid folder path. Please enter a valid folder address following the prompt's template text.")


prompt_input_folder = "Enter the source folder path (e.g., C:\\path\\to\\folder): "
prompt_output_folder = "Enter the destination folder path (e.g., C:\\path\\to\\folder): "

input_folder = prompt_for_folder(prompt_input_folder)
output_folder = prompt_for_folder(prompt_output_folder)

word_doc = win32com.client.Dispatch("Word.Application")

def convert_to_pdf(file_directory, output_folder):
    document = word_doc.Documents.Open(file_directory)

    pdf_file = os.path.splitext(os.path.basename(file_directory))[0] + ".pdf"
    pdf_file = os.path.join(output_folder, pdf_file)

    document.SaveAs(pdf_file, FileFormat=17)

    document.Close()
    print(f"Successfully converted from {file_directory} to {pdf_file}")

for root, dirs, files in os.walk(input_folder):
    for document_name in files:
        if not (document_name.lower().endswith((".doc", ".docx")) and not document_name.startswith("~$")):
            continue
        file_directory = os.path.join(root, document_name)
        relative_path = os.path.relpath(file_directory, input_folder)
        output_subfolder = os.path.join(output_folder, os.path.dirname(relative_path))
        os.makedirs(output_subfolder, exist_ok=True)

        convert_to_pdf(file_directory, output_subfolder)

word_doc.Quit()
