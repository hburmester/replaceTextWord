import os
import sys
from docx import Document

# Define constants
SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))
OLD_TEXT = 'Company Name'

def handle_file_errors(func):
    def wrapper(*args, **kwargs):
        try:
            # Call the wrapped function
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            # Handle file not found error
            print(f"Error: File not found in '{SCRIPT_FOLDER}'.")
            sys.exit(1)
    return wrapper

@handle_file_errors
def replace_text_in_docx(new_text):
    file_path = os.path.join(SCRIPT_FOLDER, 'Cover letter.docx')

    # Open the document
    doc = Document(file_path)

    # Replace text in the document
    for paragraph in doc.paragraphs:
        if OLD_TEXT in paragraph.text:
            paragraph.text = paragraph.text.replace(OLD_TEXT, new_text)

    # Save the modified document
    new_file_path = os.path.join(SCRIPT_FOLDER+'/{new_text}', f'Cover letter.docx')
    doc.save(new_file_path)

    print(f"New document saved as '{new_text}.docx'.")

if __name__ == "__main__":
    # Check if correct number of arguments are provided
    if len(sys.argv) != 2:
        print("Usage: python replace_text.py <new_text>")
        sys.exit(1)
    
    # Extract command-line argument for the new text
    new_text = sys.argv[1]

    # Perform text replacement and save the new document
    replace_text_in_docx(new_text)
