import os
import io
import docx
from pdfminer.high_level import extract_text_to_fp


def pdf_to_word(pdf_file_name):
    # Get the full path of the PDF file
    pdf_file_path = os.path.join(os.getcwd(), pdf_file_name)

    # Open PDF file in read-binary mode
    with open(pdf_file_path, 'rb') as pdf_file:
        # Use io.BytesIO() to get the bytes stream of the PDF content
        pdf_stream = io.BytesIO(pdf_file.read())

    # Use pdfminer to extract text from the PDF
    text_stream = io.StringIO()
    extract_text_to_fp(pdf_stream, text_stream)

    # Create a new Word document
    doc = docx.Document()

    # Add each line of text from the PDF to the Word document
    for line in text_stream.getvalue().split('\n'):
        doc.add_paragraph(line)

    # Get the full path of the Word file
    word_file_name = os.path.splitext(pdf_file_name)[0] + '.docx'
    word_file_path = os.path.join(os.getcwd(), word_file_name)

    # Save the Word document
    doc.save(word_file_path)

    print(f'Successfully converted {pdf_file_name} to {word_file_name}.')


# Get the PDF file name from user input
pdf_file_name = input('Enter the name of the PDF file to convert: ')

# Convert the PDF file to Word
pdf_to_word(pdf_file_name)
