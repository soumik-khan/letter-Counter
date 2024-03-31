import re
from collections import Counter
from PyPDF2 import PdfFileReader
from docx import Document

def count_letters(file_path):
    if file_path.endswith('.txt'):
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
    elif file_path.endswith('.pdf'):
        with open(file_path, 'rb') as file:
            pdf_reader = PdfFileReader(file)
            text = ''
            for page_num in range(pdf_reader.numPages):
                text += pdf_reader.getPage(page_num).extractText()
    elif file_path.endswith('.docx'):
        doc = Document(file_path)
        text = ''
        for para in doc.paragraphs:
            text += para.text
    
    # Remove non-alphabetic characters and convert text to lowercase
    text = re.sub(r'[^a-zA-Z]', '', text.lower())
    
    # Count the occurrences of each letter
    letter_counts = Counter(text)
    
    # Calculate the total number of letters
    total_letters = sum(letter_counts.values())
    
    # Calculate the percentage of each letter
    letter_percentages = {letter: (count / total_letters) * 100 for letter, count in letter_counts.items()}
    
    return letter_counts, letter_percentages

def save_results_to_variables(file_path, save_format):
    letter_counts, letter_percentages = count_letters(file_path)
    if save_format == 'variables':
        return letter_counts, letter_percentages
    elif save_format == 'files':
        # You can add file-saving functionality here if needed
        print("Results saved to files successfully.")

def main():
    file_path = "sample.txt"  # Predefined file path for testing
    save_format = "variables"  # Predefined save format for testing
    if os.path.exists(file_path):
        letter_counts, letter_percentages = save_results_to_variables(file_path, save_format)
        print("Results saved to variables successfully.")
        return letter_counts, letter_percentages
    else:
        print("File not found.")
        return None, None

if __name__ == "__main__":
    main()
