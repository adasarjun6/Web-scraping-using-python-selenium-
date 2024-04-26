import csv
from docx import Document
import re

def extract_text_and_links_from_docx(docx_file_path):
    doc = Document(docx_file_path)
    extracted_data = []

    # Iterate over paragraphs in the document
    for paragraph in doc.paragraphs:
        # Split the paragraph text based on bullet points and whitespace
        sections = re.split(r'\s*\n\s*[-\*\u2022]\s*', paragraph.text.strip())
        
        # Extract URLs from each section and add to extracted data
        for section in sections:
            links = re.findall(r'https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+[-\w]*', section)
            text_content = section

            # Append (text, links) tuple to extracted data
            extracted_data.append((text_content, ", ".join(links)))

    return extracted_data

def save_text_and_links_to_csv(data, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Write header row
        writer.writerow(['Text Content', 'Links'])

        # Write data rows
        for text, links in data:
            writer.writerow([text, links])

# Example usage:
if __name__ == "__main__":
    docx_file_path = "data.docx"
    output_file = 'scraped_data.csv'

    # Extract text content and associated links from the DOCX file
    extracted_data = extract_text_and_links_from_docx(docx_file_path)

    # Save extracted text and links to the CSV file
    if extracted_data:
        save_text_and_links_to_csv(extracted_data, output_file)
        print(f"Extracted text and links saved to '{output_file}'")
    else:
        print("No text and links found in the document.")
