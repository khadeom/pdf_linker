from docx import Document

def extract_word_text_and_positions(word_path):
    doc = Document(word_path)
    word_data = []
    for para_idx, paragraph in enumerate(doc.paragraphs):
        for run in paragraph.runs:
            word_data.append({
                "text": run.text,
                "paragraph": para_idx,
                "offset": paragraph.text.find(run.text)
            })
    return word_data

import fitz  # PyMuPDF

def extract_pdf_text_and_positions(pdf_path):
    pdf_doc = fitz.open(pdf_path)
    pdf_data = []
    for page_num in range(len(pdf_doc)):
        page = pdf_doc.load_page(page_num)
        words = page.get_text("words")  # Extract words with positions
        for word in words:
            pdf_data.append({
                "text": word[4],  # The actual word text
                "page": page_num,
                "x": word[0],  # x-coordinate
                "y": word[1]   # y-coordinate
            })
    return pdf_data

def match_words(word_data, pdf_data):
    matches = []
    for word in word_data:
        for pdf_word in pdf_data:
            if word["text"] == pdf_word["text"]:  # Match by text
                matches.append({
                    "word_text": word["text"],
                    "word_location": {
                        "paragraph": word["paragraph"],
                        "offset": word["offset"]
                    },
                    "pdf_location": {
                        "page": pdf_word["page"],
                        "x": pdf_word["x"],
                        "y": pdf_word["y"]
                    }
                })
                break  # Stop after the first match
    return matches

from docx import Document
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink_to_word(word_path, matches, output_path):
    doc = Document(word_path)
    for match in matches:
        para_idx = match["word_location"]["paragraph"]
        offset = match["word_location"]["offset"]
        text = match["word_text"]

        # Find the paragraph and run containing the word
        paragraph = doc.paragraphs[para_idx]
        for run in paragraph.runs:
            if text in run.text:
                # Add a hyperlink to the word
                hyperlink = OxmlElement("w:hyperlink")
                hyperlink.set(qn("r:id"), f"mismatch_{match['pdf_location']['page']}")
                run._element.append(hyperlink)
                break

    doc.save(output_path)

    from docx import Document
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink_to_word(word_path, matches, output_path):
    doc = Document(word_path)
    for match in matches:
        para_idx = match["word_location"]["paragraph"]
        offset = match["word_location"]["offset"]
        text = match["word_text"]

        # Find the paragraph and run containing the word
        paragraph = doc.paragraphs[para_idx]
        for run in paragraph.runs:
            if text in run.text:
                # Add a hyperlink to the word
                hyperlink = OxmlElement("w:hyperlink")
                hyperlink.set(qn("r:id"), f"mismatch_{match['pdf_location']['page']}")
                run._element.append(hyperlink)
                break

    doc.save(output_path)



word_path = "example.docx"
pdf_path = "example.pdf"
output_path = "linked_example.docx"

def link_word_to_pdf(word_path, pdf_path, output_path):
    # Step 1: Extract data from Word
    word_data = extract_word_text_and_positions(word_path)

    # Step 2: Extract data from PDF
    pdf_data = extract_pdf_text_and_positions(pdf_path)

    # Step 3: Match words
    matches = match_words(word_data, pdf_data)

    # Step 4: Add hyperlinks to Word
    add_hyperlink_to_word(word_path, matches, output_path)

    print(f"Linked Word document saved to {output_path}")
    

link_word_to_pdf(word_path, pdf_path, output_path)