import os
import re
import docx

try:
    import docx2txt
except ImportError:
    print("Error: docx2txt library is not installed. Please install it using 'pip install docx2txt'.")
    exit(1)

def extract_sfx_lines(docx_file):
    doc = docx.Document(docx_file)
    sfx_lines = {}
    current_page_number = None
    next_line_is_sfx = False
    for paragraph in doc.paragraphs:
        if "Page" in paragraph.text:
            current_page_number = paragraph.text.strip()
            sfx_lines[current_page_number] = []
        elif "FX" in paragraph.text and current_page_number:
            next_line_is_sfx = True
        elif next_line_is_sfx and current_page_number:
            sfx_lines[current_page_number].append(paragraph.text.strip())
            next_line_is_sfx = False
    return sfx_lines

def remove_unnecessary_page_lines(sfx_lines_by_page):
    for page_number, sfx_lines in sfx_lines_by_page.items():
        new_sfx_lines = []
        for line in sfx_lines:
            if "Page" not in line:
                new_sfx_lines.append(line)
        sfx_lines_by_page[page_number] = new_sfx_lines
    return sfx_lines_by_page

def add_newline_before_page_except_first(output_file):
    with open(output_file, "r", encoding="utf-8") as file:
        lines = file.readlines()

    # Add new line before every "Page" except the first one
    for i in range(1, len(lines)):
        if lines[i].startswith("Page") and not lines[i-1].startswith("Page"):
            lines[i] = "\n" + lines[i]

    # Write the modified content back to the file
    with open(output_file, "w", encoding="utf-8") as file:
        file.writelines(lines)

def main(docx_file):

    # Get the directory and filename of the Word document
    docx_dir, docx_filename = os.path.split(docx_file)

    # Extract SFX lines associated with page numbers
    sfx_lines_by_page = extract_sfx_lines(docx_file)

    # Remove unnecessary Page lines
    sfx_lines_by_page = remove_unnecessary_page_lines(sfx_lines_by_page)

    # Write extracted lines to a text file
    output_filename = os.path.splitext(docx_filename)[0] + "_SFX.txt"
    output_file = os.path.join(docx_dir, output_filename)

    with open(output_file, "w", encoding="utf-8") as file:
        for page_number, sfx_lines in sfx_lines_by_page.items():
            if sfx_lines:  # Only write page number if there are SFX lines
                file.write(page_number + "\n")
                for line in sfx_lines:
                    file.write(line + "\n")
    
    # Add a new line before every page except the first page
    add_newline_before_page_except_first(output_file)
    
    print(f"Extracted lines saved to {output_file}")


if __name__ == "__main__":
    main()