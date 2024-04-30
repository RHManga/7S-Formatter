import os
import re
import docx

try:
    import docx2txt
except ImportError:
    print("Error: docx2txt library is not installed. Please install it using 'pip install docx2txt'.")
    exit(1)

def convert_to_tags(text, bold=False, italic=False, bold_italic=False):
    tags = ""
    if bold or italic or bold_italic:
        if bold_italic:
            tags += "[k]"
        elif bold:
            tags += "[b]"
        elif italic:
            tags += "[i]"
    return tags + text

def close_tags(text, bold=False, italic=False, bold_italic=False):
    closing_tags = ""
    if bold_italic:
        closing_tags += "[/k]"
    elif bold:
        closing_tags += "[/b]"
    elif italic:
        closing_tags += "[/i]"

    # Remove leading space before closing tags if present
    if closing_tags.startswith(' '):
        closing_tags = closing_tags.lstrip()

    return closing_tags

def extract_text_with_format(docx_file):
    doc = docx.Document(docx_file)
    extracted_text = ""
    prev_bold = False
    prev_italic = False
    prev_bold_italic = False
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            text = run.text
            bold = run.bold
            italic = run.italic
            if 'b' in run.style.name:
                bold = True
            if 'i' in run.style.name:
                italic = True
            if bold == prev_bold and italic == prev_italic and (bold and italic) == prev_bold_italic:
                extracted_text += text
            else:
                extracted_text += close_tags(paragraph.text, prev_bold, prev_italic, prev_bold and prev_italic)
                extracted_text += convert_to_tags(text, bold, italic, bold and italic)
            prev_bold = bold
            prev_italic = italic
            prev_bold_italic = bold and italic
        extracted_text += close_tags(paragraph.text, bold, italic, bold and italic) + "\n"
        prev_bold = False
        prev_italic = False
        prev_bold_italic = False
    
    # Remove text matching the specified regex
    extracted_text = re.sub(r'(FX.*\n)(.*?\n)?', r'', extracted_text)
    
    return extracted_text

def remove_b_tags_from_page_lines(text):
    lines = text.split("\n")
    cleaned_lines = []
    for line in lines:
        if line.strip().startswith("[b]Page"):
            # Remove [b] tags from "Page" lines
            line = re.sub(r'\[/?b\]', '', line)
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines)

def remove_lines_matching_patterns(text, patterns):
    lines = text.split("\n")
    cleaned_lines = []
    for line in lines:
        if not any(re.match(pattern, line.strip()) for pattern in patterns):
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)

def replace_heart_symbol(text):
    return text.replace(r"\[heart\]", "♡")

def remove_tagged_lines(text):
    patterns = [r'^\[+[bi]+\]', r'^\[\/[bik]+\]']  # Add the patterns used in remove_lines
    lines = text.split("\n")
    cleaned_lines = []
    for line in lines:
        if any(re.match(pattern, line.strip()) for pattern in patterns):
            cleaned_lines.append('')  # Remove lines with matched tags
        else:
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def remove_lines(text):
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        # Check if the line is empty or a "Page" line
        if line.strip() == '' or line.strip().startswith("[b]Page"):
            cleaned_lines.append(line)
        # For other lines, apply regex deletion
        elif not re.match(r'^\[+/+[bik]+\]+$', line.strip()) and not re.match(r'^[0-9]+\.[0-9]+|^[0-9]+', line.strip()) and not re.match(r'^[0-9]+\s+[a-zA-Z]+$', line.strip()):
            cleaned_lines.append(line)
    cleaned_text = '\n'.join(cleaned_lines)
    return cleaned_text

def add_newline_before_page(text):
    return re.sub(r'(^Page)', r'\n\1', text, flags=re.MULTILINE)

def main(docx_file):
    
    # Get the directory and filename of the Word document
    docx_dir, docx_filename = os.path.split(docx_file)
    
    # Regular expression patterns to match lines with the specified formats
    patterns = [r'^<[bB]>\d+(\.\d+)?', r'^[0-9]+(\.\d+)?']

    # Extract text with formatting
    formatted_text = extract_text_with_format(docx_file)
    
    # Write the formatted text to a temporary .txt file
    temp_output_file = os.path.join(docx_dir, "temp_formatted_text.txt")
    with open(temp_output_file, "w", encoding="utf-8") as file:
        file.write(formatted_text)
    
    # Read the generated text from the temporary .txt file
    with open(temp_output_file, "r", encoding="utf-8") as file:
        cleaned_text = file.read()
    
    # Remove [b] tags from "Page" lines
    cleaned_text = remove_b_tags_from_page_lines(cleaned_text)
    
    # Remove lines matching the specified regex patterns
    cleaned_text = remove_tagged_lines(cleaned_text)
    
    # Remove blank lines without tags and lines matching specified regex expressions
    cleaned_text = remove_lines(cleaned_text)
    
    # Construct the output file path with the formatted filename
    output_filename = os.path.splitext(docx_filename)[0] + "_formatted.txt"
    output_file = os.path.join(docx_dir, output_filename)
    
    # Write the cleaned text to the output .txt file
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(cleaned_text)
    
    with open(output_file, "w", encoding="utf-8") as file:
        # Remove spaces before closing tags, after opening tags, and add spaces between consecutive tags
        found_non_empty_line = False
        cleaned_lines = cleaned_text.split('\n')
        cleaned_lines = [line for line in cleaned_lines if (found_non_empty_line or line.strip('\r') != '') and (found_non_empty_line or line.strip() != '')]
        cleaned_text = '\n'.join(cleaned_lines)
        cleaned_text = add_newline_before_page(cleaned_text)
        cleaned_text = re.sub(r'([^\s\n])\s+(\[/[bik]\])', r'\1\2', cleaned_text)
        cleaned_text = re.sub(r'(\[[bik]\])+([^\s\n])', r'\1\2', cleaned_text)
        cleaned_text = re.sub(r'(\]\[)', '] [', cleaned_text)
        cleaned_text = re.sub(r'(\[/[bik]+\])+(\S)', r'\1 \2', cleaned_text)
        cleaned_text = re.sub(r'(\S)(\[[bik]\])', r'\1 \2', cleaned_text)
        cleaned_text = re.sub(r'\[[bik]\] \[\/[bik]\]', r'', cleaned_text)
        cleaned_text = re.sub(r'  ', r' ', cleaned_text)
        file.write(cleaned_text)

    # Remove the temporary .txt file
    os.remove(temp_output_file)

    print(f"Formatted text saved to {output_file}")

if __name__ == "__main__":
    main()
