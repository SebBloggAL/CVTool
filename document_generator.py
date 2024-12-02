# document_generator.py

import os
import logging
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import docx
import docx.oxml
import re
from datetime import datetime

def set_document_font(doc):
    """
    Sets the default font for the entire document to Calibri with a size of 10.5 pt.
    """
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(10.5)
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

def set_heading_style(doc):
    """
    Creates or modifies a style named 'Style 1' for headings with size 16 pt, amber color, bold, and underlined.
    """
    styles = doc.styles
    try:
        heading_style = styles['Style 1']
    except KeyError:
        heading_style = styles.add_style('Style 1', WD_STYLE_TYPE.PARAGRAPH)

    font = heading_style.font
    font.name = 'Calibri'
    font.size = Pt(16)
    font.color.rgb = RGBColor(226, 106, 35)  # Amber color
    font.bold = True
    font.underline = True  # Underline the headings
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

    # Ensure paragraph formatting is consistent
    paragraph_format = heading_style.paragraph_format
    paragraph_format.space_after = Pt(12)

def set_document_defaults_language(doc):
    """
    Sets the document-wide default language to British English.
    """
    styles_element = doc.styles.element
    rpr_default = styles_element.xpath('./w:docDefaults/w:rPrDefault/w:rPr')[0]
    lang_default = rpr_default.xpath('w:lang')
    if lang_default:
        lang = lang_default[0]
    else:
        lang = docx.oxml.shared.OxmlElement('w:lang')
        rpr_default.append(lang)
    lang.set(qn('w:val'), 'en-GB')
    lang.set(qn('w:eastAsia'), 'en-US')
    lang.set(qn('w:bidi'), 'ar-SA')

def set_styles_language(doc):
    """
    Sets the language for all styles to British English.
    """
    for style in doc.styles:
        if style.type in [WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.CHARACTER]:
            rpr = style.element.get_or_add_rPr()
            lang = rpr.find(qn('w:lang'))
            if lang is None:
                lang = docx.oxml.shared.OxmlElement('w:lang')
                rpr.append(lang)
            lang.set(qn('w:val'), 'en-GB')
            lang.set(qn('w:eastAsia'), 'en-US')
            lang.set(qn('w:bidi'), 'ar-SA')

def apply_run_font_style(run, paragraph, is_applicant_name=False):
    """
    Applies font and language settings to a run.
    """
    font = run.font
    font.name = 'Calibri'
    if paragraph.style.name == 'Style 1':
        font.size = Pt(16)
        font.underline = True  # Ensure headings are underlined
    elif is_applicant_name:
        font.size = Pt(20)
    else:
        font.size = Pt(10.5)
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

    # Set the language for the run using w:lang element
    rpr = run._element.get_or_add_rPr()
    lang = rpr.find(qn('w:lang'))
    if lang is None:
        lang = docx.oxml.shared.OxmlElement('w:lang')
        rpr.append(lang)
    lang.set(qn('w:val'), 'en-GB')
    lang.set(qn('w:eastAsia'), 'en-US')
    lang.set(qn('w:bidi'), 'ar-SA')

def replace_placeholders_in_paragraph(paragraph, placeholders):
    """
    Replaces placeholders in a paragraph with actual data.
    """
    # Combine all run texts to handle placeholders split across runs
    full_text = ''.join(run.text for run in paragraph.runs)
    original_text = full_text  # Keep original for comparison

    for key, value in placeholders.items():
        if key in full_text:
            full_text = full_text.replace(key, value)
            logging.debug(f"Replaced '{key}' with '{value}' in paragraph.")

    if full_text != original_text:
        # Clear existing runs properly
        for run in paragraph.runs:
            p = run._element
            p.getparent().remove(p)
        paragraph._p.clear_content()
        paragraph.runs.clear()

        # Handle special formatting for specific placeholders
        if "{ApplicantName}" in original_text:
            # Split the text to isolate {ApplicantName}
            parts = original_text.split("{ApplicantName}")
            # Create runs for each part
            for i, part in enumerate(parts):
                if part:
                    run = paragraph.add_run(part)
                    apply_run_font_style(run, paragraph)
                if i < len(parts) - 1:
                    # Insert ApplicantName with font size 20
                    run = paragraph.add_run(placeholders["{ApplicantName}"])
                    apply_run_font_style(run, paragraph, is_applicant_name=True)
        else:
            # Add a new run with the replaced text
            new_run = paragraph.add_run(full_text)
            apply_run_font_style(new_run, paragraph)

def replace_headers(doc):
    """
    Replaces header placeholders (e.g., [Summary]) with actual header text and applies 'Style 1'.
    """
    header_placeholders = {
        "[Security Clearance]": "Security Clearance:",
        "[Summary]": "Summary",
        "[Skills]": "Skills",
        "[Experience]": "Experience",
        "[Education]": "Education"
    }

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text in header_placeholders:
            # Replace the placeholder with actual header text
            paragraph.text = header_placeholders[text]

            # Clear existing runs
            paragraph.clear()

            # Apply 'Style 1' to the paragraph
            paragraph.style = 'Style 1'

            # Add new run with header text
            new_run = paragraph.add_run(header_placeholders[text])

            # Set the font and language for the new run
            font = new_run.font
            font.name = 'Calibri'
            font.size = Pt(16)
            font.bold = True
            font.underline = True  # Underline the headings
            font.color.rgb = RGBColor(226, 106, 35)  # Amber color
            font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

            # Set the language
            rpr = new_run._element.get_or_add_rPr()
            lang = rpr.find(qn('w:lang'))
            if lang is None:
                lang = docx.oxml.shared.OxmlElement('w:lang')
                rpr.append(lang)
            lang.set(qn('w:val'), 'en-GB')
            lang.set(qn('w:eastAsia'), 'en-US')
            lang.set(qn('w:bidi'), 'ar-SA')

def replace_placeholders_in_cell(cell, placeholders, data):
    """
    Replaces placeholders in a single cell and sets font and language for all runs.
    """
    for paragraph in cell.paragraphs:
        if '{Skills}' in paragraph.text:
            # Insert the skills section here
            insert_skills_section(paragraph, data.get('Skills', []))
            # Remove the placeholder paragraph
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
        elif '{Experience}' in paragraph.text:
            # Insert the experience section here
            insert_experience_section(paragraph, data.get('Experience', []))
            # Remove the placeholder paragraph
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
        else:
            replace_placeholders_in_paragraph(paragraph, placeholders)
            convert_lines_to_bullets(paragraph)

    for nested_table in cell.tables:
        replace_placeholders_in_table(nested_table, placeholders, data)

def replace_placeholders_in_table(table, placeholders, data):
    """
    Replaces placeholders in all cells of a table.
    """
    for row in table.rows:
        for cell in row.cells:
            replace_placeholders_in_cell(cell, placeholders, data)

def set_font_for_all_text(doc, placeholders):
    """
    Sets the font and language for all text in the document to Calibri and British English.
    """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            is_applicant_name = run.text == placeholders.get("{ApplicantName}", "")
            apply_run_font_style(run, paragraph, is_applicant_name=is_applicant_name)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        is_applicant_name = run.text == placeholders.get("{ApplicantName}", "")
                        apply_run_font_style(run, paragraph, is_applicant_name=is_applicant_name)

def set_list_bullet_style(doc):
    """
    Defines the 'List Bullet' style if it doesn't exist.
    """
    styles = doc.styles
    if 'List Bullet' not in styles:
        list_bullet_style = styles.add_style('List Bullet', WD_STYLE_TYPE.PARAGRAPH)
        list_bullet_style.base_style = styles['List Paragraph']
        list_bullet_style.paragraph_format.left_indent = Pt(12)
        list_bullet_style.paragraph_format.first_line_indent = Pt(-12)
        list_bullet_style.paragraph_format.space_after = Pt(0)
        list_bullet_style.paragraph_format.space_before = Pt(0)

def insert_paragraph_after(paragraph, text='', style=None):
    """Insert a new paragraph after the given paragraph."""
    new_p = OxmlElement('w:p')
    paragraph._element.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    if text:
        new_run = new_paragraph.add_run(text)
    if style is not None:
        new_paragraph.style = style
    return new_paragraph

def convert_lines_to_bullets(paragraph):
    """
    Converts lines in a paragraph that should be bullet points into separate paragraphs with bullet style.
    """
    lines = paragraph.text.split('\n')
    if len(lines) > 1:
        # Store the original style
        original_style = paragraph.style

        # Remove the original paragraph's text
        paragraph.text = ''

        prev_paragraph = paragraph
        for line in lines:
            line = line.strip()
            if not line:
                continue  # Skip empty lines
            # Determine the style for the new paragraph
            if line.startswith('-') or line.startswith('•') or line.startswith('*'):
                line = line.lstrip('-•*').strip()  # Remove bullet characters
                style = 'List Bullet'
            else:
                style = original_style

            # Insert new paragraph after the previous one
            new_para = insert_paragraph_after(prev_paragraph, text=line, style=style)
            prev_paragraph = new_para
    else:
        # If the paragraph is a single line starting with a bullet character
        text = paragraph.text.strip()
        if text.startswith('-') or text.startswith('•') or text.startswith('*'):
            paragraph.text = text.lstrip('-•*').strip()  # Remove bullet characters
            paragraph.style = 'List Bullet'

def clean_duration_string(duration_str):
    """
    Cleans the duration string by removing problematic characters.
    """
    if duration_str:
        # Replace backslashes with forward slashes or remove them
        duration_str = duration_str.replace('\\', '/')
        # Remove any characters that are not in the printable ASCII range
        duration_str = ''.join(c for c in duration_str if ord(c) >= 32 and ord(c) <= 126)
    return duration_str.strip()

def identify_date_format(date_str):
    """
    Identifies the date format of a given date string.
    """
    date_patterns = {
        '%d/%m/%Y': r'^\d{1,2}/\d{1,2}/\d{4}$',
        '%m/%Y': r'^\d{1,2}/\d{4}$',
        '%b %Y': r'^[A-Za-z]{3} \d{4}$',  # e.g., Jan 2020
        '%B %Y': r'^[A-Za-z]+ \d{4}$',    # e.g., January 2020
        '%Y': r'^\d{4}$',
    }

    for fmt, pattern in date_patterns.items():
        if re.match(pattern, date_str):
            return fmt
    return None  # Format not identified

def parse_end_date(duration_str):
    """
    Parses the end date from a duration string.
    """
    try:
        if not duration_str:
            logging.debug("Duration string is empty or None.")
            return datetime.min  # No duration provided, place at the end

        # Clean the duration string
        duration_str = clean_duration_string(duration_str)
        logging.debug(f"Parsing duration string: {duration_str}")

        # Handle "Present" or similar
        present_terms = ["Present", "Current", "Now", "Ongoing"]
        duration_str = duration_str.strip()
        # Split the duration into start and end
        parts = re.split(r'\s*[-–—]\s*', duration_str)
        logging.debug(f"Split parts: {parts}")
        if len(parts) == 2:
            end_str = parts[1]
        elif len(parts) == 1:
            # Only one date provided, could be end date
            end_str = parts[0]
        else:
            logging.debug("Unable to split duration string properly.")
            return datetime.min  # Unable to parse, place at the end

        end_str = end_str.strip()
        logging.debug(f"End date string: {end_str}")
        if any(term.lower() == end_str.lower() for term in present_terms):
            return datetime.now()
        else:
            # Identify the date format
            date_format = identify_date_format(end_str)
            if date_format:
                end_date = datetime.strptime(end_str, date_format)
                return end_date
            else:
                # If we cannot parse the date, place at the end
                logging.debug(f"Could not identify date format for end date: {end_str}")
                return datetime.min
    except Exception as e:
        logging.error(f"Error parsing duration '{duration_str}': {e}")
        return datetime.min  # On error, place at the end

def sort_experiences(experience_data):
    """
    Sorts the experiences from latest to oldest based on end date.
    """
    for item in experience_data:
        duration = item.get("Duration", "")
        try:
            end_date = parse_end_date(duration)
        except Exception as e:
            logging.error(f"Error parsing end date for duration '{duration}': {e}")
            end_date = datetime.min
        item['_end_date'] = end_date  # Store the end date in the item

    # Now sort the experiences based on '_end_date' field
    sorted_experiences = sorted(experience_data, key=lambda x: x.get('_end_date', datetime.min), reverse=True)

    # Remove the temporary '_end_date' field
    for item in sorted_experiences:
        item.pop('_end_date', None)

    return sorted_experiences

def insert_skills_section(paragraph, skills_data):
    """
    Inserts the skills section into the document at the position of the given paragraph.
    """
    prev_paragraph = paragraph
    if not skills_data:
        logging.warning("No skills data provided.")
        return

    for skill in skills_data:
        bullet_paragraph = insert_paragraph_after(prev_paragraph, skill, style='List Bullet')
        apply_run_font_style(bullet_paragraph.runs[0], bullet_paragraph)
        prev_paragraph = bullet_paragraph

    # Add a blank paragraph for spacing
    blank_paragraph = insert_paragraph_after(prev_paragraph, '')

    # Remove the original placeholder paragraph
    p_element = paragraph._element
    p_element.getparent().remove(p_element)

def insert_experience_section(paragraph, experience_data):
    """
    Inserts the experience section into the document at the position of the given paragraph.
    """
    if not experience_data:
        logging.warning("No experience data provided.")
        # Remove the original placeholder paragraph
        p_element = paragraph._element
        p_element.getparent().remove(p_element)
        return

    # Sort the experiences from latest to oldest
    experience_data = sort_experiences(experience_data)

    prev_paragraph = paragraph
    for item in experience_data:
        logging.debug(f"Inserting experience item: {item}")
        position = item.get("Position", "")
        company = item.get("Company", "")
        duration = item.get("Duration", "")
        responsibilities = item.get("Responsibilities", [])

        # Skip if essential fields are missing
        if not position and not company:
            logging.warning("Experience item missing 'Position' and 'Company'; skipping.")
            continue

        # Create the role title with bold formatting
        role_title_parts = [position]
        if company:
            role_title_parts.append(f"at {company}")
        if duration:
            role_title_parts.append(f"({duration})")
        role_title = ' '.join(role_title_parts)

        role_paragraph = insert_paragraph_after(prev_paragraph, '')
        role_paragraph.style = 'Normal'
        role_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        role_run = role_paragraph.add_run(role_title)
        role_run.bold = True
        apply_run_font_style(role_run, role_paragraph)

        prev_paragraph = role_paragraph

        # Add responsibilities as bullet points
        if isinstance(responsibilities, list):
            for responsibility in responsibilities:
                bullet_paragraph = insert_paragraph_after(prev_paragraph, responsibility, style='List Bullet')
                apply_run_font_style(bullet_paragraph.runs[0], bullet_paragraph)
                prev_paragraph = bullet_paragraph
        elif isinstance(responsibilities, str):
            bullet_paragraph = insert_paragraph_after(prev_paragraph, responsibilities, style='List Bullet')
            apply_run_font_style(bullet_paragraph.runs[0], bullet_paragraph)
            prev_paragraph = bullet_paragraph

        # Add a blank paragraph for spacing
        blank_paragraph = insert_paragraph_after(prev_paragraph, '')
        prev_paragraph = blank_paragraph

    # Remove the original placeholder paragraph
    p_element = paragraph._element
    p_element.getparent().remove(p_element)

def create_document(data, output_path):
    """
    Creates a Word document using the template and fills it with the extracted data.

    Parameters:
    - data (dict): The data to populate in the document.
    - output_path (str): The full path where the document will be saved.
    """
    template_path = 'Documents/Template/blank_template.docx'

    # Ensure output directory exists
    output_directory = os.path.dirname(output_path)
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Load the template document
    try:
        doc = Document(template_path)
    except Exception as e:
        logging.error(f"Failed to load template: {e}")
        raise

    # Set the default font, heading style, and language
    set_document_font(doc)
    set_heading_style(doc)
    set_document_defaults_language(doc)  # Set document-wide default language
    set_styles_language(doc)  # Set language for all styles
    set_list_bullet_style(doc)

    # Prepare placeholders, remove 'Skills' and 'Experience' since we'll handle them separately
    placeholders = {
        "{ApplicantName}": data.get("ApplicantName", ""),
        "{Role}": data.get("Role", ""),
        "{SecurityClearance}": data.get("SecurityClearance", "Not specified"),
        "{Summary}": data.get("Summary", ""),
        # "{Skills}": data.get("Skills", ""),  # Removed this line
        # "{Experience}": data.get("Experience", ""),  # Already removed
        "{Education}": data.get("Education", "")
    }

    logging.info("Starting placeholder replacement.")
    logging.debug(f"Skills data in create_document: {data.get('Skills', [])}")
    logging.debug(f"Experience data in create_document: {data.get('Experience', [])}")

    # Replace placeholders in paragraphs
    for paragraph in doc.paragraphs:
        if '{Skills}' in paragraph.text:
            # Insert the skills section here
            insert_skills_section(paragraph, data.get('Skills', []))
            # No need to remove the placeholder paragraph here (handled in insert_skills_section)
        elif '{Experience}' in paragraph.text:
            # Insert the experience section here
            insert_experience_section(paragraph, data.get('Experience', []))
            # No need to remove the placeholder paragraph here (handled in insert_experience_section)
        else:
            replace_placeholders_in_paragraph(paragraph, placeholders)
            convert_lines_to_bullets(paragraph)

    # Replace placeholders in tables
    for table in doc.tables:
        replace_placeholders_in_table(table, placeholders, data)

    # Replace header placeholders and apply 'Style 1'
    replace_headers(doc)

    logging.info("Placeholder replacement completed.")

    # Set font and language for all text after placeholder replacement
    set_font_for_all_text(doc, placeholders)

    # Save the filled-in document
    try:
        doc.save(output_path)
        logging.info(f"Document saved to {output_path}")
    except Exception as e:
        logging.error(f"Failed to save document: {e}")
        raise
