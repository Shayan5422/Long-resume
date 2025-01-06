# extract_text.py

import os
import PyPDF2
import pdfplumber
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar, LTFigure
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_text_from_pdf_only(pdf_path):
    """استخراج متن از یک فایل PDF و بازگرداندن متن به صورت رشته."""
    text_per_page = {}
    image_flag = False
    try:
        pdfFileObj = open(pdf_path, 'rb')
        pdfReader = PyPDF2.PdfReader(pdfFileObj)
    except Exception as e:
        logger.error(f"Failed to open PDF {pdf_path}: {e}")
        return ""

    try:
        for pagenum, page_layout in enumerate(extract_pages(pdf_path)):
            try:
                pageObj = pdfReader.pages[pagenum]
            except IndexError:
                logger.warning(f"Page {pagenum} does not exist in PDF {pdf_path}.")
                continue

            page_text = []
            text_from_images = []
            page_content = []

            try:
                pdf = pdfplumber.open(pdf_path)
                page_tables = pdf.pages[pagenum]
                tables = page_tables.find_tables()
            except Exception as e:
                logger.error(f"Failed to extract tables from page {pagenum}: {e}")
                tables = []

            # Extract tables
            for table_num in range(len(tables)):
                try:
                    table = extract_table(pdf_path, pagenum, table_num)
                    table_string = table_converter(table)
                    page_content.append(table_string)
                except Exception as e:
                    logger.error(f"Failed to extract table {table_num} from page {pagenum}: {e}")
                    continue

            # Sort elements by Y position (descending)
            page_elements = [(element.y1, element) for element in page_layout]
            page_elements.sort(key=lambda a: a[0], reverse=True)

            for component in page_elements:
                element = component[1]

                # Extract text elements
                if isinstance(element, LTTextContainer):
                    try:
                        line_text, _ = text_extraction(element)
                        page_text.append(line_text)
                        page_content.append(line_text)
                    except Exception as e:
                        logger.error(f"Failed to extract text from element on page {pagenum}: {e}")

                # Extract image elements (if any)
                if isinstance(element, LTFigure):
                    try:
                        crop_image(element, pageObj)
                        convert_to_images('cropped_image.pdf')
                        image_text = image_to_text('PDF_image.png')
                        text_from_images.append(image_text)
                        page_content.append(image_text)
                        image_flag = True
                    except Exception as e:
                        logger.error(f"Failed to extract text from image on page {pagenum}: {e}")

            # Combine the extracted content for each page
            dctkey = f'Page_{pagenum}'
            text_per_page[dctkey] = page_content

    except Exception as e:
        logger.error(f"Failed to extract text from PDF {pdf_path}: {e}")
    finally:
        pdfFileObj.close()
        if image_flag:
            if os.path.exists('cropped_image.pdf'):
                os.remove('cropped_image.pdf')
            if os.path.exists('PDF_image.png'):
                os.remove('PDF_image.png')

    # Combine all text from all pages into one string
    combined_text = '\n'.join(['\n'.join(text_per_page[page]) for page in text_per_page])
    return combined_text

def extract_text_from_pdf(pdf_path, output_path):
    """استخراج متن از یک فایل PDF و ذخیره آن در فایل متنی خروجی."""
    text = extract_text_from_pdf_only(pdf_path)
    try:
        with open(output_path, 'w', encoding='utf-8') as output_file:
            output_file.write(text)
        logger.info(f"Extracted text from {pdf_path} saved to {output_path}")
    except Exception as e:
        logger.error(f"Failed to write extracted text to {output_path}: {e}")
    return text

def extract_text_from_folder(folder_path, output_txt_path):
    combined_text = ''
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            logger.info(f"Extracting text from: {pdf_path}")
            pdf_text = extract_text_from_pdf_only(pdf_path)
            combined_text += f"\n--- Text from {filename} ---\n" + pdf_text + "\n"

    # Save the combined text to a .txt file
    try:
        with open(output_txt_path, 'w', encoding='utf-8') as output_file:
            output_file.write(combined_text)
        logger.info(f"All text extracted and saved to {output_txt_path}")
    except Exception as e:
        logger.error(f"Failed to write extracted text to {output_txt_path}: {e}")

    return combined_text

def crop_image(element, pageObj):
    image_left, image_top, image_right, image_bottom = element.x0, element.y0, element.x1, element.y1
    pageObj.mediabox.lower_left = (image_left, image_bottom)
    pageObj.mediabox.upper_right = (image_right, image_top)

    cropped_pdf_writer = PyPDF2.PdfWriter()
    cropped_pdf_writer.add_page(pageObj)
    with open('cropped_image.pdf', 'wb') as cropped_pdf_file:
        cropped_pdf_writer.write(cropped_pdf_file)

def convert_to_images(input_file):
    try:
        images = convert_from_path(input_file)
        if images:
            image = images[0]
            output_file = 'PDF_image.png'
            image.save(output_file, 'PNG')
            logger.info(f"Converted {input_file} to image {output_file}.")
    except Exception as e:
        logger.error(f"Failed to convert {input_file} to image: {e}")

def image_to_text(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        logger.info(f"Extracted text from image {image_path}.")
        return text
    except Exception as e:
        logger.error(f"Failed to extract text from image {image_path}: {e}")
        return ""

def text_extraction(element):
    line_text = element.get_text()
    line_formats = []
    for text_line in element:
        if isinstance(text_line, LTChar):
            line_formats.append(text_line.fontname)
            line_formats.append(text_line.size)
    format_per_line = list(set(line_formats))
    return (line_text, format_per_line)

def extract_table(pdf_path, page_num, table_num):
    try:
        pdf = pdfplumber.open(pdf_path)
        table_page = pdf.pages[page_num]
        table = table_page.extract_tables()[table_num]
        return table
    except Exception as e:
        logger.error(f"Failed to extract table {table_num} from page {page_num}: {e}")
        return []

def table_converter(table):
    table_string = ''
    try:
        for row in table:
            cleaned_row = [item.replace('\n', ' ') if item is not None and '\n' in item else ('None' if item is None else item) for item in row]
            table_string += ('|' + '|'.join(cleaned_row) + '|\n')
        table_string = table_string.rstrip('\n')
        return table_string
    except Exception as e:
        logger.error(f"Failed to convert table to string: {e}")
        return ""

if __name__ == "__main__":
    folder_path = "path_to_your_pdfs"
    output_txt_path = "extracted_text.txt"
    extract_text_from_folder(folder_path, output_txt_path)
