import fitz  # PyMuPDF
import pandas as pd
import os
import requests
from io import BytesIO
from PIL import Image
import re

def extract_placeholder_coordinates(template_pdf):
    coordinates = {}
    pdf_document = fitz.open(template_pdf)
    pattern = re.compile(r'#\w+')  # Pattern to match placeholders starting with #

    for page_number, page in enumerate(pdf_document):
        text = page.get_text()  # Extract text from the page
        found = pattern.findall(text)  # Find all occurrences matching the pattern
        
        for placeholder in found:
            if placeholder not in coordinates:
                coordinates[placeholder] = []
            coordinates[placeholder].append((page_number, page.search_for(placeholder)))

    pdf_document.close()
    return coordinates

def download_image(image_url):
    try:
        response = requests.get(image_url)
        response.raise_for_status()  # Raise an error for bad responses
        image = Image.open(BytesIO(response.content))
        return image
    except Exception as e:
        print(f"Error downloading image from {image_url}: {e}")
        return None

def insert_image(page, image, position):
    if image:
        # Convert the image to a format suitable for PyMuPDF
        img_byte_arr = BytesIO()
        image.save(img_byte_arr, format='PNG')  # Save as PNG
        img_byte_arr.seek(0)
        
        # Insert the image into the PDF
        img_rect = fitz.Rect(position[0], position[1], position[0] + image.width, position[1] + image.height)
        page.insert_image(img_rect, stream=img_byte_arr.getvalue())

def wrap_text(text, font_size, page_width):
    """Wrap text to fit within the specified page width."""
    wrapped_lines = []
    words = text.split(' ')
    current_line = ""

    for word in words:
        # Check the width of the current line with the new word
        test_line = f"{current_line} {word}".strip()
        test_width = fitz.get_text_length(test_line, fontsize=font_size)

        if test_width <= page_width:
            current_line = test_line
        else:
            wrapped_lines.append(current_line)
            current_line = word  # Start a new line with the current word

    if current_line:
        wrapped_lines.append(current_line)  # Add the last line

    return wrapped_lines

def map_data_to_pdf(template_pdf, output_pdf, replacements, coordinates):
    # Open the template PDF
    fresh_pdf_document = fitz.open(template_pdf)

    for page_number, page in enumerate(fresh_pdf_document):
        for placeholder, instances in coordinates.items():
            if placeholder in replacements:
                actual_value = str(replacements[placeholder])
                
                # Check if the current page has instances for this placeholder
                for inst_page_number, inst_list in instances:
                    if inst_page_number == page_number:  # Only insert on the correct page
                        # Track if we've inserted this placeholder already
                        inserted_positions = set()
                        for inst in inst_list:
                            # Check if this instance has already been inserted
                            if inst not in inserted_positions:
                                # Insert the new text at the calculated position
                                x0, y0, x1, y1 = inst
                                new_x = x0
                                new_y = y0 + (y1 - y0) / 2  # Center vertically
                                
                                # Only insert text if the placeholder is not #Image
                                if placeholder != '#Image':
                                    if placeholder in ['#Opinion', '#MdOpinion']:
                                        # Wrap the text for long placeholders
                                        page_width = 480 # Width of the placeholder
                                        wrapped_lines = wrap_text(actual_value, font_size=12, page_width=page_width)
                                        for line in wrapped_lines:
                                            page.insert_text((new_x, new_y), line, fontsize=12, color=(0, 0, 0))
                                            new_y += 15  # Move down for the next line
                                    else:
                                        page.insert_text((new_x, new_y), actual_value, fontsize=12, color=(0, 0, 0))
                                
                                # Mark this instance as inserted
                                inserted_positions.add(inst)

        # Insert images only on the second page at the position of #Image placeholder
        if page_number == 1 and '#Image' in replacements and replacements['#Image']:
            image_url = replacements['#Image']
            image = download_image(image_url)
            if image:
                # Resize the image to 350px*450px
                image = image.resize((118, 165))
                
                # Get the position of the #Image placeholder
                for inst_page_number, inst_list in coordinates['#Image']:
                    if inst_page_number == page_number:
                        for inst in inst_list:
                            # Use the position of the #Image placeholder
                            x0, y0, x1, y1 = inst
                            image_position = (x0, y0)  # Set image position to the placeholder's position
                            insert_image(page, image, image_position)
                            break  # Exit after inserting the image at the first instance

    # Save the modified PDF
    fresh_pdf_document.save(output_pdf)
    fresh_pdf_document.close()

def main():
    # Load student data
    student_data = pd.read_csv('students.csv')

    # Extract placeholders from the template PDF
    coordinates = extract_placeholder_coordinates('templatewithplaceholder.pdf')

    # Create the output directory if it doesn't exist
    output_dir = 'StudentPdfs'
    os.makedirs(output_dir, exist_ok=True)

    for index, student in student_data.iterrows():
        # Prepare the replacements dictionary
        replacements = {}
        for placeholder in coordinates:
            if placeholder in student:
                if placeholder == '#Image':
                    # Generate the image URL based on the last four digits of #UidNumber
                    uid_number = str(student['#UidNumber'])
                    last_four_digits = uid_number[-4:]
                    image_url = f"https://sbpsranchi.in/Stn/stnImg1920/S-{last_four_digits}.jpg"
                    replacements[placeholder] = image_url
                else:
                    replacements[placeholder] = student[placeholder]
        
        # Generate output PDF name with the new directory
        output_pdf = os.path.join(output_dir, f"{student['#Name']}_certificate.pdf")
        
        # Map data to the PDF
        map_data_to_pdf('templatewithoutplaceholder.pdf', output_pdf, replacements, coordinates)

if __name__ == "__main__":
    main()