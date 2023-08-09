from bs4 import BeautifulSoup
import os
from translate import Translator
import fitz  # PyMuPDF
from PIL import Image

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib import utils
from reportlab.pdfgen import canvas

config = {
    "sending-note":
        {
            "margin":
                {
                    "left": 0,
                    "top": 0,
                    "right": 0,
                    "bottom": 700
                }
        },
    "address":
        {
            "margin":
                {
                    "left": 50,
                    "top": 75,
                    "right": 0,
                    "bottom": 575
                },
            "language-code":"fr",
            "type":
                {
                "font": "Helvetica-Bold",
                "size": "14",
                "color": "black"
                },
            "translated-type":
                {
                "font": "Helvetica-BoldOblique",
                "size": "14",
                "color": "darkgrey"
                },
            "content": 
                {
                "font": "Helvetica",
                "size": "14",
                "color": "black"
                },
            "separator": 
                {
                "font": "Helvetica",
                "size": "14",
                "color": "black"
                },
            "replacement": 
                {"Copier dans le Presse-papier":"Téléphone",
                 "Copier dans Presse-papier":"Rue",
                 "Copier le ":"",
                 "Copier la ":"",
                 "Copier l'":"",
                 "le ":"",
                 "la ":""}
            
        },
    "output-name":"output.pdf"
}


# Function to find the first HTML or TXT file in the folder
def find_html_or_txt_file(folder):
    for filename in os.listdir(folder):
        if filename.lower().endswith(('.html', '.txt')):
            return os.path.join(folder, filename)
    return None

# Function to find the first PDF file in the folder


def find_pdf_file(folder):
    for filename in os.listdir(folder):
        if filename.lower().endswith(('.pdf')) and filename != config["output-name"]:
            return os.path.join(folder, filename)
    return None


def get_type_content(soup, target):
    target_block = soup.find('div', class_=target)
    if target_block:
        button_elements = target_block.find_all('button', class_=button_class)
        button_types = target_block.find_all('span', class_=type_class)

        if not button_elements:
            print(
                f"No buttons with class '{button_class}' found in the extracted block.")
            return None
        if not button_types:
            print(
                f"No spans with class '{type_class}' found in the extracted block.")
            return None
        return zip(button_elements, button_types)
    else:
        print(f"No element with class '{target}' found.")
    return None


def retrieve_data_from_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()
        data = []
        soup = BeautifulSoup(html_content, 'html.parser')
        translator = Translator(to_lang='en', from_lang=config["address"]["language-code"])

        address = get_type_content(soup, 'address')
        
        if not address:
            print("Can't find address")
            return None
        
        for button_element, button_type in address:
            content = ' '.join(button_element.stripped_strings)
            type = ' '.join(button_type.stripped_strings)
            for search,replacement in config['address']['replacement'].items():
                type = type.replace(search, replacement)
            type = type.capitalize()
            translated_type = translator.translate(type)
            data.append(
                {"type": type, "translated_type": translated_type, "content": content})

        phone = get_type_content(soup, 'phone')
        if not phone:
            print("Can't find phone number")
            return None
        
        for button_element, button_type in phone:
            content = ' '.join(button_element.stripped_strings)
            type = ' '.join(button_type.stripped_strings)
            for search,replacement in config['address']['replacement'].items():
                type = type.replace(search, replacement)
            translated_type = translator.translate(type)
            data.append(
                {"type": type, "translated_type": translated_type, "content": content})

    return data


def write_paragraph(data):
    # Initialize style for the line
    line_style = ParagraphStyle(
        name='LineStyle',
        fontSize=0,  # Set font size to 0 to ensure no text is displayed
        leading=22,
        alignment=0
    )

    # Create a list of Paragraph objects with the specified styles
    paragraphs = []
    for entry in data:
        line = Paragraph(
            f"<font name='{config['address']['type']['font']}' size='{config['address']['type']['size']}' color='{config['address']['type']['color']}'>{entry['type']}</font>"
            f"<font name='{config['address']['separator']['font']}' size='{config['address']['separator']['size']}' color='{config['address']['separator']['color']}'>/</font>"
            f"<font name='{config['address']['translated-type']['font']}' size='{config['address']['translated-type']['size']}' color='{config['address']['translated-type']['color']}'>{entry['translated_type']}</font>"
            f"<font name='{config['address']['separator']['font']}' size='{config['address']['separator']['size']}' color='{config['address']['separator']['color']}'>: </font>"
            f"<font name='{config['address']['content']['font']}' size='{config['address']['content']['size']}' color='{config['address']['content']['color']}'>{entry['content']}</font>",
            line_style
        )
        paragraphs.append(line)

        # Add the paragraphs and spacer to the elements list
        # paragraphs.extend([type_paragraph, spacer, translated_type_paragraph, spacer, content_paragraph])

    return paragraphs


def pdf_to_image(pdf_path, output_folder, name, margin):

    pdf_document = fitz.open(pdf_path)

    page = pdf_document.load_page(0)
    zoom = 4  # zoom factor
    mat = fitz.Matrix(zoom, zoom)
    pixmap = page.get_pixmap(matrix=mat)
    image = Image.frombytes(
        "RGB", [pixmap.width, pixmap.height], pixmap.samples)
    w, h = image.size
    image = image.crop((zoom*margin[0], zoom*margin[1], w-zoom*margin[2], h-zoom*margin[3]))
    image_path = f"{output_folder}/{name}.png"
    image.save(image_path, "PNG")

    pdf_document.close()

    return image_path


def merge_images_vertically(image_path1, image_path2, save_path):
    image1 = Image.open(image_path1)
    image2 = Image.open(image_path2)

    width = min(image1.width, image2.width)
    height = image1.height + image2.height

    merged_image = Image.new('RGB', (width, height))
    merged_image.paste(image1, (0, 0))
    merged_image.paste(image2, (0, image1.height))
    merged_image.save(save_path, "PNG")
    w, h = merged_image.size
    return save_path, w, h


def create_pdf_from_images(image_path, output_pdf, w, h):
    c = canvas.Canvas(output_pdf, pagesize=A4)

    width, height = A4
    img = utils.ImageReader(image_path)
    c.drawImage(img, 0, 0, width, h/4)

    c.save()


print("Before using the program, make sure that you have created a folder with one and only one .txt or .html containing the ebay html code, and one and only one .pdf file containing the stamp.")
folder_path = input("Please type in the path to the folder: ")
button_class = 'tooltip__host'
type_class = 'tooltip__content'


# Find the first HTML or TXT file
file_path = find_html_or_txt_file(folder_path)

pdf_path = find_pdf_file(folder_path)

temp_path = "temp.pdf"

if file_path:
    data = retrieve_data_from_html(file_path)
else:
    print("No HTML or TXT file found in the specified folder.")
    
if data:
    

    paragraph = write_paragraph(data)

    if pdf_path:
        # Create a PDF writer for the generated content
        generated_doc = SimpleDocTemplate(temp_path, pagesize=A4)
        generated_doc.build(paragraph)

        margin1 = [config['sending-note']['margin']['left'], config['sending-note']['margin']['top'], config['sending-note']['margin']['right'], config['sending-note']['margin']['bottom']]
        margin2 = [config['address']['margin']['left'], config['address']['margin']['top'], config['address']['margin']['right'], config['address']['margin']['bottom']]
        image_path1 = pdf_to_image(pdf_path, folder_path, "pdf", margin1)
        image_path2 = pdf_to_image(temp_path, folder_path, "address", margin2)


        merged_path, w, h = merge_images_vertically(
            image_path1, image_path2, folder_path+"merged.png")

        create_pdf_from_images(merged_path, folder_path + "/" + config["output-name"], w, h)

        print(f"PDF generated and saved as 'output.pdf'")
    else:
        print("No PDF file found in the specified folder.")
else:
    print("Error when retrieving data, exiting without creating the document.")


input("Press Enter to continue...")
