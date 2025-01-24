import qrcode
from PIL import Image
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from io import BytesIO

# Read excel file and return a pandas DataFrame
def read_excel_data(file_path):
    """
    Reads Excel data and returns a pandas DataFrame.
    """
    df = pd.read_excel(file_path)
    return df


def generate_qr_code(url, qr_image_path):
    """
    Generate a QR code from a URL and save it as an image.
    """
    qr = qrcode.QRCode(
        version=1,  # Controls the size of the QR Code (1 is smallest, 40 is largest)
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="#A6E987", back_color="#FFFFFF")

    # img = qr.make_image(fill_color="#00FF00", back_color="transparent")
    img.save(qr_image_path)

    # Make the background transparent
    # data = img.getdata()
    # new_data = []
    # for item in data:
    #     # If pixel is white, make it transparent
    #     if item[:3] == (255, 255, 255):  # RGB for white
    #         new_data.append((255, 255, 255, 0))  # RGBA for transparent
    #     else:
    #         new_data.append(item)

    # img.putdata(new_data)
    # img.save(qr_image_path)
    ########

    # Step 3: Process to Make Background Transparent
    image = Image.open('qr_code.png')
    rgba = image.convert("RGBA")
    data = rgba.getdata()
    new_data = []
    for item in data:
        # Replace white pixels with transparent ones
        # if item[0] == 0 and item[1] == 0 and item[2] == 0: # Turn black pixels 
        if item[0] == 255 and item[1] == 255 and item[2] == 255: # Turn white pixels 
        # if item[:3] == (255, 255, 255):  # RGB for white
            new_data.append((255, 255, 255, 0))  # into Transparent pixels
        else:
            new_data.append(item)  # QR Code color with full opacity

    rgba.putdata(new_data)
    #rgba.save(qr_image_path, "PNG")
    rgba.save('rgba_qr.png', "PNG")
    ###########################
def add_custom_text_to_pdf(template_path, output_path, custom_data):
    """
    Add custom text to a pre-made PDF template.

    Args:
        template_path (str): Path to the pre-made PDF template.
        output_path (str): Path to save the modified PDF.
        custom_data (dict): Dictionary containing custom text fields and their values.
    """
    # Step 1: Load the pre-made PDF template
    reader = PdfReader(template_path)
    writer = PdfWriter()

    # Step 2: Create a new PDF in memory with custom data
    packet = BytesIO()
    can = canvas.Canvas(packet)
    can.setFont("Helvetica", 12)
    
    # Example custom data positions
    can.drawString(270, 265, f"{custom_data['square_no']}")
    can.drawString(270, 245, f"{custom_data['square_size']} m²")
    can.drawString(270, 225, f"{custom_data['location']}")
    can.drawString(270, 204, f"{custom_data['co2_capture']} kg pa *")
    can.drawString(270, 184, f"{custom_data['freshwater_produced']} L per annum")

    # Watermark
    can.setFont("Helvetica-Bold", 70)
    can.saveState()
    can.setFillColorRGB(245/255, 245/255, 220/255)
    can.translate(200, 400)  # position
    can.rotate(45)           # rotate 45'
    can.drawString(0, 0, "sample only")  # Text
    can.restoreState()
        

    #generate_qr_code(custom_data['url'], qr_image_path)
    generate_qr_code(custom_data['url'], 'rgba_qr.png')
    can.drawImage(qr_image_path, qr_position[0], qr_position[1], width=100, height=100)  # Adjust width/height as needed
    


    can.save()
    
    # Merge the custom PDF with the template
    packet.seek(0)
    custom_pdf = PdfReader(packet)
    
    for page in reader.pages:
        overlay = custom_pdf.pages[0]
        page.merge_page(overlay)
        writer.add_page(page)

    # Step 4: Save the modified PDF
    with open(output_path, "wb") as output_pdf:
        writer.write(output_pdf)

# Example Usage
# square_size = "25"
# serial_no="1"
# template_path = "QRcode top blank.pdf"  # Path to your pre-made template
# output_path = "customized_certificate9.pdf"  # Output path for the customized PDF
# custom_data = {
#     "square_no": f"002-{square_size}-{serial_no} m2",
#     "square_size": "25",
#     "location": "Djarawong, QLD, Lot 1RP807619",
#     "co2_capture": "37.5",
#     "freshwater_produced": "27,500"
# }

def process_excel_to_pdf(template_path, excel_path):
    """
    Process each row from Excel data and generate PDFs with custom data.
    """
    df = read_excel_data(excel_path)
    
    
    # Convert each row of the dataframe into custom_data -> create PDF
    for index, row in df.iterrows():
        # custom_data 생성
        custom_data = {
            "square_no": f"002-{int(row['square_size']):02}-{int(row['sales_no_per_square']):04} ",
            "square_size": f"{row['square_size']}",
            "location": f"Djarawong, QLD, Lot {row['lot']}",
            "co2_capture": f"{1.5*int(row['square_size'])}",
            "freshwater_produced": f"{1100*int(row['square_size'])}",
            "url": f"{row['QR']}"
        }

        # Create PDF for every row
        output_path = f"certificate_{custom_data['square_no']}.pdf"
        add_custom_text_to_pdf(template_path, output_path, custom_data)
        print(f"Generated PDF: {output_path}")

# input values setting
excel_path = 'Data2.xlsx'
template_path = "QRcode top blank.pdf"  # Path to your pre-made template
qr_image_path = "qr_code.png"
qr_position=(204, 641)


process_excel_to_pdf(template_path, excel_path)
#add_custom_text_to_pdf(template_path, output_path, custom_data)
print("pdfs created")
