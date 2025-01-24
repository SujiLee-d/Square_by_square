# Square_by_square
# HOWTO: Automatically Generate Certificate PDFs with QR Codes from Excel

## **Purpose**

This guide explains how to create certificate PDF files automatically using a Python script, a pre-made PDF template, and an Excel input file containing relevant data.

---

## **Requirements**

- **Python Environment**
    - Install the required Python libraries:
        
        ```bash
        pip install qrcode pillow pandas pypdf2 reportlab
        
        ```
        
- **Required Files**
    - **Python Script**: Place the Python script in the working directory.
    - **Template PDF**: `QRcode top blank.pdf` should be in the same folder.
    - **Excel File**: The file `Data.xlsx` should contain the required input data.
- **Input Data**
The script expects an Excel file with the following structure:
    
    
    | Row Name | Description |
    | --- | --- |
    | `square_size` | The size of the square (1, 5, or 25) |
    | `lot` | Lot number of each square |
    | `QR` | URL of each square |
- **Output**
The script will generate PDFs with:
    - A unique QR code for each row.
    - Five fields filled with the corresponding data.

---

## **Input Data Example**

### Data.xlsx

| square_size | lot | QR |
| --- | --- | --- |
| 1 | 12345 | https://example.com/square/12345 |
| 5 | 23456 | https://example.com/square/23456 |
| 25 | 34567 | https://example.com/square/34567 |

---

## **Script Overview**

### Main Features

1. **Read Excel Data**: Extracts the input data from `Data.xlsx`.
2. **Generate QR Codes**: Creates a QR code for each row's URL and saves it with a transparent background.
3. **Fill Template PDF**: Places the QR code and text fields into the template PDF.
4. **Save Output**: Saves each certificate with a unique name based on the data.

---

## **Script Details**

### Key Sections

### 1. **Read Excel Data**

The function `read_excel_data` reads the Excel file into a pandas DataFrame:

```python
def read_excel_data(file_path):
    """
    Reads Excel data and returns a pandas DataFrame.
    """
    df = pd.read_excel(file_path)
    return df

```

### 2. **Generate QR Codes**

The `generate_qr_code` function creates QR codes with a transparent background:

```python
def generate_qr_code(url, qr_image_path):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)

    img = qr.make_image(fill_color=(0, 128, 0), back_color=(255, 255, 255, 0))
    img.save(qr_image_path)

```

### 3. **Fill PDF Template**

The `add_custom_text_to_pdf` function inserts the QR code and custom text into the PDF template:

```python
def add_custom_text_to_pdf(template_path, output_path, custom_data):
    reader = PdfReader(template_path)
    writer = PdfWriter()

    packet = BytesIO()
    can = canvas.Canvas(packet)
    can.setFont("Helvetica", 12)

    can.drawString(270, 265, f"{custom_data['square_no']}")
    can.drawString(270, 245, f"{custom_data['square_size']} m²")
    can.drawString(270, 225, f"{custom_data['location']}")
    can.drawString(270, 204, f"{custom_data['co2_capture']} kg pa *")
    can.drawString(270, 184, f"{custom_data['freshwater_produced']} L per annum")

    generate_qr_code(custom_data['url'], 'rgba_qr.png')
    can.drawImage('rgba_qr.png', 204, 641, width=100, height=100)

    can.save()

    packet.seek(0)
    custom_pdf = PdfReader(packet)

    for page in reader.pages:
        overlay = custom_pdf.pages[0]
        page.merge_page(overlay)
        writer.add_page(page)

    with open(output_path, "wb") as output_pdf:
        writer.write(output_pdf)

```

### 4. **Process Excel Data**

The `process_excel_to_pdf` function processes each row in the Excel file to generate PDFs:

```python
def process_excel_to_pdf(template_path, excel_path):
    df = read_excel_data(excel_path)

    for index, row in df.iterrows():
        custom_data = {
            "square_no": f"002-{int(row['square_size']):02}-{int(row['lot']):04}",
            "square_size": f"{row['square_size']}",
            "location": f"Djarawong, QLD, Lot {row['lot']}",
            "co2_capture": f"{1.5 * int(row['square_size'])}",
            "freshwater_produced": f"{1100 * int(row['square_size'])}",
            "url": f"{row['QR']}"
        }

        output_path = f"certificate_{custom_data['square_no']}.pdf"
        add_custom_text_to_pdf(template_path, output_path, custom_data)
        print(f"Generated PDF: {output_path}")

```

---

## **How to Run the Script**

1. Place the Python script, `Data.xlsx`, and `QRcode top blank.pdf` in the same directory.
2. Run the script:
    
    ```bash
    python generate_certificate.py
    
    ```
    
3. Generated PDF files will be saved in the same directory with names like `certificate_002-01-0001.pdf`.

---

## **To-Do**

- Ensure QR code background is transparent.
- Customise font style or size as needed.

---

If you encounter any issues, feel free to reach out for support!
