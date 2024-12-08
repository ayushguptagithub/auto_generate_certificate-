import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

# Load the Excel data
excel_file = "certificates_data.xlsx"  # Replace with your Excel file name
data = pd.read_excel(excel_file)

# Define a function to create a certificate
def create_certificate(name, certificate_id, date, output_file):
    # Set up the PDF canvas
    c = canvas.Canvas(output_file, pagesize=landscape(letter))
    
    # Add a background image (replace with your own image path)
    background_image = "certificate_bg.png"  # Path to your background image
    c.drawImage(background_image, 0, 0, width=landscape(letter)[0], height=landscape(letter)[1])

    # Add the certificate title
    c.setFillColor(HexColor("#2E4053"))  # Dark blue color
    c.setFont("Helvetica-Bold", 40)
    c.drawCentredString(415, 450, "Certificate of Completion")
    
    # Add recipient name
    c.setFillColor(HexColor("#117A65"))  # Greenish color
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(415, 350, f"{name}")
    
    # Add completion text
    c.setFillColor(HexColor("#000000"))  # Black color
    c.setFont("Helvetica", 20)
    c.drawCentredString(415, 300, "has successfully completed the program on")
    
    # Add date
    c.setFillColor(HexColor("#7D3C98"))  # Purple color
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(415, 250, f"{date}")
    
    # Add certificate ID
    # Add certificate ID with italic style
    c.setFillColor(HexColor("#616A6B"))  # Gray color
    c.setFont("Helvetica-Oblique", 16)  # Use Helvetica-Oblique for italic text
    c.drawCentredString(415, 200, f"Certificate ID: {certificate_id}")


    # Save the PDF
    c.save()

# Iterate over each row in the Excel file and generate certificates
for index, row in data.iterrows():
    name = row["Name"]
    certificate_id = row["Certificate ID"]
    date = row["Date"]
    
    output_file = f"Certificate_{name}.pdf"
    create_certificate(name, certificate_id, date, output_file)
    print(f"Generated: {output_file}")

print("All certificates have been generated!")
