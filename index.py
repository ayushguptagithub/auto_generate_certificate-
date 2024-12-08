import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas

# Load the Excel data
excel_file = "certificates_data.xlsx"  # Maam you can replace with your Excel file name
data = pd.read_excel(excel_file)

# I have just tried it so please create a canvas according to your requiremnet maam
def create_certificate(name, certificate_id, date, output_file):
    # in python I know canvas to create basic one so I did if u want I ccan try Some thing different
    c = canvas.Canvas(output_file, pagesize=landscape(letter))
    
    # I have designed very basic certificate just to try
    c.setFont("Helvetica-Bold", 36)
    c.drawString(200, 500, "Certificate of Completion")
    
    c.setFont("Helvetica", 24)
    c.drawString(180, 400, f"This is to certify that {name}")
    
    c.setFont("Helvetica", 18)
    c.drawString(180, 350, f"has successfully completed the program on {date}.")
    c.drawString(180, 300, f"Certificate ID: {certificate_id}")
    
    c.save()


for index, row in data.iterrows():
    name = row["Name"]
    certificate_id = row["Certificate ID"]
    date = row["Date"]
    
    output_file = f"Certificate_{name}.pdf"
    create_certificate(name, certificate_id, date, output_file)
    print(f"Generated: {output_file}")

print("All certificates have been generated!")
