import os
import pandas as pd
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Load Excel Data
df = pd.read_excel("data.xlsx")  # Make sure Excel has the required columns

# Certificate template (Canva PDF)
template_pdf = "CERTIFICATES.pdf"

# Function to overlay text on the PDF
def create_certificate(data, output_pdf):
    # Create a temporary PDF for text overlay
    temp_pdf = "temp.pdf"
    c = canvas.Canvas(temp_pdf, pagesize=letter)

    # Set font and size
    c.setFont("Helvetica-Bold", 14)

    # Draw text on certificate at exact coordinates
    c.drawString(120, 775, str(data.get("Regd_No", "")))  # Regd. No first
    c.drawString(520, 775, str(data.get("Batch_No", "")))  # Then Batch No
    c.drawString(340, 440, str(data.get("Name", "")))  # Name of Candidate
    c.drawString(230, 400, str(data.get("Parent_Name", "")))  # Default to empty string if not found
    c.drawString(130, 360, str(data.get("Resident", "")))  # Resident of
    c.drawString(350, 315, str(data.get("Training_Programme", "")))  # Training Programme
    c.drawString(200, 270, data.get("Start_Date", "").strftime("%d/%m/%Y") if pd.notna(data.get("Start_Date", "")) else "")  # From (Start Date)
    c.drawString(400, 270, data.get("End_Date", "").strftime("%d/%m/%Y") if pd.notna(data.get("End_Date", "")) else "")  # To (End Date)
    c.drawString(450, 230, "NRLM")  # Sponsored by NRLM (Fixed Value)
    c.drawString(80, 75, data.get("Date", "").strftime("%d/%m/%Y") if pd.notna(data.get("Date", "")) else "")  # Date (Bottom Right Corner)

    c.save()

    # Open the template and overlay PDFs
    template_doc = fitz.open(template_pdf)
    overlay_doc = fitz.open(temp_pdf)

    # Combine the pages (overlay the text onto the template)
    page = template_doc.load_page(0)  # Load the first page of the template
    page.show_pdf_page(page.rect, overlay_doc, 0)  # Show the overlay on top of the template page

    # Save the final PDF
    template_doc.save(output_pdf)

# Generate certificates for each row in Excel
for index, row in df.iterrows():
    batch_folder = str(row["Batch_No"])  # Get Batch No as folder name
    if not os.path.exists(batch_folder):
        os.makedirs(batch_folder)  # Create folder if it doesn't exist
    
    output_filename = os.path.join(batch_folder, f"certificate_{row['Name'].replace(' ', '_')}.pdf")
    create_certificate(row, output_filename)

print("✅ Certificates Generated Successfully!")
