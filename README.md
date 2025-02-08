# Excel2Cert

Excel2Cert is a Python-based certificate generator that reads participant details from an Excel file and generates personalized certificates in PDF format.

## Features
- Reads participant data from an Excel file (`data.xlsx`).
- Overlays text on a pre-designed certificate template (`CERTIFICATES.pdf`).
- Generates certificates for each participant with details like Name, Registration Number, Batch Number, Parent Name, etc.
- Saves certificates in a folder named after the Batch Number.
- Uses `PyMuPDF` (fitz) and `ReportLab` for PDF manipulation.

## Installation

### Prerequisites
Ensure you have Python installed (Python 3.x recommended). Install the required dependencies:
```sh
pip install pandas pymupdf reportlab openpyxl
```

## Usage

1. **Prepare the Excel File (`data.xlsx`)**
   - Ensure the Excel file contains the required columns:
     - `Regd_No` (Registration Number)
     - `Batch_No` (Batch Number)
     - `Name` (Participant Name)
     - `Parent_Name` (Parent Name)
     - `Resident` (Address/Location)
     - `Training_Programme` (Course Name)
     - `Start_Date` (Format: `dd/mm/yyyy`)
     - `End_Date` (Format: `dd/mm/yyyy`)
     - `Date` (Certificate Issue Date, Format: `dd/mm/yyyy`)

2. **Run the Script**
```sh
python generate_certificates.py
```

3. **Output**
   - Certificates will be generated inside folders named after each batch.
   - Example: If a participant is from Batch 2024, their certificate will be saved in `2024/certificate_John_Doe.pdf`.

## GitHub Setup (Optional)
To push your project to GitHub:
```sh
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/ubaidzargar/Excel2Cert.git
git push -u origin main
```

## License
This project is open-source and available under the MIT License.

## Author
[ubaidzargar](https://github.com/ubaidzargar)
