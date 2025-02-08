

# Excel2Cert 🚀  
### Automate Certificate Generation from Excel Data 

**Excel2Cert** is a simple yet powerful tool that automates certificate creation by extracting data from an **Excel sheet** and overlaying it onto a **PDF certificate template**.  

This eliminates **manual work**, reduces **errors**, and allows bulk certificate generation **in seconds**.  

---

## **📌 Features**  
✅ **Bulk Certificate Generation** – No need for manual edits.  
✅ **Excel to PDF Automation** – Fetches details like name, date, and batch number from an Excel file.  
✅ **Customizable Positioning** – Place text dynamically on a PDF template.  
✅ **Supports Canva Templates** – Works with non-editable PDF templates from Canva or other design tools.  
✅ **Easy Setup** – No advanced coding required.  

---

## **🛠 Installation**  

### **1️⃣ Clone the Repository**  
```sh
git clone https://github.com/ubaidzargar/Excel2Cert.git
cd Excel2Cert
```

### **2️⃣ Install Dependencies**  
Make sure you have **Python 3** installed. Then, run:  
```sh
pip install pandas pymupdf reportlab openpyxl
```

### **3️⃣ Place Your Files in the Project Folder**  
- **Excel File (`data.xlsx`)** – Contains certificate details (e.g., Name, Batch No, Date).  
- **Certificate Template (`certificate_template.pdf`)** – Your Canva-designed certificate layout.  

---

## **📜 How to Use**  

### **1️⃣ Open the Excel File (`data.xlsx`) and Fill Data**
The first row should have column headers:  
| Regd_No | Batch_No | Name | Parent_Name | Resident | Training_Programme | Start_Date | End_Date | Date |
|---------|---------|------|-------------|----------|---------------------|------------|----------|------|
| 12345   | 2024-01 | John Doe | Richard Doe | New York | Digital Marketing | 2024-01-10 | 2024-01-20 | 2024-01-21 |

### **2️⃣ Run the Script**
Once the Excel file is updated, generate certificates by running:  
```sh
python3 script.py
```

### **3️⃣ Check the Output**
The script will create **individual certificates** for each entry in the Excel file, named like:  
```
certificate_John_Doe.pdf  
certificate_Jane_Smith.pdf  
```

---

## **🛠 Customization**
- Adjust **text placement** in `script.py` by modifying the coordinates in `drawString(x, y, text)`.  
- Change **font size** or style inside the `reportlab` section.  

---

## **❓ Troubleshooting**
❌ **pip command not found?** → Install pip:  
```sh
python3 -m ensurepip --default-pip
```
❌ **Text not aligned properly?** → Adjust X, Y coordinates in `script.py`.  
❌ **Generated certificates are blank?** → Ensure `certificate_template.pdf` is not an image-based PDF.  

---

## **🤝 Contributing**
🚀 Want to improve this tool? Feel free to **fork** the repository, make changes, and submit a **pull request**!  

---

## **📜 License**
This project is **open-source** under the **MIT License**.  

---

## **📩 Connect With Me**
💼 **[LinkedIn](https://www.linkedin.com/in/yasirzargar)**  
📧 **zargaryasir@gmail.com**  

🚀 Let’s automate more and work smarter!  

---
