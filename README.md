AuditRAM Highlighter is a lightweight and efficient tool built using Python + Streamlit that allows users to highlight specific words inside multiple file formats — without modifying the original file.

This tool is perfect for audits, inspections, document analysis, and compliance workflows where quick text identification is required.

 Features
 Word Highlighting Across Multiple Formats

 PDF (text-based) → Red bounding box

 Scanned PDF → OCR + red bounding box

 Images (PNG/JPG) → OCR + red bounding box

 Word (.docx) → Yellow text highlight

 Excel (.xlsx) → Red bold color highlight for matched text

 Safe & Non-Destructive

Your original file never gets modified.
The tool generates a new highlighted version for download.

Simple Web-Based UI

Runs locally using Streamlit.
No complex setup or coding needed.

Supported File Types
Format	Extension	Highlight Method
Text PDF	.pdf	Bounding box
Scanned PDF	.pdf	OCR + bounding box
Image	.png, .jpg, .jpeg	OCR + bounding box
Word Document	.docx	Yellow highlight
Excel Sheet	.xlsx	Red-bold word highlight
⚙️ Installation Guide

Follow these steps carefully to install all required dependencies and run the program successfully.

 Install Python

Make sure Python 3.8+ is installed.
Download from: https://www.python.org/downloads/

Install Required Libraries

Run this command in terminal / CMD / PowerShell:

pip install streamlit pymupdf pillow pytesseract openpyxl python-docx

Install Tesseract OCR (Required for Images & Scanned PDFs)
For Windows:

Download installer:
https://github.com/UB-Mannheim/tesseract/wiki

After installation, ensure this path exists:

C:\Program Files\Tesseract-OCR\tesseract.exe


If not recognized, add it to Environment Variables.

 Project Structure

Your folder should look like:

AuditRAM-Highlighter/
   ├── app.py
   └── README.md

 How to Run the Application

In your terminal, navigate to your project folder:

cd AuditRAM-Highlighter


Run the Streamlit app:

streamlit run app.py


Your browser will automatically open at:

http://localhost:8501

🖱️ How to Use the Tool

Click "Upload File"

Select your file (PDF / Image / DOCX / XLSX)

Enter the word you want to highlight

Click “Highlight”

Download the processed file

 Output Examples
Input	Output
PDF	PDF with red bounding boxes
Image	PNG/JPG with red bounding boxes
Word	DOCX with yellow-highlighted text
Excel	XLSX with red-bold highlighted text
Troubleshooting Guide
 Error: “Tesseract not found”

Install Tesseract from the link above and restart your system.

 PDF not highlighting

Check if the word is present.
Case-insensitive, partial matching is supported.

 Excel not highlighting

Your Excel file must be .xlsx (not .xls).
Merged cells or formulas may cause matching issues.

Word not highlighting

Your file must be .docx (not .doc).

Contributing

Pull requests are welcome.
For major changes, open an issue first.

 License

This project is free and open-source.

 Credits

Made using:

Python

Streamlit

PyMuPDF

PIL

Tesseract OCR

OpenPyXL

python-docx