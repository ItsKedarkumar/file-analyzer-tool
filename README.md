📁 File Analyzer Tool

A Python-based File Analyzer Tool that processes Text (.txt), CSV (.csv), and PDF (.pdf) files, performs data analysis, uses OCR to extract Aadhaar details, generates reports and charts, and exports results to Excel for automation and documentation.

🔥 Features

✔ Analyze Text files (word count, frequent words, email/number detection)
✔ Analyze CSV files (statistical summary using pandas)
✔ Scan PDF using OCR (Aadhaar scan support)
✔ Search keywords in text files
✔ Detect special characters
✔ Export full analysis to Excel
✔ Generate a final Project Summary report
✔ Fully terminal-based menu-driven tool

📦 Project Structure
FILE ANALYZER TOOL/
├── file_analyzer.py         # Main program
├── Output/                  # Generated reports, charts, Excel files
├── samples/                 # Sample test files
├── test_files/              # Development test data
├── test_pdfs/               # Aadhaar test PDF (if scanned)
├── analysis_output.xlsx     # Excel summary (generated)
├── final_project_summary.txt# Final project summary
├── aadhar_output.xlsx       # OCR output (if scanned)
├── README.md                # Project documentation
└── .gitignore               # Git ignored files

🔧 Requirements

Install the required libraries before running the tool:

pip install pandas matplotlib python-docx openpyxl pytesseract pillow


⚠ For Aadhaar (OCR) scanning:

Install Tesseract OCR

Add Tesseract.exe path to System Environment Variables

▶️ How to Run

1️⃣ Navigate to the project directory

cd "C:\kedar\python\Mini project\FILE ANALYZER TOOL"


2️⃣ Run the script

python file_analyzer.py


3️⃣ Use the menu options

==== FILE ANALYZER TOOL ====
1) Analyze Text File & Generate Report
2) Analyze CSV File
3) Scan PDF for Aadhaar
4) Search Keyword in Text File
5) Analyze Special Characters
6) Export All Analysis to Excel
7) Generate Final Summary Report
8) Exit



TEXT FILE ANALYSIS RESULT
-------------------------
Total Words: 12
Unique Words: 10
Most Frequent Word: is (2)
Emails Found: ['example@test.com']
Numbers Found: ['9876543210']

📊 Final Project Summary Example
FILE ANALYZER TOOL 🧪 FINAL PROJECT SUMMARY
Developer : Kedar Kumar Trivedi
Version   : v1.0 (Mini Project | SEM 4)

TEXT ANALYSIS
- Total Words: 12
- Unique Words: 10

CSV ANALYSIS
- Column Names: ['name', 'age', 'marks']
- Age Mean = 18.5, Marks Mean = 90.0

OCR Aadhaar
Status: Extracted successfully (if scanned)

🎯 Future Enhancements

🔹 Add GUI using Tkinter / PyQt
🔹 Support more file formats (JSON, XML)
🔹 Direct database export
🔹 Email report automation

💡 Developed By

👨‍💻 Kedar Kumar Trivedi
📚 Electronics & Communication | 4th Semester
🏫 GTU College
