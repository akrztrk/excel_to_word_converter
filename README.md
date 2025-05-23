# Excel to Word Converter (Web App)

This project is a simple web application that allows users to upload Excel files and automatically convert each sheet into a separate Word document. It is built using Flask and follows an object-oriented design.

## 🚀 Features

- Upload Excel files (.xlsx, .xls)
- Convert each sheet into a separate Word (.docx) document
- Download converted Word files through the browser
- Simple and user-friendly web interface

## 🛠️ Technologies Used

- Python 3.x
- Flask (web framework)
- Pandas (for reading Excel files)
- python-docx (for writing Word documents)
- HTML with Jinja2 templating

## 📂 Project Structure

excel_to_word_web/
├── app.py # Main Flask application
├── converter/ # Conversion modules
│ ├── excel_reader.py
│ ├── word_writer.py
│ └── controller.py
├── templates/ # HTML templates
│ ├── index.html
│ └── downloads.html
├── uploads/ # Uploaded Excel files
├── outputs/ # Generated Word documents
└── README.md
