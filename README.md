# pdf-to-excel-convert
<img width="916" height="639" alt="image" src="https://github.com/user-attachments/assets/4e6c4c00-0760-4951-9d6b-1dd93dcd64ca" />
A simple Python application that converts PDF tables into well-structured Excel files.

## Features

- Convert tables from PDF to Excel (.xlsx)
- Automatically detects table headers
- Extracts metadata above each table
- Groups similar tables into the same Excel worksheet
- Removes empty and total rows
- Automatically adjusts column widths
- Modern graphical interface built with CustomTkinter
- Progress bar during conversion

## Technologies

- Python
- CustomTkinter
- pdfplumber
- openpyxl

## Installation

Clone the repository:

```bash
git clone https://github.com/your-username/pdf-to-excel-converter.git
cd pdf-to-excel-converter
```

Install the required packages:

```bash
pip install -r requirements.txt
```

## Usage

Run the application:

```bash
python main.py
```

1. Select a PDF file.
2. Choose where to save the Excel file.
3. Click **Convert**.
4. Wait for the conversion to finish.

## Project Structure

```
.
├── main.py
├── requirements.txt
└── README.md
```

## Output

The generated Excel file contains:

- Extracted metadata
- Cleaned tables
- Automatic formatting
- One worksheet per detected group
