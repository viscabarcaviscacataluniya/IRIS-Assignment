# 📊 FastAPI Excel Document Parser

This project is a lightweight and efficient FastAPI-based web service to parse and analyze Excel files. It exposes simple APIs to:

- 📋 List all sheet names (tables) in the Excel file  
- 🔍 Fetch row names from a given sheet  
- ➕ Calculate the sum of numeric values in a specified row

## 🚀 Features

- FastAPI backend
- Supports `.xls` files
- Error handling for missing sheets or rows
- Easy to extend for more Excel operations

## 📁 Project Structure

IRIS Assignment/
├── main.py # FastAPI app and endpoints
├── utils/
│ └── excel_utils.py # Utility functions for Excel handling
├── capbudg.xls # Sample Excel file
└── README.md # Project documentation
## 🔧 Setup & Run

1. **Install dependencies**  
```bash
pip install fastapi uvicorn pandas openpyxl xlrd
Start the server

bash
Copy
Edit
uvicorn main:app --reload --port 9090
API Endpoints
GET /list_tables
Returns all sheet names in the Excel file

GET /get_table_details?table_name=Sheet1
Returns row names (first column values) of the specified sheet

GET /row_sum?table_name=Sheet1&row_name=SomeRow
Returns the sum of all numeric values in the specified row
