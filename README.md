# AI-Powered Excel Business Analyst

An AI-powered business analysis application that transforms sales data from Excel and CSV files into structured business insights and downloadable Excel reports.

The application uses Python and Pandas for data processing and calculations, Streamlit for the user interface, and Ollama with a local LLM for AI-powered business explanations.

## Features

- **Excel and CSV Upload:** Import sales datasets in XLSX or CSV formats.
- **Data Quality Checks:** Identify duplicate records, missing values, and invalid numeric values.
- **Automated Sales Analysis:** Calculate completed revenue, completed orders, and average order value.
- **Product Analysis:** Analyze product revenue and sales quantities.
- **Regional Analysis:** Compare revenue across different regions.
- **Order Status Analysis:** Review completed, pending, and cancelled orders.
- **AI Business Analyst:** Ask business questions in natural language and receive explanations based on verified analytical results.
- **Excel Report Generation:** Export a formatted Excel report containing six analytical sheets.

## Technology Stack

- Python
- Pandas
- Streamlit
- Ollama
- Qwen 2.5 7B
- OpenPyXL

## Project Structure

```text
AI-Powered-Excel-Business-Analyst/
│
├── app.py
├── AI_Business_Analysis_Report.xlsx
├── requirements.txt
├── README.md
│
└── data/
    ├── sample_sales.xlsx
    └── test_customers.xlsx
```

## How It Works

1. Upload a sales dataset in Excel or CSV format.
2. The application validates the dataset and checks data quality.
3. Pandas processes the data and calculates business metrics.
4. Explore sales performance through analytical tables and charts.
5. Ask business questions using the AI Business Analyst.
6. Download the generated Excel report.

## Excel Report

The generated report contains six worksheets:

1. Summary
2. Data Quality
3. Product Analysis
4. Region Analysis
5. Order Status
6. Prepared Data

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/milad-mardani/AI-Powered-Excel-Business-Analyst.git
```

### 2. Navigate to the Project Directory

```bash
cd AI-Powered-Excel-Business-Analyst
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Install and Configure Ollama

Install Ollama from:

https://ollama.com/

Download the required model:

```bash
ollama pull qwen2.5:7b
```

### 5. Run the Application

```bash
streamlit run app.py
```

The application will open in your browser.

## Sample Datasets

The `data/` directory contains:

- **sample_sales.xlsx:** Sample sales dataset compatible with the current application's sales analysis features.
- **test_customers.xlsx:** An additional customer dataset included for testing. Customer-data analysis is not currently supported by the application.

## Important Notes

- The current application is designed for sales data.
- Required sales columns include OrderID, Product, Region, Quantity, UnitPrice, and Status.
- Revenue is calculated using Quantity multiplied by UnitPrice.
- Numerical analysis is performed using Python and Pandas.
- The language model explains analytical results; it does not independently calculate business metrics.
- Ollama is used for local AI inference.

## Future Improvements

- Support for additional business datasets.
- Dynamic data profiling for different data structures.
- More advanced business analytics.
- Additional AI-powered analytical capabilities.

## Author

**Milad Mardani**

GitHub: https://github.com/milad-mardani
