# Excel Transaction Combiner

A Python script to combine multiple Coop Mastercard transaction Excel files into a single file. The script handles both `.xls` and `.xlsx` files, translates Swedish column names to English, and formats dates consistently.

I use this script to combine the transactions from my Coop Mastercard statements into a single file that I can then import into my accounting software (Buxfer)

The script will:

- Read all Excel files in the current directory
- Combine them into `Combined_Transactions.xlsx`
- Skip the source file, date, and purchase date columns
- Output without headers

## Column Translations

- Datum → Date
- Valuta → Currency
- Plats → Location
- Kortinnehavare → Cardholder
- Fakturabelopp → Invoice Amount
- Detaljer → Details
- Transaktionsbelopp → Transaction Amount

## Features

- Combines multiple Excel files into one
- Translates Swedish column names to English
- Formats dates as YYYY/MM/DD
- Handles both .xls and .xlsx files
- Converts amount columns to numeric values
- Skips temporary Excel files (starting with ~$)

## Requirements

- Python 3.x
- pandas
- openpyxl
- xlrd

## Installation

Install dependencies using pip and the requirements.txt file:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python combine_excel.py
```
