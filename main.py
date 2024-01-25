import pandas as pd
import numpy as np
import os

# Change Directory to where the raw CSVs are stored
directory_path = r'H:\Shared drives\97 - Finance Only\02 - COGS Pricing & Forecasting\00 - The Pool\The Pool CSV'
os.chdir(directory_path)

# Load Files
invoice_data = pd.read_csv('Invoice Data.csv', sep=';', encoding='utf-8')
credit_notes_data = pd.read_csv('Credit Notes.csv', sep=';', encoding='utf-8')
price_change_data = pd.read_csv('Price Change.csv', sep=';', encoding='utf-8')
quantity_aligner = pd.read_csv('The Pool Quantity Converter.csv', sep=',', encoding='utf-8')

# Rename Columns Names
def column_rename(df):
    df = df.rename(columns={
        'HKd.-Nr.': 'Rx ID',
        'HKd.-Kurzbez. LA oder RA': 'Rx Name',
        'Auftragseingangs-Datum': 'Order Date',
        'Kd.-Bestell-Nr.': 'Order ID',
        'Liefer-Datum': 'Delivered Date',
        'Rechnungs-Nr.': 'Invoice ID',
        'Artikel-Nr.': 'Item ID',
        'VArt.-Bez.': 'Item Name',
        'VK-Preis': 'Unit Price',
        'PB-Einheit': 'Packaging Unit',
        'Menge': 'Quantity',
        'BE': 'Packing Unit',
        'abweichende Bestellmenge': 'Deviating Quantity',
        'BE.1': 'GS Packing Unit',
        'GS-Grund-Nr.': 'GS Number',
        'GS-Grund-Bez.': 'GS Reason'})
    return df

try:
    invoice_data = column_rename(invoice_data)
    credit_notes_data = column_rename(credit_notes_data)
except KeyError:
    pass

def convert_value(value):
    if isinstance(value, str):  # Check if the value is a string
        try:
            # Attempt to convert value to a float
            return float(value.replace(' €', '').replace(',', '.'))
        except ValueError:
            # If conversion fails, check if it's 'Wochenpreis' and return 'Weekly Price'
            if value.strip().lower() == 'wochenpreis':
                return 'Weekly Price'
            else:
                return value

# Convert To Float Prices
invoice_data['Unit Price'] = invoice_data['Unit Price'].apply(convert_value)
credit_notes_data['Unit Price'] = credit_notes_data['Unit Price'].apply(convert_value)
credit_notes_data['Unit Price'] = credit_notes_data['Unit Price'].apply(convert_value)

price_change_list = price_change_data.iloc[:, 3:]
price_change_list = price_change_list.columns.to_list()

for column in price_change_list:
    if column in price_change_data.columns:
        price_change_data[column] = price_change_data[column].apply(convert_value)

# Convert Dates
def convert_date(column):
    return pd.to_datetime(column, format='%d.%m.%y')  # Adjusted format to match "dd.mm.yy"

invoice_data['Order Date'] = invoice_data['Order Date'].apply(convert_date)
invoice_data['Delivered Date'] = invoice_data['Delivered Date'].apply(convert_date)
credit_notes_data['Order Date'] = credit_notes_data['Order Date'].apply(convert_date)
credit_notes_data['Delivered Date'] = credit_notes_data['Delivered Date'].apply(convert_date)

# Create Consolidated Output
output_data = invoice_data

# Align invoice_data and quantity_aligner
output_data = pd.merge(output_data, quantity_aligner[['Item ID', 'Price Conversion', 'Quantity Conversion', 'DE Name', 'EN Name', 'Category', 'VAT Rate']], on='Item ID', how='left')

def calculate_quantity(quantity, conversion):
    try:
        return quantity / conversion
    except TypeError:
        return quantity

def calculate_net_price(unit_price, conversion):
    try:
        return unit_price * conversion
    except TypeError:
        return unit_price

def calculate_gross_price(unit_price, vat_rate):
    try:
        return unit_price * (1 + vat_rate)
    except TypeError:
        return unit_price

def calculate_total_price(unit_price, quantity):
    try:
        return unit_price * quantity
    except TypeError:
        return unit_price

# Apply the function to calculate the new 'Unit Price'
output_data['Quantity'] = output_data.apply(lambda row: calculate_quantity(row['Quantity'], row['Quantity Conversion']), axis=1)

# Apply the function to calculate the new 'Unit Price'
output_data['Unit Net Price'] = output_data.apply(lambda row: calculate_net_price(row['Unit Price'], row['Price Conversion']), axis=1)
output_data['Unit Gross Price'] = output_data.apply(lambda row: calculate_gross_price(row['Unit Net Price'], row['VAT Rate']), axis=1)

# # Apply the function for 'Total Price', but now multiply with 'Quantity'
output_data['Total Net Price'] = output_data.apply(lambda row: calculate_total_price(row['Unit Net Price'], row['Quantity']), axis=1)
output_data['Total Gross Price'] = output_data.apply(lambda row: calculate_total_price(row['Unit Gross Price'], row['Quantity']), axis=1)

# Re-Order Master Output
output_data = output_data[['Invoice ID', 'Order ID', 'Order Date', 'Delivered Date', 'Item ID', 'Item Name', 'Category',
                           'Quantity', 'Unit Net Price', 'Unit Gross Price', 'Total Net Price', 'Total Gross Price']]

# Creat Excel File Name
excel_file_path = 'consolidated_data.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
    # Write each DataFrame to a different worksheet
    output_data.to_excel(writer, sheet_name='Consolidated Output', index=False)
    invoice_data.to_excel(writer, sheet_name='Invoice Data', index=False)
    credit_notes_data.to_excel(writer, sheet_name='Credit Notes', index=False)
    price_change_data.to_excel(writer, sheet_name='Price Change', index=False)
    quantity_aligner.to_excel(writer, sheet_name='Quantity Aligner', index=False)

    # Access the XlsxWriter workbook and worksheet objects from the dataframe
    workbook = writer.book
    consolidated_output_worksheet = writer.sheets['Consolidated Output']

    # Create a date format
    date_format = workbook.add_format({'num_format': 'dd-mmm-yy'})

    # Set the width of columns in the 'Consolidated Output' worksheet
    consolidated_output_worksheet.set_column('A:A', 10)  # Set width of 'Invoice ID' column to 10
    consolidated_output_worksheet.set_column('B:B', 20)  # Example for another column
    consolidated_output_worksheet.set_column('C:C', 12)  # Example for a range of columns
    consolidated_output_worksheet.set_column('C:C', date_format)  # Example for a range of columns
    consolidated_output_worksheet.set_column('D:D', 12)  # Example for a range of columns
    consolidated_output_worksheet.set_column('D:D', date_format)  # Example for a range of columns
    consolidated_output_worksheet.set_column('E:E', 12)  # Example for a range of columns
    consolidated_output_worksheet.set_column('F:F', 60)  # Example for another column
    consolidated_output_worksheet.set_column('H:L', 12)  # Example for a range of columns

    # You can add similar lines to set the column widths for other sheets if needed
