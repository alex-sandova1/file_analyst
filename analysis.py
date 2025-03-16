import pandas as pd
from openpyxl import load_workbook

def read_tables_from_sheet(sheet):
    tables = []
    current_table = []
    row_count = 0  # Counter to track the number of rows processed
    for row in sheet.iter_rows(values_only=True):
        row_count += 1
        #print(f"Processing row {row_count}: {row}")  # Debugging statement
        # Check if the row contains any non-None values
        if any(cell is not None for cell in row):
            current_table.append(row)
        else:
            if current_table:
                tables.append(current_table)
                current_table = []
    if current_table:
        tables.append(current_table)
    #print(f"Total rows processed: {row_count}")  # Debugging statement
    return tables

def remove_none_values(table):
    # Remove cells that are None 
    # and remove rows that are completely empty
    return [[cell for cell in row if cell is not None] for row in table if any(cell is not None for cell in row)] if table else table

def remove_total_row(table):
    # Remove the last row if it contains the word "Total" in any cell, skipping the first row
    return [row for i, row in enumerate(table) if i == 0 or not any('total' in str(cell).lower() for cell in row)] if table else table

def Highest_Payment(table):
    # Find column that contains the word "Payment" or "Monthly"
    payment_col_index = None
    for i, header in enumerate(table[0]):
        normalized_header = header.strip().lower() if header else ''
        if 'payment' in normalized_header or 'monthly' in normalized_header:
            payment_col_index = i
            break
    if payment_col_index is None:
        print("No payment column found.")
        return
    # Print payment column index for debugging
    #print(f"Payment column index: {payment_col_index}")
    
    # Verify that the payment column index is valid
    if payment_col_index >= len(table[0]):
        print("Invalid payment column index.")
        return
    
    # Store the column index with their respective row
    payment_values = []
    for row in table[1:]:
        if row[payment_col_index] is not None:
            payment_values.append((payment_col_index, row[payment_col_index]))
        else:
            print(f"Found None value in payment column at row: {row}")
    
    # Find the highest payment value
    if payment_values:
        highest_payment = max(payment_values, key=lambda x: x[1])
        print(f"Highest payment value: {highest_payment[1]}")
    else:
        print("No valid payment values found.")
    #return value of highest payment
    return highest_payment[1] if payment_values else None

def Lowest_Payment(table):
    
    # Find column that contains the word "Payment" or "Monthly"
    payment_col_index = None
    for i, header in enumerate(table[0]):
        normalized_header = header.strip().lower() if header else ''
        if 'payment' in normalized_header or 'monthly' in normalized_header:
            payment_col_index = i
            break
    if payment_col_index is None:
        print("No payment column found.")
        return 
    # Print payment column index for debugging
    #print(f"Payment column index: {payment_col_index}")
    
    # Verify that the payment column index is valid
    if payment_col_index >= len(table[0]):
        print("Invalid payment column index.")
        return
    
    # Store the column index with their respective row
    payment_values = []
    for row in table[1:]:
        if row[payment_col_index] is not None:
            payment_values.append((payment_col_index, row[payment_col_index]))
        else:
            print(f"Found None value in payment column at row: {row}")
    
    # Find the lowest payment value
    if payment_values:
        lowest_payment = min(payment_values, key=lambda x: x[1])
        print(f"Lowest payment value: {lowest_payment[1]}")
    else:
        print("No valid payment values found.")
   
    #return value of lowest payment
    return lowest_payment[1] if payment_values else None