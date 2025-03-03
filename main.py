import os
import openpyxl
from google.api_core.exceptions import ResourceExhausted
import google.generativeai as genai
import json
from collections import defaultdict
from preprocess import process_excel_and_extract_data
from ReviewAnalysis2 import process_reviews
from datetime import datetime, timedelta
import calendar

def get_previous_month(date):
    """Returns the previous month of the given date."""
    if isinstance(date, str):
        date = datetime.strptime(date, "%Y-%m-%d")

    first_of_current_month = date.replace(day=1)
    last_day_of_previous_month = first_of_current_month - timedelta(days=1)
    return last_day_of_previous_month.month  

def filter_rows_by_previous_month(file_path, previous_month):
    """Filters the rows of the Excel file based on the previous month."""
    workbook = openpyxl.load_workbook(file_path)
    filtered_rows = []
    
    for sheet in workbook.worksheets:
        month_column_index = None
        for cell in sheet[1]:
            if cell.value and cell.value.strip().lower() == 'month':
                month_column_index = cell.column
                break
        
        if month_column_index is None:
            print(f"Month column not found in sheet {sheet.title}. Skipping...")
            continue

        # Filter rows based on the 'Month' column value
        for row in sheet.iter_rows(min_row=2, values_only=True):
            month_value = row[month_column_index - 1]
            if month_value == previous_month:
                filtered_rows.append(row)

    return filtered_rows

def main():
    """Main function to execute the sentiment analysis."""
    api_key = "AIzaSyAxk2Wog2ylp7wuQgTGdQCakzJXMoRHzO8"
    genai.configure(api_key=api_key)

    current_date = datetime.now()  
    previous_month = get_previous_month(current_date)
    print(f"Previous month is: {previous_month}")

    file_path = "/Users/yash/Downloads/Today/Splitted/A2b January month.xlsx"

    filtered_rows = filter_rows_by_previous_month(file_path, previous_month)
    
    if not filtered_rows:
        print("No rows found for the previous month. Exiting...")
        return

    print(f"Found {len(filtered_rows)} rows for the previous month.")

    process_reviews(filtered_rows)

    updated_rows = filtered_rows

    output_file_path = "/Users/yash/Downloads/Today/Splitted/output_summary_competitor_analysis.xlsx"

    process_excel_and_extract_data(updated_rows, output_file_path, api_key)

if __name__ == "__main__":
    main()
