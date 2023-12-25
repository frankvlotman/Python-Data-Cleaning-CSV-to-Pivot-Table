import pandas as pd

def process_csv_and_create_pivot(csv_file_path, excel_file_path):
    # Read the CSV file, skip the first 3 rows, and trim leading spaces in the top row
    df = pd.read_csv(csv_file_path, skiprows=2)
    df = df.rename(columns=lambda x: x.strip())

    # Include only the specified columns & removal of unwanted columns
    columns_to_include = ['car_make', 'car_model', 'car_vin_no', 'color', 'quantity', 'amount']
    df_filtered = df[columns_to_include]

    # Create a pivot table grouped by 'car_make'
    pivot_table = df_filtered.pivot_table(index='car_make', aggfunc='sum')

    # Save the pivot table to an Excel file
    pivot_table.to_excel(excel_file_path)

# Source file & output file paths
csv_file_path = 'C:\\Users\\Frank\\Desktop\\dataset.csv'  # Path to CSV file
excel_file_path = 'C:\\Users\\Frank\\Desktop\\results.xlsx'  # Path for the output Excel file

process_csv_and_create_pivot(csv_file_path, excel_file_path)
