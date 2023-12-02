import os
import pandas as pd
import requests

# Constants
EXCEL_URL = 'https://bvp-cloud-index.s3.amazonaws.com/BVP-Nasdaq-Emerging-Cloud-Index.xlsx'
TEMP_EXCEL_FILE = 'temp_excel_file.xlsx'
OUTPUT_CSV_FILE = 'emc.csv'
FINAL_OUTPUT_FILE = 'emcloud.csv'

# Function to download and save the Excel file
def download_file(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as file:
            file.write(response.content)
    else:
        raise Exception(f"Failed to download file: {url}")

# Function to process the Excel file and save it as a CSV
def process_excel_to_csv(input_file, output_csv):
    excel_file = pd.ExcelFile(input_file)
    second_sheet_name = excel_file.sheet_names[1]
    second_sheet_df = pd.read_excel(input_file, sheet_name=second_sheet_name, skiprows=7)
    second_sheet_df.drop(second_sheet_df.index[[0, 1]], inplace=True)
    second_sheet_df = second_sheet_df.iloc[:, 1:]
    second_sheet_df.to_csv(output_csv, index=False)

# Function to calculate and assign rankings
def calculate_rankings(csv_file_path):
    cloud_data = pd.read_csv(csv_file_path)

    weights = {
        'Market Cap': 0.20, 'EV / Annualized Revenue': 0.10, 'EV / Forward Revenue': 0.10,
        'Efficiency': 0.10, 'Revenue Growth Rate': 0.15, 'Gross Margin': 0.10,
        'LTM FCF Margin': 0.15, 'SGA Margin': 0.025, 'R&D Revenue': 0.025,
        'Sales Marketing Revenue': 0.05
    }

    cloud_data['Score'] = (
        cloud_data['EV / Annualized Revenue'].rank(ascending=True) * weights['EV / Annualized Revenue']
        + cloud_data['EV / Forward Revenue'].rank(ascending=True) * weights['EV / Forward Revenue']
        + cloud_data['SGA Margin'].rank(ascending=True) * weights['SGA Margin']
        + cloud_data['Market Cap'].rank(ascending=False) * weights['Market Cap']
        + cloud_data['Efficiency'].rank(ascending=False) * weights['Efficiency']
        + cloud_data['Revenue Growth Rate'].rank(ascending=False) * weights['Revenue Growth Rate']
        + cloud_data['Gross Margin'].rank(ascending=False) * weights['Gross Margin']
        + cloud_data['LTM FCF Margin'].rank(ascending=False) * weights['LTM FCF Margin']
        + cloud_data['R&D Revenue'].rank(ascending=False) * weights['R&D Revenue']
        + cloud_data['Sales Marketing Revenue'].rank(ascending=False) * weights['Sales Marketing Revenue']
    ).round(2)

    cloud_data.sort_values(by='Score', ascending=True, inplace=True)
    cloud_data['Rank'] = cloud_data['Score'].rank(method='first', ascending=True).astype(int)
    cloud_data.to_csv(FINAL_OUTPUT_FILE, index=False)

# Main execution
if __name__ == "__main__":
    download_file(EXCEL_URL, TEMP_EXCEL_FILE)
    process_excel_to_csv(TEMP_EXCEL_FILE, OUTPUT_CSV_FILE)
    calculate_rankings(OUTPUT_CSV_FILE)
    os.remove(TEMP_EXCEL_FILE)
    os.remove(OUTPUT_CSV_FILE)
