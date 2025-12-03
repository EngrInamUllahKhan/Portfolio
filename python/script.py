import os
import pandas as pd
import zipfile
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

def process_local(input_file_path, output_folder):
    logging.info(f"Running JBCompetitor_Matrix process on: {input_file_path}")

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)
    
    # Extract input file name
    input_blob = os.path.basename(input_file_path)
    date_string = input_blob[:8]

    # Set output file names
    output_psv = input_blob.replace('.xlsx','.psv').replace(' - ', '_').replace(' ','_')
    output_zip = input_blob.replace('.xlsx','.zip').replace(' - ', '_').replace(' ','_')
    output_psv_path = os.path.join(output_folder, output_psv)
    output_zip_path = os.path.join(output_folder, output_zip)

    # Load Excel sheets
    data_matrix = pd.read_excel(input_file_path, sheet_name=0, skiprows=2)
    data_priority = pd.read_excel(input_file_path, sheet_name=1, skiprows=2)

    # Column names and department IDs
    column_names = ['Date','Department_ID','Department', 'Retailer', 'Competitor', 'Competitor Priority']
    department_ids = {
        "ACCESSORIES":62,
        "ACCESSORIES HW":60,
        "AIR CONDITIONING":46,
        "AUDIO":10,
        "CAMERAS":70,
        "COMMUNICATIONS":80,
        "COMPUTERS":75,
        "SMART HOME": 20,
        "COOKING":42,
        "FITNESS":35,
        "GAMES HW":90,
        "GAMES SW":92,
        "IN CAR": 50,
        "IT":65,
        "MISCELLANEOUS":120,
        "MOVIES":110,
        "MUSIC":100,
        "MUSICAL INSTRUMENTS":15,
        "SMALL APPLIANCES":55,
        "VISUAL":30,
        "WHITEGOODS":85
    }

    n_rows = len(data_matrix)
    n_cols = len(data_matrix.columns)
    final_data = []

    # Transform data
    for col in range(2, n_cols):
        for row in range(0, n_rows):
            temp_item = [
                date_string,
                department_ids.get(data_matrix.iloc[row,1]),
                data_matrix.iloc[row,1],
                data_matrix.columns[col],
                data_matrix.iloc[row,col],
                data_priority.iloc[row,col]
            ]
            final_data.append(temp_item)

    # Create DataFrame and save PSV
    final_df = pd.DataFrame(final_data, columns = column_names)
    final_df.to_csv(output_psv_path, sep='|', index=False, encoding='utf-8', line_terminator='\r\n')
    logging.info(f"PSV file created: {output_psv_path}")

    # Create ZIP containing PSV
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipped_file:
        zipped_file.write(output_psv_path, arcname=output_psv)
    logging.info(f"ZIP file created: {output_zip_path}")


# -----------------------------
# Example usage with your paths
# -----------------------------
if __name__ == "__main__":
    input_file_path = r"C:\Users\InamUllah\Downloads\input\20251203 - Competitor Matrix.xlsx"
    output_folder = r"C:\Users\InamUllah\Downloads\output"

    process_local(input_file_path, output_folder)
