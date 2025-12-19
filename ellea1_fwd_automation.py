import pandas as pd
import xlwings as xw
import glob, os

# --- CONFIGURATION ---
EXCEL_CALCULATOR_FILE = r'C:\Path\To\Your\Ellea1_FWD_Excel_Worksheet.xlsx'

###################################################################
# INPUT_FILES = r"C:\Path\To\Inputs\Your_Inputs.csv"
###################################################################
# In case of specific files, uncomment abvoe and comment below#####
###################################################################
INPUT_FILES = glob.glob(r"C:\Path\To\Inputs\*.csv")
###################################################################

OUTPUT_DIRECTORY = r"C:\Path\To\Outputs"
SHEET_NAME = 'Main'
BATCH_SIZE = 5000  # <<< SAVE EVERY 5000 ROWS

INPUT_CELL_MAP = {
    'Thickness layer 1 (mm)': 'D2',
    'Thickness layer 2 (mm)': 'D3',
    'Thickness layer 3 (mm)': 'D4',
    'Stiffness layer 1 (MPa)': 'B2',
    'Stiffness layer 2 (MPa)': 'B3',
    'Stiffness layer 3 (MPa)': 'B4',
    'Stiffness layer 4 (MPa)': 'B5',
    'Poisson ratio layer 1': 'C2',
}

OUTPUT_CELL_MAP = {
    'd1(nm)': 'C13', 'd2(nm)': 'C14', 'd3(nm)': 'C15', 'd4(nm)': 'C16',
    'd5(nm)': 'C17', 'd6(nm)': 'C18', 'd7(nm)': 'C19', 'd8(nm)': 'C20',
    'd9(nm)': 'C21', 'd10(nm)': 'C22',
}

def save_chunk(data, output_path, columns, is_first_save):
    """Helper to save a chunk of data to CSV"""
    df_chunk = pd.DataFrame(data)
    # Ensure columns match the desired order
    df_chunk = df_chunk.reindex(columns=columns)
    
    # If first save, 'w' (write) + Header. If later save, 'a' (append) + No Header.
    mode = 'w' if is_first_save else 'a'
    header = True if is_first_save else False
    
    df_chunk.to_csv(output_path, mode=mode, header=header, index=False)
    print(f"    -> Checkpoint: Saved {len(data)} rows to disk.")

def process_file(input_path, output_path, excel_wb):
    print(f"  > Processing: {os.path.basename(input_path)}")
    
    df_inputs = pd.read_csv(input_path, sep=',')
    sheet = excel_wb.sheets[SHEET_NAME]
    
    current_batch = []
    total_rows = len(df_inputs)
    file_initialized = False # Tracks if we have created the file yet
    
    # Define Column Order ahead of time
    final_columns = list(df_inputs.columns) + list(OUTPUT_CELL_MAP.keys())

    for index, row in df_inputs.iterrows():
        # Notification
        if (index + 1) % 1000 == 0:
            print(f"    ... row {index + 1}/{total_rows}")

        # Input
        for csv_col, excel_cell in INPUT_CELL_MAP.items():
            sheet.range(excel_cell).value = row[csv_col]
        
        # Output
        current_result = {}
        for output_name, excel_cell in OUTPUT_CELL_MAP.items():
            current_result[output_name] = sheet.range(excel_cell).value
        
        # Add context
        for col_name in df_inputs.columns:
            current_result[col_name] = row[col_name]

        current_batch.append(current_result)

        # --- CHECKPOINT: SAVE EVERY 5000 ROWS ---
        if len(current_batch) >= BATCH_SIZE:
            save_chunk(current_batch, output_path, final_columns, not file_initialized)
            file_initialized = True
            current_batch = [] # Clear memory

    # --- SAVE REMAINING ROWS ---
    if current_batch:
        save_chunk(current_batch, output_path, final_columns, not file_initialized)
    
    print(f"  > DONE: Finished {input_path}")

def main():
    os.makedirs(OUTPUT_DIRECTORY, exist_ok=True)
    print("--- Starting Automation (Safe Mode: Auto-Saving) ---")

    with xw.App(visible=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        
        workbook = app.books.open(EXCEL_CALCULATOR_FILE)
        
        for input_file in INPUT_FILES:
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_path = os.path.join(OUTPUT_DIRECTORY, f"{base_name}_output.csv")
            process_file(input_file, output_path, workbook)
            
        workbook.close()
    
    print("--- Finished Successfully ---")

if __name__ == '__main__':

    main()
