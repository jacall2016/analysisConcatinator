import os
import pandas as pd
from tkinter import Tk, filedialog

def get_output_path():
    # Create a Tkinter root window, but don't display it
    root = Tk()
    root.withdraw()

    # Ask the user to select a directory for the output file
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    return output_path

def get_folder_path():
    # Create a Tkinter root window, but don't display it
    root = Tk()
    root.withdraw()

    # Ask the user to select a directory for the input Excel files
    folder_path = filedialog.askdirectory()
    
    return folder_path

def extract_plate_number(file_name):
    if "analysis_" in file_name:
        # Find the index of the first capital L
        start_index = file_name.find("L")
        # Find the index of the character one past the first capital P after the capital L
        end_index = file_name.find("P", start_index) + 2
        # Extract the plate number
        if start_index != -1 and end_index != -1:
            plate_number = file_name[start_index:end_index]
            return plate_number
    # Return default plate number if there's an error
    return "ERROR_GETTING_FILE_NAME"

def IsHighControl(well_number):
    if len(well_number) < 3:
        return False
    if well_number[1] not in ['1', '2'] or well_number[2] != '.':
        return False
    return True

def process_excel_folder(folder_path):
    all_dfs = []
    
    # Loop through all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(folder_path, filename)
            # Read Excel file
            xls = pd.ExcelFile(filepath)
            
            # Check if 'Analysis' sheet exists
            if 'Analysis' in xls.sheet_names:
                # Read 'Analysis' sheet into DataFrame
                df = pd.read_excel(xls, 'Analysis')
                
                # Add 'high_controls' and 'plate_number' columns
                df['high_controls'] = None
                df['plate_number'] = extract_plate_number(filename)
                
                # Append DataFrame to the list
                all_dfs.append(df)
    
    # Combine all DataFrames into one
    combined_df = pd.concat(all_dfs, ignore_index=True)

    combined_df = combined_df[['plate_number', 'well_number', 
                               'yemk_z_score', 'hits_yemk_z_score', 
                               'high_controls', 
                               'phl_z_score', 'hits_phl_z_score', 
                               'flip700_z_score', 'hits_flip700_z_score',
                               'live_z_score', 'hits_live_z_score']]
    
    return combined_df

def separate_dataframes(combined_df):
    # Create separate DataFrames for each type
    yemk_df = combined_df[['plate_number', 'well_number', 'yemk_z_score', 'hits_yemk_z_score', 'high_controls']]
    phl_df = combined_df[['plate_number', 'well_number', 'phl_z_score', 'hits_phl_z_score', 'high_controls']]
    flip700_df = combined_df[['plate_number', 'well_number', 'flip700_z_score', 'hits_flip700_z_score', 'high_controls']]
    live_df = combined_df[['plate_number', 'well_number', 'live_z_score', 'hits_live_z_score', 'high_controls']]
    
    return yemk_df, phl_df, flip700_df, live_df

def populate_high_controls(df):
    z_score_column = [col for col in df.columns if 'z_score' in col and 'hits' not in col][0]
    # Iterate over rows and populate 'high_controls' column
    for index, row in df.iterrows():
        if IsHighControl(row['well_number']):
            # Populate high_controls with the corresponding z_score column
            df.at[index, 'high_controls'] = row[z_score_column]
    return df

def concatenate_columns(folder_path, output_path, sheet_columns):
    combined_df = process_excel_folder(folder_path)

    yemk_df, phl_df, flip700_df, live_df = separate_dataframes(combined_df)

    yemk_df = populate_high_controls(yemk_df)
    phl_df = populate_high_controls(phl_df)
    flip700_df = populate_high_controls(flip700_df)
    live_df = populate_high_controls(live_df)

    # Write each DataFrame to separate sheets in the Excel file
    with pd.ExcelWriter(output_path) as writer:
        yemk_df.to_excel(writer, index=False, sheet_name='yemk')
        phl_df.to_excel(writer, index=False, sheet_name='phl')
        flip700_df.to_excel(writer, index=False, sheet_name='flip700')
        live_df.to_excel(writer, index=False, sheet_name='live')

if __name__ == "__main__":
    # Example usage
    folder_path = get_folder_path()  # Prompt the user to choose the folder containing Excel files
    output_path = get_output_path()  # Empty to prompt the user for the output location

    # Specify the sheets and columns for each sheet
    sheet_columns = {
        'yemk': ['plate_number', 'well_number', 'yemk_z_score', 'hits_yemk_z_score', 'high_controls'],
        'phl': ['plate_number', 'well_number', 'phl_z_score', 'hits_phl_z_score','high_controls'],
        'flip700': ['plate_number', 'well_number', 'flip700_z_score', 'hits_flip700_z_score', 'high_controls'],
        'live': ['plate_number', 'well_number', 'live_z_score', 'hits_live_z_score', 'high_controls']
    }
    concatenate_columns(folder_path, output_path, sheet_columns)