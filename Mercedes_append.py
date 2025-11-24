import xlwings as xw
import pandas as pd
import os
import glob

def append_Mercedes():

    path = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Mercedes_appends'

    file_patterns = ['*.csv']

    # List to store all dataframes
    dataframes = []
    all_files = []

    for pattern in file_patterns:
        all_files.extend(glob.glob(os.path.join(path, pattern)))


    
    for file_path in all_files:
        try:
            # Extract filename without extension
            filename = os.path.basename(file_path)
            filename_no_ext = os.path.splitext(filename)[0]
            
            # Extract store name (everything before the first underscore)
            if '_' in filename_no_ext:
                store_name = filename_no_ext.split('_')[1]
            else:
                store_name = filename_no_ext  # Use full filename if no underscore
            
            # Read the file based on extension
            if file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                continue
            
            # Add store column
            df['Store'] = store_name
            
            # Add the dataframe to our list
            dataframes.append(df)
            print(f"Processed: {filename} -> Store: {store_name}")
            
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
            continue
    
    if not dataframes:
        print("No valid files were processed.")
        return None
    
    # Combine all dataframes
    combined_df = pd.concat(dataframes, ignore_index=True)
    
    # Save the combined file
    Final_path = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Mercedes_appends\Final'
    output_path = os.path.join(Final_path, 'Mercedes_Combined.xlsx')
    combined_df.to_excel(output_path, index=False)
    
    print(f"Successfully combined {len(dataframes)} files into: {output_path}")
    print(f"Total rows: {len(combined_df)}")
    
    return combined_df

if __name__ == "__main__":
    append_Mercedes()