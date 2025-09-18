# %%
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta


# Convert AN_Store from decimal to integer string, preserving nulls
def clean_store(val):
    if pd.isna(val):
        return np.nan
    try:
        return str(int(float(val)))
    except:
        return val  # If it can't convert, keep original

def Update_Historic_MS():

    #---------------------------------Load files to df--------------------------------------------------

    # File paths
    file_path = r"W:\Corporate\Inventory\Urban Science\AutoNation_SalesFile_NationalSales.txt"
    file_path2 = r"C:\Development\R\Market_Share\Historical\AutoNation_SalesFile_NationalSales_Historical_202201_202506.txt"

    # Load files

    df_curr = pd.read_csv(file_path, delimiter=',', quotechar='"', encoding='utf-8') 
    df_hist = pd.read_csv(file_path2, delimiter=',', quotechar='"', encoding='utf-8')

    print("Files loaded")

    #---------------------------------Transformation--------------------------------------------------
    print("Transformation ongoing")
    df_filtered = df_hist[(df_hist['SALE_MONTH'].isin([5])) & (df_hist['SALE_YEAR'].isin([2025, 2024]))]


    df_combined = pd.concat([df_curr, df_filtered], ignore_index=True)

    df_combined['AN_Store'] = df_combined['AN_Store'].apply(clean_store)

    print(df_combined)

    #---------------------------------Save output file--------------------------------------------------
    print("Saving file. Please wait")
    output_path = fr"W:\Corporate\Inventory\Urban Science\AutoNation_SalesFile_NationalSales.txt"

    df_combined.to_csv(output_path, sep=',', index=False, encoding='utf-8-sig')

    print(f"File savedat: {output_path}")
    
    return

#run function
if __name__ == '__main__':
    
    Update_Historic_MS()