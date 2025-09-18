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
    file_path = r"W:\Corporate\Inventory\Urban Science\PowerBI\V2\AutoNation_SalesFile_NationalSales_Historical.txt"

    # Load files
    df_curr = pd.read_csv(file_path, delimiter=',', quotechar='"', encoding='utf-8')
    # Remove store 2071
    df_filtered = df_curr[(df_curr['AN_Store'] != 2071.0)] 

    #---------------------------------Save output file--------------------------------------------------
    output_path = fr"W:\Corporate\Inventory\Urban Science\PowerBI\V2\AutoNation_SalesFile_NationalSales_Historical.txt"
    df_filtered.to_csv(output_path, sep=',', index=False, encoding='utf-8-sig')

    print(f"File savedat: {output_path}")
    
    return

#run function
if __name__ == '__main__':
    
    Update_Historic_MS()