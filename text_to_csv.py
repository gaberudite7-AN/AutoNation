import os
import csv
import shutil

# Mode: 'copy' = copy/rename .txt -> .csv (keeps file contents unchanged)
#       'parse' = read lines, split by commas and write proper CSV rows
mode = 'copy'  # set to 'parse' if you want to split lines into CSV cells

# Define source and destination folders
source_folder = r'W:\Corporate\Inventory\Urban Science\Historics\Industry'
destination_folder = r'W:\Corporate\Inventory\Urban Science\Historics\Industry\CSV_Formatted'

# Create destination folder if it doesn't exist
os.makedirs(destination_folder, exist_ok=True)


# Iterate through all .txt files in the source folder
for filename in os.listdir(source_folder):
    if filename.endswith('.txt'):
        txt_path = os.path.join(source_folder, filename)
        csv_filename = os.path.splitext(filename)[0] + '.csv'
        csv_path = os.path.join(destination_folder, csv_filename)

        if mode == 'copy':
            # Copy the file and give it a .csv extension (same content as .txt)
            shutil.copy2(txt_path, csv_path)
        else:
            # Read the text file and write to CSV (split by commas)
            with open(txt_path, 'r', encoding='utf-8-sig') as txt_file:
                lines = txt_file.readlines()

            with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file)
                for line in lines:
                    writer.writerow([cell.strip() for cell in line.strip().split(',')])


print("All .txt files have been converted to .csv and copied to the destination folder.")