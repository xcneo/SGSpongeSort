import os
import pandas as pd
import shutil

def process_tiff_files():
    source_dir = 'spongephotos'
    output_dir = 'output_repository'
    metadata_file = 'metadata_sponge.xlsx'

    # Read metadata
    df = pd.read_excel(metadata_file, sheet_name="Sorted Slides List")
    df["SpeciesID"] = df["Parent ZRC"].apply(extract_zrc_codes)
    df = df.explode("SpeciesID")
    df.dropna(subset=["SpeciesID"], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Find all .TIF files
    for filename in os.listdir(source_dir):
        if filename.upper().endswith('.TIF'):
            accession_number = filename.split('_')[0].replace("ZRCPOR", "ZRC")
            
            # Find the order for the given accession number
            order_series = df[df['SpeciesID'] == accession_number]['Order']
            
            if not order_series.empty:
                order = str(order_series.iloc[0]).strip()
                
                # Create destination directory
                dest_dir = os.path.join(output_dir, order)
                os.makedirs(dest_dir, exist_ok=True)
                
                # Move the file
                source_path = os.path.join(source_dir, filename)
                dest_path = os.path.join(dest_dir, filename)
                shutil.move(source_path, dest_path)
                print(f"Moved {filename} to {dest_dir}")
            else:
                print(f"No order found for {accession_number}")

def extract_zrc_codes(text):
    if not isinstance(text, str):
        return []
    # Normalize variations to a standard ZRC prefix
    text = text.replace("ZRC.POR.", "ZRC")
    text = text.replace("ZRCPOR", "ZRC")
    # Find all occurrences of ZRC followed by digits
    import re
    codes = re.findall(r'ZRC\d+', text)
    return codes

if __name__ == "__main__":
    process_tiff_files()
