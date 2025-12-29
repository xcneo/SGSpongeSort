import os
import shutil
import pandas as pd
import re

def extract_zrc_codes(text):
    if not isinstance(text, str):
        return []
    # Normalize variations to a standard ZRC prefix
    text = text.replace("ZRC.POR.", "ZRC")
    text = text.replace("ZRCPOR", "ZRC")
    # Find all occurrences of ZRC followed by digits
    codes = re.findall(r'ZRC\d+', text)
    return codes

def organize_spicule_images(raw_img_dir, metadata_file, output_dir):
    # Load the master sheet with the mapping
    mapping_df = pd.read_excel(metadata_file, sheet_name="Sorted Slides List")

    # Apply the function to extract clean ZRC codes and explode the dataframe
    # to handle rows with multiple codes
    mapping_df["SpeciesID"] = mapping_df["Parent ZRC"].apply(extract_zrc_codes)
    mapping_df = mapping_df.explode("SpeciesID")

    # Drop rows where no valid SpeciesID could be extracted
    mapping_df.dropna(subset=["SpeciesID"], inplace=True)
    mapping_df.reset_index(drop=True, inplace=True)

    # Load all sheets into a dictionary
    all_sheets = pd.read_excel(metadata_file, sheet_name=None)

    # Remove the master mapping sheet from the dict and normalize sheet names
    species_sheets = {}
    for k, v in all_sheets.items():
        key_str = str(k).strip()
        # Handle the specific case of the mislabeled sheet
        if key_str == 'Sheet3':
            key_str = '0741'
        if key_str.isdigit():
            # Standardize sheet keys to ZRCXXXX format for consistent matching
            species_sheets[f"ZRC{key_str}"] = v

    # Track missing mappings
    missing_sheets = []   # images but no sheet
    sheet_no_images = []  # sheet but no images

    # Get all available images
    image_files = [f for f in os.listdir(raw_img_dir) if f.lower().endswith((".jpeg", ".jpg"))]


    # Iterate through mapping table
    for _, row in mapping_df.iterrows():
        species_id = row["SpeciesID"]   # e.g. ZRC1234
        order = str(row["Order"]).strip()
        family = str(row["Family"]).strip()
        genus = str(row["Genus "]).strip()

        # Construct species folder path
        species_folder = os.path.join(output_dir, order, family, genus, species_id)
        os.makedirs(species_folder, exist_ok=True)

        # ---- Copy images ----
        numeric_id = species_id.replace("ZRC", "")
        # Match images that start with ZRCxxxx or ZRCPORxxxx
        matched_imgs = [
            img for img in image_files
            if img.startswith(f"ZRC{numeric_id}") or img.startswith(f"ZRCPOR{numeric_id}")
        ]
        for fname in matched_imgs:
            src_path = os.path.join(raw_img_dir, fname)
            dst_path = os.path.join(species_folder, fname)
            shutil.copy2(src_path, dst_path)

        # ---- Save species sheet ----
        if species_id in species_sheets:
            sheet_df = species_sheets[species_id]
            out_excel = os.path.join(species_folder, f"{species_id}_measurements.xlsx")
            sheet_df.to_excel(out_excel, index=False)
        else:
            missing_sheets.append(species_id)

    # ---- Handle sheets that have no images ----
    mapped_species_ids = set(mapping_df["SpeciesID"].tolist())
    for species_id in species_sheets.keys():
        numeric_id = species_id.replace("ZRC", "")
        # Check if any image file matches either ZRC or ZRCPOR prefix
        has_image = any(
            img.startswith(f"ZRC{numeric_id}") or img.startswith(f"ZRCPOR{numeric_id}")
            for img in image_files
        )

        if not has_image:
            # Find taxonomic info from mapping
            row = mapping_df[mapping_df["SpeciesID"] == species_id]
            if not row.empty:
                order = str(row.iloc[0]["Order"]).strip()
                family = str(row.iloc[0]["Family"]).strip()
                genus = str(row.iloc[0]["Genus "]).strip()

                species_folder = os.path.join(output_dir, order, family, genus, species_id)
                os.makedirs(species_folder, exist_ok=True)

                # Save sheet into folder
                sheet_df = species_sheets[species_id]
                out_excel = os.path.join(species_folder, f"{species_id}_measurements.xlsx")
                sheet_df.to_excel(out_excel, index=False)

                sheet_no_images.append(species_id)

    # ---- Print report ----
    if missing_sheets:
        print("⚠️ Species with images but missing measurement sheets:")
        unique_missing = sorted(list(set(missing_sheets)))
        for sp in unique_missing:
            print("   -", sp)

    if sheet_no_images:
        print("ℹ️ Species with measurement sheets but no images:")
        unique_sheet_no_images = sorted(list(set(sheet_no_images)))
        for sp in unique_sheet_no_images:
            print("   -", sp)

    print("✅ Image + metadata organization complete.")


if __name__ == "__main__":
    raw_images_folder = "spongephotos"         # your images folder
    metadata_path = "metadata_sponge.xlsx"     # Excel file
    output_folder = "output_repository"        # output directory

    organize_spicule_images(raw_images_folder, metadata_path, output_folder)
