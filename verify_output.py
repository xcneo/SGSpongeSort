import os
import pandas as pd

def verify_output_repository(output_dir, raw_img_dir, metadata_file):
    """
    Verifies the contents of the output repository and checks for unused source files.
    """
    # --- 1. Verify contents of the output repository ---
    photos_no_excel = []
    excel_no_photos = []
    dest_photos = set()
    dest_excels = set()

    if os.path.exists(output_dir):
        for root, dirs, files in os.walk(output_dir):
            if not dirs:  # Deepest level directories
                has_photos = any(f.lower().endswith((".jpg", ".jpeg")) for f in files)
                has_excel = any(f.lower().endswith(".xlsx") for f in files)
                species_folder = os.path.basename(root)

                if has_photos and not has_excel:
                    photos_no_excel.append(species_folder)
                elif not has_photos and has_excel:
                    excel_no_photos.append(species_folder)
                
                for f in files:
                    if f.lower().endswith((".jpg", ".jpeg")):
                        dest_photos.add(f)
                    elif f.lower().endswith(".xlsx"):
                        dest_excels.add(f)

    # --- 2. Check for unused source files ---
    # Get source photos
    source_photos = set([f for f in os.listdir(raw_img_dir) if f.lower().endswith((".jpeg", ".jpg"))])
    unused_photos = source_photos - dest_photos

    # Get source excel sheets
    all_sheets = pd.read_excel(metadata_file, sheet_name=None)
    source_sheet_keys = {str(k).strip() for k in all_sheets if str(k) != "Sorted Slides List" and str(k).strip().isdigit()}
    
    # Normalize destination excel names to match source sheet keys
    # e.g., "ZRC0006_measurements.xlsx" -> "0006"
    dest_sheet_keys = {f.split('_')[0].replace("ZRC", "") for f in dest_excels}
    unused_sheets = source_sheet_keys - dest_sheet_keys


    # --- 3. Print combined report ---
    print("--- Verification Report ---")

    # Section 1: Output Repository Integrity
    print("\n--- Output Repository Integrity ---")
    if photos_no_excel:
        print("\n⚠️ Species with photos but no Excel file:")
        for species in sorted(photos_no_excel):
            print(f"   - {species}")
    else:
        print("\n✅ All species with photos have a corresponding Excel file.")

    if excel_no_photos:
        print("\n⚠️ Species with an Excel file but no photos:")
        for species in sorted(excel_no_photos):
            print(f"   - {species}")
    else:
        print("\n✅ All species with an Excel file have corresponding photos.")

    # Section 2: Source File Usage
    print("\n--- Source File Usage ---")
    if unused_photos:
        print(f"\n⚠️ {len(unused_photos)} unused photos found in '{raw_img_dir}':")
        for photo in sorted(list(unused_photos)):
            print(f"   - {photo}")
    else:
        print(f"\n✅ All photos from '{raw_img_dir}' were used.")

    if unused_sheets:
        print(f"\n⚠️ {len(unused_sheets)} unused measurement sheets found in '{metadata_file}':")
        for sheet in sorted(list(unused_sheets)):
            print(f"   - Sheet '{sheet}'")
    else:
        print(f"\n✅ All measurement sheets from '{metadata_file}' were used.")

    print("\n--- Verification Complete ---")


if __name__ == "__main__":
    output_folder = "output_repository"
    raw_images_folder = "spongephotos"
    metadata_path = "metadata_sponge.xlsx"
    verify_output_repository(output_folder, raw_images_folder, metadata_path)
