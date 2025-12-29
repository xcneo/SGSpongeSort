import os
import pandas as pd
import re
from collections import defaultdict

def extract_zrc_codes(text):
    """Extracts and normalizes ZRC codes from a string."""
    if not isinstance(text, str):
        return []
    text = text.replace("ZRC.POR.", "ZRC").replace("ZRCPOR", "ZRC")
    return re.findall(r'ZRC\d+', text)

def check_genus_representation(output_dir, metadata_file):
    """
    Checks if at least one species from each genus has photos and measurements,
    and provides a summary of species counts per genus.
    """
    # --- 1. Find all species with photos and/or measurements in the output directory ---
    species_with_photos = set()
    species_with_measurements = set()
    if os.path.exists(output_dir):
        for root, dirs, files in os.walk(output_dir):
            if not dirs:  # Deepest level directories
                species_id = os.path.basename(root)
                if any(f.lower().endswith((".jpg", ".jpeg")) for f in files):
                    species_with_photos.add(species_id)
                if any(f.lower().endswith(".xlsx") for f in files):
                    species_with_measurements.add(species_id)

    # --- 2. Load metadata and analyze by genus ---
    try:
        mapping_df = pd.read_excel(metadata_file, sheet_name="Sorted Slides List")
        # Clean up column names, especially the one with a trailing space
        mapping_df.columns = [col.strip() for col in mapping_df.columns]
    except FileNotFoundError:
        print(f"Error: Metadata file not found at '{metadata_file}'")
        return
    except ValueError:
        print(f"Error: 'Sorted Slides List' sheet not found in '{metadata_file}'")
        return

    # Create a mapping from genus to a list of its species IDs
    genus_to_species = defaultdict(list)
    for _, row in mapping_df.iterrows():
        genus = str(row.get("Genus", "Unknown")).strip()
        species_ids = extract_zrc_codes(row.get("Parent ZRC", ""))
        if genus and species_ids:
            genus_to_species[genus].extend(species_ids)

    # --- 3. Generate the report data ---
    report_data = []
    for genus, species_list in sorted(genus_to_species.items()):
        unique_species = set(species_list)
        has_picture = any(sp_id in species_with_photos for sp_id in unique_species)
        has_measurement = any(sp_id in species_with_measurements for sp_id in unique_species)
        report_data.append({
            "Genus": genus,
            "Total Species": len(unique_species),
            "Has Picture": "Yes" if has_picture else "No",
            "Has Measurement": "Yes" if has_measurement else "No"
        })

    # --- 4. Print the formatted report ---
    print("--- Genus Representation Report ---")
    if not report_data:
        print("\nNo genus information found or processed.")
    else:
        # Create a formatted table
        headers = ["Genus", "Total Species", "Has Picture", "Has Measurement"]
        # Calculate column widths
        col_widths = {h: len(h) for h in headers}
        for item in report_data:
            for h in headers:
                col_widths[h] = max(col_widths[h], len(str(item[h])))
        
        # Print header
        header_line = " | ".join(h.ljust(col_widths[h]) for h in headers)
        print("\n" + header_line)
        print("-" * len(header_line))

        # Print rows
        for item in report_data:
            row_line = " | ".join(str(item[h]).ljust(col_widths[h]) for h in headers)
            print(row_line)

    print("\n--- Report Complete ---")


if __name__ == "__main__":
    output_folder = "output_repository"
    metadata_path = "metadata_sponge.xlsx"
    check_genus_representation(output_folder, metadata_path)
