import os
import pandas as pd
def find_measurement_files(root_dir):
    files = []
    for subdir, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith('_measurements.xlsx'):
                files.append(os.path.join(subdir, filename))
    return files

def extract_spicule_types(file_path):
    try:
        df = pd.read_excel(file_path, header=None)
        spicules = set()
        for val in df[0]:
            if isinstance(val, str):
                v = val.strip().lower()
                if v and v != 'spicule type':
                    spicules.add(v)
        return spicules
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return set()

def main():
    root = 'output_repository'
    outdir = 'character_matrix/spicule_samples'
    os.makedirs(outdir, exist_ok=True)
    files = find_measurement_files(root)
    spicule_to_samples = {}
    for f in files:
        sample_name = os.path.splitext(os.path.basename(f))[0].replace('_measurements','')
        spicules = extract_spicule_types(f)
        for s in spicules:
            if s not in spicule_to_samples:
                spicule_to_samples[s] = []
            spicule_to_samples[s].append(sample_name)
    for spicule, samples in spicule_to_samples.items():
        fname = os.path.join(outdir, f"{spicule.replace(' ','_')}_samples.txt")
        with open(fname, 'w') as f:
            for s in sorted(samples):
                f.write(s + '\n')
    print(f"Done! Created {len(spicule_to_samples)} spicule sample lists in {outdir}")

if __name__ == '__main__':
    main()
