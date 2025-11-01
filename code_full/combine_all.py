import os
import pandas as pd

# === USER INPUT ===
main_path = input("Please enter the main directory path: ").strip()

# === SUBFOLDERS ===
folders = ["ترازنامه", "سود و زیان", "نسبت های مالی"]

# === DETECT ALL COMPANIES ===
companies = set()

for folder in folders:
    folder_path = os.path.join(main_path, folder)
    if not os.path.exists(folder_path):
        print(f"⚠️ Folder not found: {folder_path}")
        continue
    
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            # Extract company name (everything before the first space)
            company_name = file.split(" ")[0]
            companies.add(company_name)

companies = sorted(list(companies))
print(f"✅ Found {len(companies)} companies: {companies}")

# === PREPARE OUTPUT FILES ===
output_all = os.path.join(main_path, "Merged_All.xlsx")
missing_log = []

# Create the all-in-one Excel writer
writer_all = pd.ExcelWriter(output_all, engine='openpyxl')

# === MAIN LOOP ===
for company_name in companies:
    print(f"\n🔹 Processing company: {company_name}")
    data_parts = {}
    missing_parts = []

    for folder in folders:
        file_path = os.path.join(main_path, folder, f"{company_name} {folder}.xlsx")

        if not os.path.exists(file_path):
            print(f"  ❌ Missing file in {folder}")
            missing_parts.append(folder)
            continue

        try:
            df = pd.read_excel(file_path)
            data_parts[folder] = df
            print(f"  ✅ Loaded: {folder}")
        except Exception as e:
            print(f"  ⚠️ Error reading {folder}: {e}")
            missing_parts.append(folder)

    # === COMBINE IF ANY DATA EXISTS ===
    if data_parts:
        combined_data = pd.DataFrame()
        for section_name, df in data_parts.items():
            separator = pd.DataFrame([[f"--- {section_name} ---"]], columns=[df.columns[0]])
            combined_data = pd.concat([combined_data, separator, df], ignore_index=True)

        # (1) Save individual merged file
        output_single = os.path.join(main_path, f"{company_name}_Merged.xlsx")
        combined_data.to_excel(output_single, index=False)
        print(f"  💾 Created individual merged file: {output_single}")

        # (2) Add to combined file (all companies)
        combined_data.to_excel(writer_all, sheet_name=company_name[:31], index=False)

    else:
        print(f"  ⚠️ No data found for {company_name} — skipped")

    # Record missing info
    if missing_parts:
        missing_log.append({"Company": company_name, "Missing Sections": ", ".join(missing_parts)})

# Save the all-in-one Excel
writer_all.close()
print(f"\n✅ All-in-one file created: {output_all}")

# === CREATE MISSING REPORT ===
if missing_log:
    missing_df = pd.DataFrame(missing_log)
    output_missing = os.path.join(main_path, "Missing_Report.xlsx")
    missing_df.to_excel(output_missing, index=False)
    print(f"⚠️ Missing data report saved: {output_missing}")
else:
    print("✅ No missing files detected.")

print("\n🎯 Merging process completed successfully.")
