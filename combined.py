import os
import pandas as pd

# === USER INPUT ===
main_path = input("Please enter the main directory path: ").strip()
company_name = "دعبید"  # change this to your test company name

# === SUBFOLDERS ===
folders = ["ترازنامه", "سود و زیان", "نسبت های مالی"]

# === READ AND MERGE ===
data_parts = {}

for folder in folders:
    file_path = os.path.join(main_path, folder, f"{company_name} {folder}.xlsx")
    
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        continue
    
    df = pd.read_excel(file_path)
    data_parts[folder] = df

# === COMBINE DATA ===
combined_data = pd.DataFrame()

for section_name, df in data_parts.items():
    separator = pd.DataFrame([[f"--- {section_name} ---"]], columns=[df.columns[0]])
    combined_data = pd.concat([combined_data, separator, df], ignore_index=True)

# === SAVE MERGED OUTPUTS ===
# (1) Individual merged file
output_single = os.path.join(main_path, f"{company_name}_Merged.xlsx")
combined_data.to_excel(output_single, index=False)
print(f"✅ Merged file created for {company_name}: {output_single}")

# (2) All-in-one file (with one sheet per company)
output_all = os.path.join(main_path, "Merged_All.xlsx")
with pd.ExcelWriter(output_all, engine='openpyxl', mode='w') as writer:
    combined_data.to_excel(writer, sheet_name=company_name, index=False)
print(f"✅ Combined file created: {output_all}")
