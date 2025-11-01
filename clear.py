import os
import pandas as pd

# === USER INPUT ===
main_path = input("Please enter the main directory path: ").strip()
company_name = "دعبید"  # change this to your test company name

# === SUBFOLDERS ===
folders = ["ترازنامه", "سود و زیان", "نسبت های مالی"]

# === ROWS TO REMOVE (exact or partial match) ===
rows_to_remove = [
    "تصویر اطلاعیه",
    "مالیات سال قبل",
    "سهم اقلیت از سود سال جاری",
    "سود(زیان)فروش دارائیهای زیستی مولد",
    "اقلام غیر مترقبه",
    "اثرات انباشته تغییر در اصول و روشهای"
]

# === FILE PATHS ===
output_single = os.path.join(main_path, f"{company_name}_Merged.xlsx")
output_all = os.path.join(main_path, "Merged_All.xlsx")

# === FUNCTION: CLEAN DATAFRAME ===
def clean_dataframe(df):
    if df.empty:
        return df
    # Find which column contains row labels (usually the first one)
    first_col = df.columns[0]
    # Drop rows where the first column contains any of the unwanted phrases
    for phrase in rows_to_remove:
        df = df[~df[first_col].astype(str).str.contains(phrase, na=False)]
    df.reset_index(drop=True, inplace=True)
    return df

# === IF MERGED FILE ALREADY EXISTS, LOAD AND CLEAN ===
if os.path.exists(output_single):
    print(f"⚙️ Existing merged file found for {company_name}. Applying cleaning only...")
    combined_data = pd.read_excel(output_single)
    combined_data = clean_dataframe(combined_data)
else:
    print(f"🧩 Merging data for {company_name} ...")
    data_parts = {}
    for folder in folders:
        file_path = os.path.join(main_path, folder, f"{company_name} {folder}.xlsx")
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            continue
        df = pd.read_excel(file_path)
        data_parts[folder] = df

    combined_data = pd.DataFrame()
    for section_name, df in data_parts.items():
        separator = pd.DataFrame([[f"--- {section_name} ---"]], columns=[df.columns[0]])
        combined_data = pd.concat([combined_data, separator, df], ignore_index=True)

    combined_data = clean_dataframe(combined_data)

# === SAVE UPDATED / CLEANED FILES ===
combined_data.to_excel(output_single, index=False)
print(f"✅ Cleaned and saved merged file: {output_single}")

with pd.ExcelWriter(output_all, engine='openpyxl', mode='w') as writer:
    combined_data.to_excel(writer, sheet_name=company_name, index=False)
print(f"✅ Cleaned data also saved in: {output_all}")
