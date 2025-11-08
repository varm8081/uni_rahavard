import os
import pandas as pd
import glob

# === USER INPUT ===
main_path = input("Please enter the main directory path: ").strip()

# === ROWS TO REMOVE (exact or partial match) ===
rows_to_remove = [
    "تصویر اطلاعیه",
    "مالیات سال قبل",
    "سهم اقلیت از سود سال جاری",
    "سود(زیان)فروش دارائیهای زیستی مولد",
    "اقلام غیر مترقبه",
    "اثرات انباشته تغییر در اصول و روشهای"
]

# === OUTPUT FILE PATHS ===
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

# === FUNCTION: EXTRACT COMPANY NAME FROM FILENAME ===
def extract_company_name(filename):
    """Extract company name from filename (format: CompanyName_Merged.xlsx)"""
    base_name = os.path.basename(filename)
    if '_Merged.xlsx' in base_name:
        return base_name.split('_Merged.xlsx')[0]
    elif '_' in base_name:
        return base_name.split('_')[0]
    else:
        return base_name.split('.')[0]  # Fallback: remove extension only

# === FIND ALL MERGED FILES ===
pattern = os.path.join(main_path, "*_Merged.xlsx")
company_files = glob.glob(pattern)

print(f"📁 Found {len(company_files)} company files in directory")

if not company_files:
    print("❌ No company files found with pattern '*_Merged.xlsx'")
    exit()

# === PROCESS EACH COMPANY FILE ===
all_company_data = {}

for file_path in company_files:
    company_name = extract_company_name(file_path)
    print(f"⚙️ Processing: {company_name}")
    
    try:
        # Read the company file
        df = pd.read_excel(file_path)
        
        # Clean the dataframe
        cleaned_df = clean_dataframe(df)
        
        # Save cleaned version (overwrite original)
        cleaned_df.to_excel(file_path, index=False)
        print(f"  ✅ Cleaned and saved: {os.path.basename(file_path)}")
        
        # Store for combined all file
        all_company_data[company_name] = cleaned_df
        
    except Exception as e:
        print(f"  ❌ Error processing {company_name}: {str(e)}")

# === CREATE MERGED ALL FILE ===
if all_company_data:
    print(f"\n📊 Creating combined file with {len(all_company_data)} companies...")
    
    with pd.ExcelWriter(output_all, engine='openpyxl') as writer:
        for company_name, df in all_company_data.items():
            # Truncate sheet name if too long (Excel limit: 31 characters)
            sheet_name = company_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  ✅ Added sheet: {sheet_name}")
    
    print(f"🎉 Successfully created: {output_all}")
    
    # Summary
    print(f"\n📈 PROCESSING SUMMARY:")
    print(f"   Total companies processed: {len(all_company_data)}")
    print(f"   Combined file: {output_all}")
    
else:
    print("❌ No data available to create combined file")

print("\n✅ All operations completed!")