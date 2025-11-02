# compute_z_and_targets.py
import os
import pandas as pd

print("Place the Cleaned_Features.xlsx in same folder as this script (or give filename).")
file_name = input("Enter cleaned features filename (e.g., combined_Cleaned_Features.xlsx): ").strip()
if not os.path.exists(file_name):
    print("File not found. Exiting.")
    raise SystemExit

df = pd.read_excel(file_name)

# Ensure numeric columns exist and numeric typed
num_cols = ["CurrentAssets","CurrentLiabilities","TotalAssets","TotalLiabilities",
            "RetainedEarnings","EBIT","Sales","Equity",
            "X1","X2","X3","X4","X5",
            "ROA","ROE","CurrentRatio","DebtRatio","OperatingMargin"]

for c in num_cols:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce')

# Compute Altman Z (use X1..X5 if present, otherwise compute from base numbers if available)
def compute_Xs(row):
    # if X already present, keep; else try to compute
    try:
        X1 = row["X1"] if pd.notna(row.get("X1")) else ((row["CurrentAssets"] - row["CurrentLiabilities"]) / row["TotalAssets"]) if pd.notna(row.get("TotalAssets")) and row["TotalAssets"] != 0 else float('nan')
    except:
        X1 = float('nan')
    try:
        X2 = row["X2"] if pd.notna(row.get("X2")) else (row["RetainedEarnings"] / row["TotalAssets"]) if pd.notna(row.get("TotalAssets")) and row["TotalAssets"] != 0 else float('nan')
    except:
        X2 = float('nan')
    try:
        X3 = row["X3"] if pd.notna(row.get("X3")) else (row["EBIT"] / row["TotalAssets"]) if pd.notna(row.get("TotalAssets")) and row["TotalAssets"] != 0 else float('nan')
    except:
        X3 = float('nan')
    try:
        X4 = row["X4"] if pd.notna(row.get("X4")) else (row["Equity"] / row["TotalLiabilities"]) if pd.notna(row.get("TotalLiabilities")) and row["TotalLiabilities"] != 0 else float('nan')
    except:
        X4 = float('nan')
    try:
        X5 = row["X5"] if pd.notna(row.get("X5")) else (row["Sales"] / row["TotalAssets"]) if pd.notna(row.get("TotalAssets")) and row["TotalAssets"] != 0 else float('nan')
    except:
        X5 = float('nan')
    return pd.Series([X1,X2,X3,X4,X5])

df[["X1_calc","X2_calc","X3_calc","X4_calc","X5_calc"]] = df.apply(compute_Xs, axis=1)

# prefer existing X if present; else use calc
for i in range(1,6):
    df[f"X{i}"] = df[f"X{i}"] if f"X{i}" in df.columns else df[f"X{i}_calc"]
    # if both exists, fillna from calc
    if f"X{i}_calc" in df.columns and f"X{i}" in df.columns:
        df[f"X{i}"] = df[f"X{i}"].fillna(df[f"X{i}_calc"])

# Compute Altman Z
df["Altman_Z"] = 1.2*df["X1"] + 1.4*df["X2"] + 3.3*df["X3"] + 0.6*df["X4"] + 1.0*df["X5"]

# Save with Z
out_with_z = os.path.splitext(file_name)[0] + "_WithZ.xlsx"
df.to_excel(out_with_z, index=False)
print(f"Saved features with Altman_Z to: {out_with_z}")

# Create ML target Z_next by shifting within each company by Year (ensure Year is sortable numeric)
# first make Year numeric
df['Year_num'] = pd.to_numeric(df['Year'], errors='coerce')
df = df.sort_values(['Company','Year_num'])

df['Z_next'] = df.groupby('Company')['Altman_Z'].shift(-1)

# ML ready: drop rows where Z_next is NaN (no next-year available)
ml_df = df.dropna(subset=['Z_next']).copy()

out_ml = os.path.splitext(file_name)[0] + "_ML_ready.xlsx"
ml_df.to_excel(out_ml, index=False)
print(f"Saved ML-ready dataset to: {out_ml}")

# Also save a small report of rows where Altman_Z could not be computed
no_z = df[df['Altman_Z'].isna()][['Company','Year']]
if not no_z.empty:
    noz_file = os.path.splitext(file_name)[0] + "_noZ_report.csv"
    no_z.to_csv(noz_file, index=False, encoding='utf-8-sig')
    print(f"Report of rows without Altman_Z saved to: {noz_file}")
else:
    print("Altman_Z computed for all rows (where enough inputs existed).")
