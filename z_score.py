import pandas as pd
import matplotlib.pyplot as plt
import os

# Step 1: Input
print("Please place your Excel file in the same folder as this script.")
file_name = input("Enter your Excel file name (e.g., pharma_data.xlsx): ").strip()

# Step 2: Load File
try:
    df = pd.read_excel(file_name)
    print(f"✅ File '{file_name}' loaded successfully.")
except FileNotFoundError:
    print("❌ File not found. Make sure it is in the same folder.")
    exit()

# Step 3: Detect Years Automatically
years = [col for col in df.columns if any(str(y) in str(col) for y in range(1390, 1500))]
if not years:
    print("⚠️ No year-like columns found. Please check your file structure.")
    exit()
print(f"📆 Detected years: {years}")

# Step 4: Convert to row-wise dataframe
df_rows = df.set_index(df.columns[0])

# Step 5: Required rows
req_rows = {
    "CurrentAssets": "جمع دارایی‌های جاری",
    "CurrentLiabilities": "جمع بدهی‌های جاری",
    "TotalAssets": "جمع کل دارایی‌ها",
    "TotalLiabilities": "جمع کل بدهی‌ها",
    "RetainedEarnings": "سود و زیان انباشته در پایان دوره",
    "EBIT": "سود (زیان) عملیاتی",
    "Sales": "جمع درآمدها",
    "Equity": "جمع حقوق صاحبان سهام"
}

# Step 6: Compute Altman Z per year
z_scores = {}
missing_years = []  # to store years that couldn't be calculated

for year in years:
    try:
        X1 = (df_rows.loc[req_rows["CurrentAssets"], year] - df_rows.loc[req_rows["CurrentLiabilities"], year]) / df_rows.loc[req_rows["TotalAssets"], year]
        X2 = df_rows.loc[req_rows["RetainedEarnings"], year] / df_rows.loc[req_rows["TotalAssets"], year]
        X3 = df_rows.loc[req_rows["EBIT"], year] / df_rows.loc[req_rows["TotalAssets"], year]
        X4 = df_rows.loc[req_rows["Equity"], year] / df_rows.loc[req_rows["TotalLiabilities"], year]
        X5 = df_rows.loc[req_rows["Sales"], year] / df_rows.loc[req_rows["TotalAssets"], year]

        if any(pd.isna([X1, X2, X3, X4, X5])):
            missing_years.append(year)
            continue

        z = 1.2*X1 + 1.4*X2 + 3.3*X3 + 0.6*X4 + 1.0*X5
        z_scores[year] = z

    except KeyError:
        missing_years.append(year)
    except Exception as e:
        print(f"⚠️ Unexpected error for year {year}: {e}")
        missing_years.append(year)

# Step 7: Create Z DataFrame
if not z_scores:
    print("❌ No valid Z-Scores could be calculated. Please check your data.")
    exit()

z_df = pd.DataFrame(list(z_scores.items()), columns=["Year", "Altman Z"]).sort_values("Year")

# Step 8: Risk Category
def classify_z(z):
    if z > 2.99:
        return "Safe Zone"
    elif z >= 1.81:
        return "Grey Zone"
    else:
        return "Distress Zone"

z_df["Risk Category"] = z_df["Altman Z"].apply(classify_z)

# Step 9: Save with chart in Excel
base_name = os.path.splitext(file_name)[0]
output_file = f"{base_name}_AltmanZ_Report.xlsx"
version = 1
while os.path.exists(output_file):
    version += 1
    output_file = f"{base_name}_AltmanZ_Report_v{version}.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook = writer.book

    # Write Z table
    z_df.to_excel(writer, sheet_name="Z_Score_Table", index=False)
    worksheet = writer.sheets["Z_Score_Table"]

    # Create a chart
    chart = workbook.add_chart({"type": "line"})
    chart.add_series({
        "name": "Altman Z-Score",
        "categories": f"=Z_Score_Table!$A$2:$A${len(z_df)+1}",
        "values": f"=Z_Score_Table!$B$2:$B${len(z_df)+1}",
        "data_labels": {"value": True},
        "line": {"color": "#008080"}
    })
    chart.set_title({"name": "Altman Z-Score Trend Over Years"})
    chart.set_x_axis({"name": "Year"})
    chart.set_y_axis({"name": "Z-Score"})
    chart.set_style(10)

    # Insert chart next to the table
    worksheet.insert_chart("E2", chart)

print(f"💾 Report saved as: {output_file}")

# Step 10: Inform about missing years
if missing_years:
    print(f"⚠️ Z-Score could not be calculated for the following years (missing data): {missing_years}")
else:
    print("✅ Z-Score calculated successfully for all detected years.")
