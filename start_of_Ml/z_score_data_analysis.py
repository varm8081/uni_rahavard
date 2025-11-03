# zscore_data_analysis.py
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import skew, kurtosis

# === Load Data ===
file_path = input("Enter the Excel file name (e.g., Cleaned_Features_WithZ_ML_ready.xlsx): ").strip()
df = pd.read_excel(file_path)

print("\n✅ Data Loaded Successfully!")
print(f"Shape: {df.shape}")
print("\n📋 Columns:", df.columns.tolist())

# === Basic Info ===
print("\n📊 Basic Info:")
print(df.describe().T)
print("\nMissing Values (%):")
print((df.isnull().mean() * 100).round(2))

# === Distribution Check ===
numeric_cols = df.select_dtypes(include=np.number).columns
df[numeric_cols].hist(bins=30, figsize=(18,12))
plt.suptitle("Numeric Feature Distributions")
plt.show()

# === Skewness & Kurtosis ===
skewness = df[numeric_cols].apply(lambda x: skew(x.dropna()))
kurt = df[numeric_cols].apply(lambda x: kurtosis(x.dropna()))
sk_table = pd.DataFrame({'Skewness': skewness, 'Kurtosis': kurt})
print("\n📈 Skewness & Kurtosis:")
print(sk_table.sort_values('Skewness', ascending=False).head(10))

# === Correlation Matrix ===
corr = df[numeric_cols].corr()
plt.figure(figsize=(12,10))
sns.heatmap(corr, cmap='coolwarm', center=0, annot=False)
plt.title("Correlation Matrix Heatmap")
plt.show()

# === Top Correlations with Z_next ===
target = 'Z_next'
if target in df.columns:
    corr_target = corr[target].sort_values(ascending=False)
    print("\n🔥 Top correlations with Z_next:")
    print(corr_target.head(15))
    corr_target.drop(target).head(15).plot(kind='bar', figsize=(10,5), title='Top Features correlated with Z_next')
    plt.show()

# === Trend Example for a Single Company ===
if 'Company' in df.columns and 'Year' in df.columns:
    company = df['Company'].unique()[0]
    df_company = df[df['Company'] == company].sort_values('Year')
    plt.figure(figsize=(10,5))
    plt.plot(df_company['Year'], df_company['Altman_Z'], marker='o', label='Altman_Z')
    plt.plot(df_company['Year'], df_company['Z_next'], marker='x', label='Z_next')
    plt.title(f"Trend of Z for {company}")
    plt.xlabel('Year')
    plt.ylabel('Z Score')
    plt.legend()
    plt.show()

print("\n✅ EDA Completed.")
