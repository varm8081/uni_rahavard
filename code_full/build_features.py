# build_features.py
import os
import pandas as pd
import re

# ---------- Utility functions ----------
# map Persian/Arabic digits to ASCII digits
PERSIAN_DIGITS = {ord(x): ord(y) for x, y in zip(
    "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")}

def normalize_text(s):
    if pd.isna(s):
        return s
    s = str(s)
    # normalize arabic y/k to persian if present
    s = s.replace('ي', 'ی').replace('ك', 'ک')
    # remove zero-width and non-printable
    s = re.sub(r'[\u200c\u200b\u200e\u200f]', '', s)
    return s.strip()

def to_number(x):
    """Convert a cell to float if possible: handle Persian digits, commas, parentheses for negatives."""
    if pd.isna(x):
        return float('nan')
    s = str(x).strip()
    if s == '':
        return float('nan')
    # replace Persian/Arabic digits
    s = s.translate(PERSIAN_DIGITS)
    # remove currency symbols or non-numeric letters
    s = s.replace(',', '').replace('٬', '')  # comma variants
    s = s.replace(' ', '')
    # handle parentheses negative e.g. (123) => -123
    if re.match(r'^\(.*\)$', s):
        s = '-' + s[1:-1]
    # if s contains non-digit except - and ., coerce
    try:
        return pd.to_numeric(s, errors='coerce')
    except:
        return float('nan')

def find_row_index(df_index, target_phrase):
    """Find an index in df_index that matches target_phrase robustly (exact then substring)."""
    target = normalize_text(target_phrase)
    # exact match
    for idx in df_index:
        if normalize_text(idx) == target:
            return idx
    # substring match
    for idx in df_index:
        if target in normalize_text(idx):
            return idx
    # try reverse: index contained in target (rare)
    for idx in df_index:
        if normalize_text(idx) in target:
            return idx
    return None

# ---------- Mappings (Persian row names used in your file) ----------
REQ_ROWS = {
    "CurrentAssets": "جمع دارایی‌های جاری",
    "CurrentLiabilities": "جمع بدهی‌های جاری",
    "TotalAssets": "جمع کل دارایی‌ها",
    "TotalLiabilities": "جمع کل بدهی‌ها",
    "RetainedEarnings": "سود و زیان انباشته در پایان دوره",
    "EBIT": "سود (زیان) عملیاتی",
    "Sales": "جمع درآمدها",
    "Equity": "جمع حقوق صاحبان سهام"
}

# 5 supplemental variables (from your list)
SUPP_ROWS = {
    "ROA": "بازده دارایی‌ها ROA",
    "ROE": "بازدهی سرمایه ROE",
    "CurrentRatio": "نسبت جاری",
    "DebtRatio": "نسبت بدهی",
    "OperatingMargin": "حاشیه سود عملیاتی"
}

# ---------- Main ----------
def main():
    print("Place the combined Excel (each sheet = one company) in same folder as this script.")
    file_path = input("Enter Excel filename (e.g., combined_all_companies.xlsx): ").strip()
    if not os.path.exists(file_path):
        print("File not found. Exiting.")
        return

    # Prepare output structures
    rows_out = []
    missing_log = []

    # read all sheet names (company names)
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names
    print(f"Found {len(sheets)} sheets (companies).")

    for sheet in sheets:
        print(f"Processing company: {sheet} ...")
        df = pd.read_excel(xls, sheet_name=sheet, header=0)
        # first column is row labels (e.g., "سال مالی" header then date columns)
        first_col = df.columns[0]
        # set rows as index
        df_rows = df.set_index(first_col)
        # identify year columns: assume the columns except first_col are the date columns
        year_cols = list(df_rows.columns)
        # convert year labels to simple year (first part before '/'), map column -> year_str
        year_map = {}
        for col in year_cols:
            s = str(col)
            s = normalize_text(s)
            # extract leading year digits
            m = re.match(r'(\d{3,4})', s)
            year = m.group(1) if m else s
            year_map[col] = year

        # for each year column, extract required rows
        for col in year_cols:
            year = year_map[col]
            missing = []
            vals = {}
            # required metrics
            for key, persian_name in REQ_ROWS.items():
                idx = find_row_index(df_rows.index, persian_name)
                if idx is None:
                    vals[key] = float('nan')
                    missing.append((year, sheet, key, persian_name, "row_not_found"))
                else:
                    raw = df_rows.at[idx, col]
                    num = to_number(raw)
                    vals[key] = num
                    if pd.isna(num):
                        missing.append((year, sheet, key, persian_name, "value_na_or_unparseable"))
            # supplemental
            for key, persian_name in SUPP_ROWS.items():
                idx = find_row_index(df_rows.index, persian_name)
                if idx is None:
                    vals[key] = float('nan')
                    missing.append((year, sheet, key, persian_name, "row_not_found"))
                else:
                    raw = df_rows.at[idx, col]
                    num = to_number(raw)
                    vals[key] = num
                    if pd.isna(num):
                        missing.append((year, sheet, key, persian_name, "value_na_or_unparseable"))
            # compute X1..X5 if possible (we will compute in features file too)
            # X1 = (CurrentAssets - CurrentLiabilities) / TotalAssets
            try:
                X1 = (vals["CurrentAssets"] - vals["CurrentLiabilities"]) / vals["TotalAssets"] if not pd.isna(vals["TotalAssets"]) and vals["TotalAssets"] != 0 else float('nan')
            except:
                X1 = float('nan')
            try:
                X2 = vals["RetainedEarnings"] / vals["TotalAssets"] if not pd.isna(vals["TotalAssets"]) and vals["TotalAssets"] != 0 else float('nan')
            except:
                X2 = float('nan')
            try:
                X3 = vals["EBIT"] / vals["TotalAssets"] if not pd.isna(vals["TotalAssets"]) and vals["TotalAssets"] != 0 else float('nan')
            except:
                X3 = float('nan')
            try:
                X4 = vals["Equity"] / vals["TotalLiabilities"] if not pd.isna(vals["TotalLiabilities"]) and vals["TotalLiabilities"] != 0 else float('nan')
            except:
                X4 = float('nan')
            try:
                X5 = vals["Sales"] / vals["TotalAssets"] if not pd.isna(vals["TotalAssets"]) and vals["TotalAssets"] != 0 else float('nan')
            except:
                X5 = float('nan')

            out_row = {
                "Company": sheet,
                "Year": year,
                # raw base numbers
                "CurrentAssets": vals.get("CurrentAssets"),
                "CurrentLiabilities": vals.get("CurrentLiabilities"),
                "TotalAssets": vals.get("TotalAssets"),
                "TotalLiabilities": vals.get("TotalLiabilities"),
                "RetainedEarnings": vals.get("RetainedEarnings"),
                "EBIT": vals.get("EBIT"),
                "Sales": vals.get("Sales"),
                "Equity": vals.get("Equity"),
                # X components
                "X1": X1, "X2": X2, "X3": X3, "X4": X4, "X5": X5,
                # supplemental
                "ROA": vals.get("ROA"),
                "ROE": vals.get("ROE"),
                "CurrentRatio": vals.get("CurrentRatio"),
                "DebtRatio": vals.get("DebtRatio"),
                "OperatingMargin": vals.get("OperatingMargin")
            }
            rows_out.append(out_row)
            # append missing info
            for m in missing:
                missing_log.append({
                    "Year": m[0], "Company": m[1], "Key": m[2], "PersianName": m[3], "Reason": m[4]
                })

    # build DataFrame and save
    df_out = pd.DataFrame(rows_out)
    # sort for readability
    df_out = df_out.sort_values(["Company", "Year"]).reset_index(drop=True)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    out_file = f"{base_name}_Cleaned_Features.xlsx"
    df_out.to_excel(out_file, index=False)
    print(f"Saved features to: {out_file}")

    # save missing log
    if missing_log:
        log_df = pd.DataFrame(missing_log)
        log_file = f"{base_name}_missing_log.csv"
        log_df.to_csv(log_file, index=False, encoding='utf-8-sig')
        print(f"Saved missing-log to: {log_file} (rows: {len(log_df)})")
    else:
        print("No missing entries logged.")

if __name__ == "__main__":
    main()
