# zscore_model_train.py
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, GroupKFold
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import r2_score, mean_squared_error, mean_absolute_percentage_error
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
import matplotlib.pyplot as plt
import seaborn as sns

# --- STEP 1: Load Data ---
file_path = input("Enter the Excel file name (e.g., Cleaned_Features_WithZ_ML_ready.xlsx): ").strip()
df = pd.read_excel(file_path)

# --- STEP 2: Clean Data ---
# drop columns not used for modeling
drop_cols = ['Year', 'Company']
df = df.drop(columns=drop_cols, errors='ignore')

# target variable
y = df['Z_next']
X = df.drop(columns=['Z_next'])

# optional: drop columns with extremely high missing rate (>70%)
missing_ratio = X.isna().mean()
X = X.loc[:, missing_ratio < 0.7]

# --- STEP 3: Imputation + Scaling Pipeline ---
preprocess = Pipeline([
    ('imputer', SimpleImputer(strategy='median')),
    ('scaler', StandardScaler())
])

# --- STEP 4: Split Data ---
# Group by company if available, else random
if 'Year_num' in df.columns:
    df['Year_num'] = pd.to_numeric(df['Year_num'], errors='coerce')

if 'Company' in df.columns:
    groups = df['Company']
    gkf = GroupKFold(n_splits=5)
    splits = list(gkf.split(X, y, groups=groups))
    train_idx, test_idx = splits[0]  # first fold
    X_train, X_test = X.iloc[train_idx], X.iloc[test_idx]
    y_train, y_test = y.iloc[train_idx], y.iloc[test_idx]
else:
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# --- STEP 5: Train Model ---
model = RandomForestRegressor(
    n_estimators=300,
    max_depth=10,
    random_state=42,
    n_jobs=-1
)

pipe = Pipeline([
    ('prep', preprocess),
    ('model', model)
])

pipe.fit(X_train, y_train)
y_pred = pipe.predict(X_test)

# --- STEP 6: Evaluation ---
r2 = r2_score(y_test, y_pred)
#rmse = mean_squared_error(y_test, y_pred, squared=False)
mape = mean_absolute_percentage_error(y_test, y_pred)

print("\n📊 Model Performance:")
print(f"R² Score: {r2:.3f}")
#print(f"RMSE: {rmse:.3f}")
print(f"MAPE: {mape:.2%}")

# --- STEP 7: Feature Importance ---
model_fitted = pipe.named_steps['model']
importances = model_fitted.feature_importances_
features = X.columns

imp_df = pd.DataFrame({'Feature': features, 'Importance': importances})
imp_df = imp_df.sort_values('Importance', ascending=False)

plt.figure(figsize=(10,6))
sns.barplot(x='Importance', y='Feature', data=imp_df.head(15))
plt.title("Top 15 Important Features for Predicting Z_next")
plt.tight_layout()
plt.show()

# --- STEP 8: Compare Predictions ---
comparison = pd.DataFrame({'Actual': y_test, 'Predicted': y_pred})
plt.figure(figsize=(6,6))
sns.scatterplot(x='Actual', y='Predicted', data=comparison)
plt.plot([y.min(), y.max()], [y.min(), y.max()], 'r--')
plt.title("Actual vs Predicted Z_next")
plt.show()
