# zscore_xgboost_shap.py
import pandas as pd
import numpy as np
from sklearn.model_selection import GroupKFold, cross_val_score
from sklearn.metrics import r2_score, mean_squared_error, mean_absolute_percentage_error
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import StandardScaler
from xgboost import XGBRegressor
import shap
import matplotlib.pyplot as plt
import seaborn as sns

# === Load Data ===
file_path = input("Enter the Excel file name (e.g., Cleaned_Features_WithZ_ML_ready.xlsx): ").strip()
df = pd.read_excel(file_path)

# === Prepare Data ===
drop_cols = ['Year', 'Company']
df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors='ignore')

y = df['Z_next']
X = df.drop(columns=['Z_next'])

# Handle missing values
imputer = SimpleImputer(strategy='median')
X = pd.DataFrame(imputer.fit_transform(X), columns=X.columns)

# Scale
scaler = StandardScaler()
X_scaled = pd.DataFrame(scaler.fit_transform(X), columns=X.columns)

# === Cross-validation (Group by company if available) ===
if 'Company' in df.columns:
    groups = df['Company']
else:
    groups = np.arange(len(df))

gkf = GroupKFold(n_splits=5)

model = XGBRegressor(
    n_estimators=400,
    learning_rate=0.05,
    max_depth=6,
    subsample=0.8,
    colsample_bytree=0.8,
    random_state=42
)

scores = cross_val_score(model, X_scaled, y, cv=gkf, groups=groups, scoring='r2')
print(f"\nCross-validated R²: {scores.mean():.3f} ± {scores.std():.3f}")

# === Fit Final Model ===
model.fit(X_scaled, y)

# === SHAP Explainability ===
explainer = shap.Explainer(model, X_scaled)
shap_values = explainer(X_scaled)

# Summary plot
shap.summary_plot(shap_values, X_scaled, plot_type="bar")
shap.summary_plot(shap_values, X_scaled)

# Feature importance (XGBoost native)
importance = model.feature_importances_
imp_df = pd.DataFrame({'Feature': X.columns, 'Importance': importance}).sort_values('Importance', ascending=False)

plt.figure(figsize=(10,6))
sns.barplot(x='Importance', y='Feature', data=imp_df.head(15))
plt.title("Top 15 Important Features (XGBoost)")
plt.show()
