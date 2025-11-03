# zscore_xgboost_full_evaluation.py
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import GroupKFold, train_test_split, cross_val_score
from sklearn.metrics import (
    r2_score, mean_squared_error, mean_absolute_error, mean_absolute_percentage_error
)
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import StandardScaler
from xgboost import XGBRegressor
import shap

# === Load Data ===
file_path = input("Enter the Excel file name (e.g., Cleaned_Features_WithZ_ML_ready.xlsx): ").strip()
df = pd.read_excel(file_path)

print(f"\n✅ Data Loaded. Shape: {df.shape}")

# === Prepare Data ===
drop_cols = ['Year', 'Company']
df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors='ignore')

# Target & Features
y = df['Z_next']
X = df.drop(columns=['Z_next'])

# === Handle missing values ===
imputer = SimpleImputer(strategy='median')
X_imputed = pd.DataFrame(imputer.fit_transform(X), columns=X.columns)

# === Scale features ===
scaler = StandardScaler()
X_scaled = pd.DataFrame(scaler.fit_transform(X_imputed), columns=X.columns)

# === Train/Test Split ===
X_train, X_test, y_train, y_test = train_test_split(X_scaled, y, test_size=0.2, random_state=42)

# === Define model ===
model = XGBRegressor(
    n_estimators=400,
    learning_rate=0.05,
    max_depth=6,
    subsample=0.8,
    colsample_bytree=0.8,
    random_state=42,
    objective='reg:squarederror'
)

# === Cross-validation ===
gkf = GroupKFold(n_splits=5) if 'Company' in df.columns else None
if gkf:
    print("\n🔁 Performing Group K-Fold cross-validation...")
    groups = df['Company'] if 'Company' in df.columns else np.arange(len(df))
    cv_scores = cross_val_score(model, X_scaled, y, cv=gkf, groups=groups, scoring='r2')
    print(f"R² (Cross-validated): {cv_scores.mean():.4f} ± {cv_scores.std():.4f}")

# === Fit Model ===
model.fit(X_train, y_train)
y_pred = model.predict(X_test)

# === Evaluation Metrics ===
r2 = r2_score(y_test, y_pred)
adj_r2 = 1 - (1 - r2) * (len(y_test) - 1) / (len(y_test) - X_test.shape[1] - 1)
rmse = np.sqrt(mean_squared_error(y_test, y_pred))
mae = mean_absolute_error(y_test, y_pred)
mape = mean_absolute_percentage_error(y_test, y_pred) * 100

print("\n📊 Model Evaluation Results:")
print(f"R² Score: {r2:.4f}")
print(f"Adjusted R²: {adj_r2:.4f}")
print(f"RMSE: {rmse:.4f}")
print(f"MAE: {mae:.4f}")
print(f"MAPE: {mape:.2f}%")

# === Visualization: Actual vs Predicted ===
plt.figure(figsize=(7,6))
sns.scatterplot(x=y_test, y=y_pred, alpha=0.7)
plt.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'r--', label="Perfect Prediction")
plt.xlabel("Actual Z_next")
plt.ylabel("Predicted Z_next")
plt.title("Actual vs Predicted Z_next")
plt.legend()
plt.show()

# === Residual Plot ===
residuals = y_test - y_pred
plt.figure(figsize=(7,5))
sns.histplot(residuals, bins=30, kde=True)
plt.title("Residuals Distribution")
plt.xlabel("Residual (Actual - Predicted)")
plt.show()

# === Feature Importance ===
importance = model.feature_importances_
imp_df = pd.DataFrame({'Feature': X.columns, 'Importance': importance}).sort_values('Importance', ascending=False)

plt.figure(figsize=(10,6))
sns.barplot(x='Importance', y='Feature', data=imp_df.head(15))
plt.title("Top 15 Important Features (XGBoost)")
plt.show()

# === SHAP Explainability ===
print("\n🔍 Computing SHAP values (this may take a bit)...")
explainer = shap.Explainer(model, X_train)
shap_values = explainer(X_test)

shap.summary_plot(shap_values, X_test, plot_type="bar", show=False)
plt.title("SHAP Feature Importance (mean absolute value)")
plt.show()

shap.summary_plot(shap_values, X_test, show=False)
plt.title("SHAP Summary Plot")
plt.show()

print("\n✅ Full evaluation completed successfully!")
