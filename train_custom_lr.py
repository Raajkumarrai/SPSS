import pandas as pd
import joblib
from data.ml.custom_linear_regression import CustomLinearRegression

# Load your dataset
data = pd.read_csv("data/students_dataset_with_predictions.csv")

X = data[['Attendance (%)','Assignment (%)','Study Time (hrs)','Previous Score (%)','Extracurricular (%)']].values
y = data['Prediction Score'].values

# Train model
model = CustomLinearRegression(
    learning_rate=0.01,
    epochs=1500
)
model.fit(X, y)

# Save trained model
joblib.dump(model, "models/custom_linear_model.pkl")

print("Custom Linear Regression model saved successfully!")
