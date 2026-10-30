import pandas as pd
import joblib
from data.ml.custom_linear_regression import CustomLinearRegression

# Load your dataset
data = pd.read_csv("student_dataset_with_predictions.csv")

X = data.drop("final_score", axis=1).values
y = data["final_score"].values

# Train model
model = CustomLinearRegression(
    learning_rate=0.01,
    epochs=1500
)
model.fit(X, y)

# Save trained model
joblib.dump(model, "models/custom_linear_model.pkl")

print("Custom Linear Regression model saved successfully!")
