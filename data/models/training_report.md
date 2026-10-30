
# Student Performance Prediction - Machine Learning Report

## Training Summary
- **Date**: 2025-11-26 16:19:24
- **Dataset**: 50000 records
- **Features**: 8 features
- **Target**: Prediction Score

## Feature Columns
Attendance (%), Assignment (%), Study Time (hrs), Previous Score (%), Extracurricular (%), Department, Year, Status

## Model Performance Results


### LinearRegression
- **R² Score**: 1.0000
- **Mean Squared Error**: 0.0000
- **Mean Absolute Error**: 0.0025
- **Cross-Validation Score**: 1.0000 (±0.0000)

### KNNRegressor
- **R² Score**: 0.9691
- **Mean Squared Error**: 2.0285
- **Mean Absolute Error**: 1.0855
- **Cross-Validation Score**: 0.9663 (±0.0006)
- **Best Parameters**: {'metric': 'euclidean', 'n_neighbors': 11, 'weights': 'distance'}

### DecisionTreeRegressor
- **R² Score**: 0.9830
- **Mean Squared Error**: 1.1192
- **Mean Absolute Error**: 0.8250
- **Cross-Validation Score**: 0.9823 (±0.0004)
- **Best Parameters**: {'criterion': 'squared_error', 'max_depth': 20, 'min_samples_leaf': 2, 'min_samples_split': 5}

## Best Model
**LinearRegression** with R² score of **1.0000**

## Files Generated
- `models/` - Trained models and preprocessing objects
- `plots/` - Performance visualization plots
- `models/training_results.csv` - Detailed results
