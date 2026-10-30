import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score, GridSearchCV
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.neighbors import KNeighborsRegressor
from sklearn.linear_model import LinearRegression
from sklearn.tree import DecisionTreeRegressor
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
import matplotlib.pyplot as plt
import seaborn as sns
import joblib
import os
from datetime import datetime

class StudentPerformancePredictor:
    def __init__(self):
        self.models = {}
        self.scaler = StandardScaler()
        self.label_encoders = {}
        self.feature_columns = []
        self.target_column = 'Prediction Score'
        self.results = {}
        
    def load_data(self, file_path):
        """Load and prepare the dataset"""
        print(" Loading dataset...")
        self.df = pd.read_csv(file_path)
        print(f" Dataset loaded: {len(self.df)} records, {len(self.df.columns)} columns")
        return self.df
    
    def preprocess_data(self):
        """Preprocess the data for machine learning"""
        print("\n Preprocessing data...")
        
        # Create a copy for preprocessing
        df_processed = self.df.copy()
        
        # Encode categorical variables
        categorical_columns = ['Department', 'Year', 'Status']
        for col in categorical_columns:
            if col in df_processed.columns:
                le = LabelEncoder()
                df_processed[col] = le.fit_transform(df_processed[col])
                self.label_encoders[col] = le
                print(f"   - Encoded {col}: {len(le.classes_)} categories")
        
        # Select feature columns (exclude target and non-feature columns)
        exclude_columns = [self.target_column, 'Student ID', 'Student Name', 'Recommendation', 'Result']
        self.feature_columns = [col for col in df_processed.columns if col not in exclude_columns]
        
        # Prepare features and target
        X = df_processed[self.feature_columns]
        y = df_processed[self.target_column]
        
        print(f"   - Features: {self.feature_columns}")
        print(f"   - Target: {self.target_column}")
        print(f"   - Feature shape: {X.shape}")
        
        # Split the data
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42
        )
        
        # Scale the features
        X_train_scaled = self.scaler.fit_transform(X_train)
        X_test_scaled = self.scaler.transform(X_test)
        
        self.X_train = X_train
        self.X_test = X_test
        self.y_train = y_train
        self.y_test = y_test
        self.X_train_scaled = X_train_scaled
        self.X_test_scaled = X_test_scaled
        
        print(f" Data preprocessed: Train={X_train.shape[0]}, Test={X_test.shape[0]}")
        return X_train_scaled, X_test_scaled, y_train, y_test
    
    def train_linear_regression(self):
        """Train Linear Regression model"""
        print("\n Training Linear Regression...")
        
        # Train model
        lr_model = LinearRegression()
        lr_model.fit(self.X_train_scaled, self.y_train)
        
        # Make predictions
        y_pred = lr_model.predict(self.X_test_scaled)
        
        # Calculate metrics
        mse = mean_squared_error(self.y_test, y_pred)
        mae = mean_absolute_error(self.y_test, y_pred)
        r2 = r2_score(self.y_test, y_pred)
        
        # Cross-validation
        cv_scores = cross_val_score(lr_model, self.X_train_scaled, self.y_train, cv=5)
        
        self.models['LinearRegression'] = lr_model
        self.results['LinearRegression'] = {
            'MSE': mse,
            'MAE': mae,
            'R2': r2,
            'CV_Mean': cv_scores.mean(),
            'CV_Std': cv_scores.std(),
            'predictions': y_pred
        }
        
        print(f"    Linear Regression trained")
        print(f"    R² Score: {r2:.4f}")
        print(f"    MSE: {mse:.4f}")
        print(f"    MAE: {mae:.4f}")
        print(f"    CV Score: {cv_scores.mean():.4f} (+/- {cv_scores.std() * 2:.4f})")
        
        return lr_model
    
    def train_knn_regressor(self):
        """Train KNN Regressor with hyperparameter tuning"""
        print("\n Training KNN Regressor...")
        
        # Hyperparameter tuning
        param_grid = {
            'n_neighbors': [3, 5, 7, 9, 11, 13, 15],
            'weights': ['uniform', 'distance'],
            'metric': ['euclidean', 'manhattan']
        }
        
        knn = KNeighborsRegressor()
        grid_search = GridSearchCV(
            knn, param_grid, cv=5, scoring='r2', n_jobs=-1, verbose=0
        )
        grid_search.fit(self.X_train_scaled, self.y_train)
        
        best_knn = grid_search.best_estimator_
        y_pred = best_knn.predict(self.X_test_scaled)
        
        # Calculate metrics
        mse = mean_squared_error(self.y_test, y_pred)
        mae = mean_absolute_error(self.y_test, y_pred)
        r2 = r2_score(self.y_test, y_pred)
        
        # Cross-validation
        cv_scores = cross_val_score(best_knn, self.X_train_scaled, self.y_train, cv=5)
        
        self.models['KNNRegressor'] = best_knn
        self.results['KNNRegressor'] = {
            'MSE': mse,
            'MAE': mae,
            'R2': r2,
            'CV_Mean': cv_scores.mean(),
            'CV_Std': cv_scores.std(),
            'predictions': y_pred,
            'best_params': grid_search.best_params_
        }
        
        print(f"    KNN Regressor trained")
        print(f"    Best parameters: {grid_search.best_params_}")
        print(f"    R² Score: {r2:.4f}")
        print(f"    MSE: {mse:.4f}")
        print(f"    MAE: {mae:.4f}")
        print(f"    CV Score: {cv_scores.mean():.4f} (+/- {cv_scores.std() * 2:.4f})")
        
        return best_knn
    
    def train_decision_tree(self):
        """Train Decision Tree Regressor with hyperparameter tuning"""
        print("\n Training Decision Tree Regressor...")
        
        # Hyperparameter tuning
        param_grid = {
            'max_depth': [3, 5, 7, 10, 15, 20, None],
            'min_samples_split': [2, 5, 10, 20],
            'min_samples_leaf': [1, 2, 4, 8],
            'criterion': ['squared_error', 'absolute_error']
        }
        
        dt = DecisionTreeRegressor(random_state=42)
        grid_search = GridSearchCV(
            dt, param_grid, cv=5, scoring='r2', n_jobs=-1, verbose=0
        )
        grid_search.fit(self.X_train, self.y_train)
        
        best_dt = grid_search.best_estimator_
        y_pred = best_dt.predict(self.X_test)
        
        # Calculate metrics
        mse = mean_squared_error(self.y_test, y_pred)
        mae = mean_absolute_error(self.y_test, y_pred)
        r2 = r2_score(self.y_test, y_pred)
        
        # Cross-validation
        cv_scores = cross_val_score(best_dt, self.X_train, self.y_train, cv=5)
        
        self.models['DecisionTreeRegressor'] = best_dt
        self.results['DecisionTreeRegressor'] = {
            'MSE': mse,
            'MAE': mae,
            'R2': r2,
            'CV_Mean': cv_scores.mean(),
            'CV_Std': cv_scores.std(),
            'predictions': y_pred,
            'best_params': grid_search.best_params_
        }
        
        print(f"    Decision Tree Regressor trained")
        print(f"    Best parameters: {grid_search.best_params_}")
        print(f"    R² Score: {r2:.4f}")
        print(f"    MSE: {mse:.4f}")
        print(f"    MAE: {mae:.4f}")
        print(f"    CV Score: {cv_scores.mean():.4f} (+/- {cv_scores.std() * 2:.4f})")
        
        return best_dt
    
    def compare_models(self):
        """Compare all trained models"""
        print("\n Model Comparison")
        print("=" * 80)
        
        comparison_data = []
        for model_name, results in self.results.items():
            comparison_data.append({
                'Model': model_name,
                'R² Score': results['R2'],
                'MSE': results['MSE'],
                'MAE': results['MAE'],
                'CV Score': results['CV_Mean'],
                'CV Std': results['CV_Std']
            })
        
        comparison_df = pd.DataFrame(comparison_data)
        print(comparison_df.to_string(index=False, float_format='%.4f'))
        
        # Find best model
        best_model = comparison_df.loc[comparison_df['R² Score'].idxmax()]
        print(f"\n Best Model: {best_model['Model']} (R² = {best_model['R² Score']:.4f})")
        
        return comparison_df
    
    def plot_results(self):
        """Plot model comparison and predictions"""
        print("\n Generating plots...")
        
        # Create plots directory
        os.makedirs('plots', exist_ok=True)
        
        # 1. Model Performance Comparison
        fig, axes = plt.subplots(2, 2, figsize=(15, 12))
        fig.suptitle('Student Performance Prediction - Model Comparison', fontsize=16)
        
        # R² Scores
        models = list(self.results.keys())
        r2_scores = [self.results[model]['R2'] for model in models]
        axes[0, 0].bar(models, r2_scores, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
        axes[0, 0].set_title('R² Scores')
        axes[0, 0].set_ylabel('R² Score')
        axes[0, 0].tick_params(axis='x', rotation=45)
        
        # MSE Scores
        mse_scores = [self.results[model]['MSE'] for model in models]
        axes[0, 1].bar(models, mse_scores, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
        axes[0, 1].set_title('Mean Squared Error')
        axes[0, 1].set_ylabel('MSE')
        axes[0, 1].tick_params(axis='x', rotation=45)
        
        # MAE Scores
        mae_scores = [self.results[model]['MAE'] for model in models]
        axes[1, 0].bar(models, mae_scores, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
        axes[1, 0].set_title('Mean Absolute Error')
        axes[1, 0].set_ylabel('MAE')
        axes[1, 0].tick_params(axis='x', rotation=45)
        
        # CV Scores
        cv_scores = [self.results[model]['CV_Mean'] for model in models]
        cv_stds = [self.results[model]['CV_Std'] for model in models]
        axes[1, 1].bar(models, cv_scores, yerr=cv_stds, capsize=5, 
                      color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
        axes[1, 1].set_title('Cross-Validation Scores')
        axes[1, 1].set_ylabel('CV Score')
        axes[1, 1].tick_params(axis='x', rotation=45)
        
        plt.tight_layout()
        plt.savefig('plots/model_comparison.png', dpi=300, bbox_inches='tight')
        print("    Model comparison plot saved: plots/model_comparison.png")
        
        # 2. Actual vs Predicted Plots
        fig, axes = plt.subplots(1, 3, figsize=(18, 6))
        fig.suptitle('Actual vs Predicted Values', fontsize=16)
        
        for i, (model_name, results) in enumerate(self.results.items()):
            axes[i].scatter(self.y_test, results['predictions'], alpha=0.6, s=20)
            axes[i].plot([self.y_test.min(), self.y_test.max()], 
                        [self.y_test.min(), self.y_test.max()], 'r--', lw=2)
            axes[i].set_xlabel('Actual Values')
            axes[i].set_ylabel('Predicted Values')
            axes[i].set_title(f'{model_name}\nR² = {results["R2"]:.4f}')
            axes[i].grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig('plots/actual_vs_predicted.png', dpi=300, bbox_inches='tight')
        print("    Actual vs predicted plot saved: plots/actual_vs_predicted.png")
        
        # Close the plot to avoid display issues
        plt.close()
    
    def save_models(self):
        """Save trained models and preprocessing objects"""
        print("\n Saving models...")
        
        # Create models directory
        os.makedirs('models', exist_ok=True)
        
        # Save models
        for model_name, model in self.models.items():
            model_path = f'models/{model_name.lower()}.pkl'
            joblib.dump(model, model_path)
            print(f"    {model_name} saved: {model_path}")
        
        # Save preprocessing objects
        joblib.dump(self.scaler, 'models/scaler.pkl')
        joblib.dump(self.label_encoders, 'models/label_encoders.pkl')
        joblib.dump(self.feature_columns, 'models/feature_columns.pkl')
        
        # Save results
        results_df = pd.DataFrame(self.results).T
        results_df.to_csv('models/training_results.csv')
        
        print("    Preprocessing objects saved")
        print("    Training results saved: models/training_results.csv")
    
    def generate_report(self):
        """Generate a comprehensive training report"""
        print("\n Generating training report...")
        
        report = f"""
# Student Performance Prediction - Machine Learning Report

## Training Summary
- **Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- **Dataset**: {len(self.df)} records
- **Features**: {len(self.feature_columns)} features
- **Target**: {self.target_column}

## Feature Columns
{', '.join(self.feature_columns)}

## Model Performance Results

"""
        
        for model_name, results in self.results.items():
            report += f"""
### {model_name}
- **R² Score**: {results['R2']:.4f}
- **Mean Squared Error**: {results['MSE']:.4f}
- **Mean Absolute Error**: {results['MAE']:.4f}
- **Cross-Validation Score**: {results['CV_Mean']:.4f} (±{results['CV_Std']:.4f})
"""
            
            if 'best_params' in results:
                report += f"- **Best Parameters**: {results['best_params']}\n"
        
        # Find best model
        best_model = max(self.results.items(), key=lambda x: x[1]['R2'])
        report += f"""
## Best Model
**{best_model[0]}** with R² score of **{best_model[1]['R2']:.4f}**

## Files Generated
- `models/` - Trained models and preprocessing objects
- `plots/` - Performance visualization plots
- `models/training_results.csv` - Detailed results
"""
        
        # Save report
        with open('models/training_report.md', 'w') as f:
            f.write(report)
        
        print("    Training report saved: models/training_report.md")
        return report

def main():
    """Main function to run the complete ML pipeline"""
    print("🚀 Starting Student Performance Prediction ML Pipeline")
    print("=" * 80)
    
    # Initialize predictor
    predictor = StudentPerformancePredictor()
    
    # Load data
    predictor.load_data('students_dataset_with_predictions.csv')
    
    # Preprocess data
    predictor.preprocess_data()
    
    # Train models
    predictor.train_linear_regression()
    predictor.train_knn_regressor()
    predictor.train_decision_tree()
    
    # Compare models
    comparison_df = predictor.compare_models()
    
    # Generate plots
    predictor.plot_results()
    
    # Save models
    predictor.save_models()
    
    # Generate report
    report = predictor.generate_report()
    
    print("\n" + "=" * 80)
    print(" ML Pipeline Completed Successfully!")
    print(" Check the 'models/' and 'plots/' directories for results")
    print("=" * 80)

if __name__ == "__main__":
    main()
