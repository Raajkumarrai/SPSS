import pandas as pd
import numpy as np
import joblib
import os
from datetime import datetime

class StudentPerformancePredictor:
    def __init__(self, models_dir='models'):
        """Initialize the predictor with trained models"""
        self.models_dir = models_dir
        self.models = {}
        self.scaler = None
        self.label_encoders = {}
        self.feature_columns = []
        
        # Load models and preprocessing objects
        self.load_models()
    
    def load_models(self):
        """Load trained models and preprocessing objects"""
        print(" Loading trained models...")
        
        try:
            # Load preprocessing objects
            self.scaler = joblib.load(f'{self.models_dir}/scaler.pkl')
            self.label_encoders = joblib.load(f'{self.models_dir}/label_encoders.pkl')
            self.feature_columns = joblib.load(f'{self.models_dir}/feature_columns.pkl')
            
            # Load models
            model_files = {
                'LinearRegression': 'linearregression.pkl',
                'KNNRegressor': 'knnregressor.pkl',
                'DecisionTreeRegressor': 'decisiontreeregressor.pkl'
            }
            
            for model_name, filename in model_files.items():
                model_path = f'{self.models_dir}/{filename}'
                if os.path.exists(model_path):
                    self.models[model_name] = joblib.load(model_path)
                    print(f"    Loaded {model_name}")
                else:
                    print(f"     {model_name} not found at {model_path}")
            
            print(f" Successfully loaded {len(self.models)} models")
            
        except Exception as e:
            print(f" Error loading models: {str(e)}")
            print("Please ensure models are trained first using train_ml_models.py")
    
    def preprocess_input(self, student_data):
        """Preprocess input data for prediction"""
        # Create a copy of the input data
        processed_data = student_data.copy()
        
        # Encode categorical variables
        for col, encoder in self.label_encoders.items():
            if col in processed_data.columns:
                # Handle unseen categories
                unique_values = processed_data[col].unique()
                for value in unique_values:
                    if value not in encoder.classes_:
                        # Add new category to encoder
                        encoder.classes_ = np.append(encoder.classes_, value)
                
                processed_data[col] = encoder.transform(processed_data[col])
        
        # Select only feature columns
        if isinstance(processed_data, pd.DataFrame):
            processed_data = processed_data[self.feature_columns]
        
        return processed_data
    
    def predict_single_student(self, student_data, model_name='LinearRegression'):
        """Predict performance for a single student"""
        if model_name not in self.models:
            raise ValueError(f"Model {model_name} not found. Available models: {list(self.models.keys())}")
        
        # Preprocess input
        processed_data = self.preprocess_input(student_data)
        
        # Scale features (except for Decision Tree)
        if model_name != 'DecisionTreeRegressor':
            processed_data_scaled = self.scaler.transform(processed_data)
            prediction = self.models[model_name].predict(processed_data_scaled)[0]
        else:
            prediction = self.models[model_name].predict(processed_data)[0]
        
        return round(prediction, 2)
    
    def predict_multiple_students(self, students_data, model_name='LinearRegression'):
        """Predict performance for multiple students"""
        if model_name not in self.models:
            raise ValueError(f"Model {model_name} not found. Available models: {list(self.models.keys())}")
        
        # Preprocess input
        processed_data = self.preprocess_input(students_data)
        
        # Scale features (except for Decision Tree)
        if model_name != 'DecisionTreeRegressor':
            processed_data_scaled = self.scaler.transform(processed_data)
            predictions = self.models[model_name].predict(processed_data_scaled)
        else:
            predictions = self.models[model_name].predict(processed_data)
        
        return [round(pred, 2) for pred in predictions]
    
    def predict_all_models(self, student_data):
        """Get predictions from all available models"""
        results = {}
        
        for model_name in self.models.keys():
            try:
                prediction = self.predict_single_student(student_data, model_name)
                results[model_name] = prediction
            except Exception as e:
                results[model_name] = f"Error: {str(e)}"
        
        return results
    
    def get_recommendation(self, prediction_score, study_time=None):
        """Generate recommendation based on prediction score"""
        if prediction_score < 50:
            return "Needs more study hours"
        elif prediction_score >= 80:
            return "Keep up the good work"
        else:
            return "Maintain consistency"
    
    def get_result(self, prediction_score):
        """Determine pass/fail result"""
        return "Pass" if prediction_score >= 40 else "Fail"

def create_sample_student():
    """Create a sample student for testing"""
    sample_student = pd.DataFrame({
        'Attendance (%)': [85.5],
        'Assignment (%)': [78.2],
        'Study Time (hrs)': [3.5],
        'Previous Score (%)': [72.8],
        'Extracurricular (%)': [25.0],
        'Department': ['Computer Science'],
        'Year': ['3rd'],
        'Status': ['Active']
    })
    return sample_student

def main():
    """Main function to demonstrate prediction functionality"""
    print(" Student Performance Prediction System")
    print("=" * 50)
    
    # Initialize predictor
    predictor = StudentPerformancePredictor()
    
    if not predictor.models:
        print(" No models loaded. Please train models first using train_ml_models.py")
        return
    
    # Create sample student data
    sample_student = create_sample_student()
    
    print("\n📊 Sample Student Data:")
    print(sample_student.to_string(index=False))
    
    # Get predictions from all models
    print("\n Predictions from All Models:")
    print("-" * 40)
    
    all_predictions = predictor.predict_all_models(sample_student)
    
    for model_name, prediction in all_predictions.items():
        if isinstance(prediction, (int, float)):
            recommendation = predictor.get_recommendation(prediction)
            result = predictor.get_result(prediction)
            print(f"{model_name}: {prediction:.2f}% - {result} - {recommendation}")
        else:
            print(f"{model_name}: {prediction}")
    
    # Interactive prediction
    print("\n" + "=" * 50)
    print(" Interactive Prediction")
    print("=" * 50)
    
    try:
        # Get user input
        print("\nEnter student information:")
        attendance = float(input("Attendance (%): "))
        assignment = float(input("Assignment (%): "))
        study_time = float(input("Study Time (hours): "))
        prev_score = float(input("Previous Score (%): "))
        extracurricular = float(input("Extracurricular (%): "))
        
        print("\nSelect Department:")
        departments = ['Computer Science', 'Mathematics', 'Physics', 'Chemistry', 
                      'Biology', 'Economics', 'Engineering', 'Literature', 'History']
        for i, dept in enumerate(departments, 1):
            print(f"{i}. {dept}")
        dept_choice = int(input("Enter choice (1-9): ")) - 1
        department = departments[dept_choice]
        
        print("\nSelect Year:")
        years = ['1st', '2nd', '3rd', '4th']
        for i, year in enumerate(years, 1):
            print(f"{i}. {year}")
        year_choice = int(input("Enter choice (1-4): ")) - 1
        year = years[year_choice]
        
        print("\nSelect Status:")
        statuses = ['Active', 'Inactive']
        for i, status in enumerate(statuses, 1):
            print(f"{i}. {status}")
        status_choice = int(input("Enter choice (1-2): ")) - 1
        status = statuses[status_choice]
        
        # Create student data
        student_data = pd.DataFrame({
            'Attendance (%)': [attendance],
            'Assignment (%)': [assignment],
            'Study Time (hrs)': [study_time],
            'Previous Score (%)': [prev_score],
            'Extracurricular (%)': [extracurricular],
            'Department': [department],
            'Year': [year],
            'Status': [status]
        })
        
        print(f"\n Student Data:")
        print(student_data.to_string(index=False))
        
        # Get predictions
        print(f"\n Performance Predictions:")
        print("-" * 40)
        
        all_predictions = predictor.predict_all_models(student_data)
        
        for model_name, prediction in all_predictions.items():
            if isinstance(prediction, (int, float)):
                recommendation = predictor.get_recommendation(prediction, study_time)
                result = predictor.get_result(prediction)
                print(f"{model_name}: {prediction:.2f}% - {result} - {recommendation}")
            else:
                print(f"{model_name}: {prediction}")
        
        # Show best model prediction
        valid_predictions = {k: v for k, v in all_predictions.items() 
                           if isinstance(v, (int, float))}
        if valid_predictions:
            best_model = max(valid_predictions.items(), key=lambda x: x[1])
            print(f"\n Best Prediction: {best_model[0]} - {best_model[1]:.2f}%")
        
    except (ValueError, IndexError) as e:
        print(f" Invalid input: {str(e)}")
    except KeyboardInterrupt:
        print("\n Goodbye!")

if __name__ == "__main__":
    main()
