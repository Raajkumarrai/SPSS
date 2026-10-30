import os
import sys
import pandas as pd
import numpy as np
import joblib
from datetime import datetime

# Add the data directory to the path to import the trained models
data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
sys.path.append(data_dir)

class StudentPerformancePredictor:
    """Service class for making student performance predictions using trained ML models"""
    
    def __init__(self):
        """Initialize the predictor with trained models"""
        self.models_dir = os.path.join(data_dir, 'models')
        self.models = {}
        self.scaler = None
        self.label_encoders = {}
        self.feature_columns = []
        
        # Load models and preprocessing objects
        self.load_models()
    
    def load_models(self):
        """Load trained models and preprocessing objects"""
        try:
            # Load preprocessing objects
            self.scaler = joblib.load(os.path.join(self.models_dir, 'scaler.pkl'))
            self.label_encoders = joblib.load(os.path.join(self.models_dir, 'label_encoders.pkl'))
            self.feature_columns = joblib.load(os.path.join(self.models_dir, 'feature_columns.pkl'))
            
            # Load models
            model_files = {
                'LinearRegression': 'linearregression.pkl',
                'KNNRegressor': 'knnregressor.pkl',
                'DecisionTreeRegressor': 'decisiontreeregressor.pkl'
            }
            
            for model_name, filename in model_files.items():
                model_path = os.path.join(self.models_dir, filename)
                if os.path.exists(model_path):
                    self.models[model_name] = joblib.load(model_path)
                else:
                    print(f"Warning: {model_name} not found at {model_path}")
            
            print(f"Successfully loaded {len(self.models)} models")
            
        except Exception as e:
            print(f"Error loading models: {str(e)}")
            print("Please ensure models are trained first using train_ml_models.py")
    
    def preprocess_input(self, student_data):
        """Preprocess input data for prediction"""
        processed_data = student_data.copy()
        
        for col, encoder in self.label_encoders.items():
            if col in processed_data.columns:
                unique_values = processed_data[col].unique()
                for value in unique_values:
                    if value not in encoder.classes_:
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
        
        # Limit prediction to realistic bounds (0-97%)
        prediction = max(0.0, min(97.0, prediction))
        
        return round(prediction, 2)
    
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
        if prediction_score < 40:
            return "Critical intervention needed - consider academic counseling"
        elif prediction_score < 50:
            return "Needs more study hours and academic support"
        elif prediction_score < 60:
            return "Below average - focus on improving attendance and assignments"
        elif prediction_score < 70:
            return "Average performance - maintain consistency and improve study habits"
        elif prediction_score < 80:
            return "Good performance - continue current study patterns"
        elif prediction_score < 90:
            return "Excellent performance - keep up the good work"
        else:
            return "Outstanding performance - maintain high standards"
    
    def get_result(self, prediction_score):
        """Determine pass/fail result"""
        return "Pass" if prediction_score >= 40 else "Fail"
    
    def predict_from_database_student(self, student_record):
        """Predict performance for a student from database record"""
        try:
            # Create student data DataFrame
            student_data = pd.DataFrame({
                'Attendance (%)': [float(student_record.get('attendance', 0))],
                'Assignment (%)': [float(student_record.get('assignment', 0))],
                'Study Time (hrs)': [float(student_record.get('study_time', 0))],
                'Previous Score (%)': [float(student_record.get('previous_score', 0))],
                'Extracurricular (%)': [float(student_record.get('extracurricular', 0))],
                'Department': [student_record.get('department', 'Unknown')],
                'Year': [student_record.get('year', '1st')],
                'Status': [student_record.get('status', 'Active')]
            })
            
            # Get predictions from all models
            predictions = self.predict_all_models(student_data)
            
            # Get the best prediction (highest score)
            valid_predictions = {k: v for k, v in predictions.items() 
                               if isinstance(v, (int, float))}
            
            if valid_predictions:
                best_model = max(valid_predictions.items(), key=lambda x: x[1])
                best_prediction = best_model[1]
                best_model_name = best_model[0]
            else:
                best_prediction = 0
                best_model_name = "No valid predictions"
            
            # Generate recommendation and result
            recommendation = self.get_recommendation(best_prediction)
            result = self.get_result(best_prediction)
            
            return {
                'predictions': predictions,
                'best_prediction': best_prediction,
                'best_model': best_model_name,
                'recommendation': recommendation,
                'result': result,
                'student_data': student_data.to_dict('records')[0]
            }
            
        except Exception as e:
            print(f"Error predicting from database student: {str(e)}")
            return {
                'error': str(e),
                'predictions': {},
                'best_prediction': 0,
                'best_model': "Error",
                'recommendation': "Unable to generate recommendation",
                'result': "Error",
                'student_data': {}
            }

# Global predictor instance
predictor = None

def get_predictor():
    """Get or create the global predictor instance"""
    global predictor
    if predictor is None:
        try:
            predictor = StudentPerformancePredictor()
        except Exception as e:
            print(f"Error initializing predictor: {e}")
            return None
    return predictor

def make_prediction(student_data):
    """Make a prediction for student data"""
    pred = get_predictor()
    if pred is None:
        return {
            'error': 'Prediction service not available',
            'predictions': {},
            'best_prediction': 0,
            'best_model': "Error",
            'recommendation': "Unable to generate recommendation",
            'result': "Error"
        }
    
    try:
        # Convert student_data to DataFrame if it's not already
        if not isinstance(student_data, pd.DataFrame):
            student_data = pd.DataFrame([student_data])
        
        # Get predictions from all models
        predictions = pred.predict_all_models(student_data)
        
        # Get the best prediction (highest score)
        valid_predictions = {k: v for k, v in predictions.items() 
                           if isinstance(v, (int, float))}
        
        if valid_predictions:
            best_model = max(valid_predictions.items(), key=lambda x: x[1])
            best_prediction = best_model[1]
            best_model_name = best_model[0]
        else:
            best_prediction = 0
            best_model_name = "No valid predictions"
        
        # Generate recommendation and result
        recommendation = pred.get_recommendation(best_prediction)
        result = pred.get_result(best_prediction)
        
        return {
            'predictions': predictions,
            'best_prediction': best_prediction,
            'best_model': best_model_name,
            'recommendation': recommendation,
            'result': result,
            'timestamp': datetime.now().isoformat()
        }
        
    except Exception as e:
        print(f"Error making prediction: {e}")
        return {
            'error': str(e),
            'predictions': {},
            'best_prediction': 0,
            'best_model': "Error",
            'recommendation': "Unable to generate recommendation",
            'result': "Error"
        }
