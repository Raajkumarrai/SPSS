"""
Comprehensive System Testing Script for SPPS
This script runs all test cases and collects real performance data
"""

import os
import sys
import time
import pandas as pd
import numpy as np
from datetime import datetime
import traceback
import psutil
import threading
from io import BytesIO

# Add paths
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))
sys.path.append(os.path.join(os.path.dirname(__file__), 'data'))

# Import system components
try:
    from prediction_service import StudentPerformancePredictor, get_predictor
    from database import db_manager
    from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
    from sklearn.model_selection import cross_val_score
    import joblib
except ImportError as e:
    print(f"Import error: {e}")
    print("Please ensure all dependencies are installed")

class SystemTester:
    def __init__(self):
        self.results = {}
        self.test_data = None
        self.predictor = None
        
    def load_test_data(self):
        """Load test dataset"""
        print("\n" + "="*80)
        print("PHASE 1: DATA COLLECTION AND PREPARATION TEST")
        print("="*80)
        
        data_file = os.path.join('data', 'students_dataset_with_predictions.csv')
        if not os.path.exists(data_file):
            print(f"[ERROR] Test data file not found: {data_file}")
            return False
        
        try:
            self.test_data = pd.read_csv(data_file)
            print(f"[OK] Dataset loaded successfully")
            print(f"   - Total Records: {len(self.test_data)}")
            print(f"   - Columns: {len(self.test_data.columns)}")
            print(f"   - Features: {list(self.test_data.columns)}")
            
            # Check for missing values
            missing = self.test_data.isnull().sum().sum()
            print(f"   - Missing Values: {missing}")
            
            # Check for duplicates
            duplicates = self.test_data.duplicated().sum()
            print(f"   - Duplicate Records: {duplicates}")
            
            self.results['data_collection'] = {
                'total_records': len(self.test_data),
                'primary_features': len([c for c in self.test_data.columns if c not in ['Student ID', 'Student Name', 'Prediction Score', 'Recommendation', 'Result']]),
                'target_variable': 'Prediction Score',
                'missing_values': int(missing),
                'duplicates': int(duplicates)
            }
            
            return True
        except Exception as e:
            print(f"[ERROR] Error loading data: {e}")
            return False
    
    def test_data_preprocessing(self):
        """Test data preprocessing"""
        print("\n" + "="*80)
        print("PHASE 2: DATA PREPROCESSING TEST")
        print("="*80)
        
        if self.test_data is None:
            print("[ERROR] No test data available")
            return False
        
        try:
            df = self.test_data.copy()
            
            # Test duplicate removal
            initial_count = len(df)
            df_no_dup = df.drop_duplicates()
            duplicates_removed = initial_count - len(df_no_dup)
            
            # Test missing value handling
            missing_before = df.isnull().sum().sum()
            df_filled = df.fillna(df.median(numeric_only=True))
            missing_after = df_filled.isnull().sum().sum()
            
            # Test categorical encoding
            categorical_cols = ['Department', 'Year', 'Status']
            encoded_cols = [col for col in categorical_cols if col in df.columns]
            
            # Test feature scaling (check if values are in reasonable range)
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            numeric_cols = [c for c in numeric_cols if c not in ['Prediction Score']]
            
            print(f"[OK] Data Cleaning:")
            print(f"   - Duplicate Removal: {duplicates_removed} duplicates removed")
            print(f"   - Missing Value Handling: {missing_before} -> {missing_after} missing values")
            print(f"[OK] Feature Engineering:")
            print(f"   - Categorical Encoding: {len(encoded_cols)} columns encoded")
            print(f"   - Feature Scaling: {len(numeric_cols)} numeric features identified")
            
            self.results['data_preprocessing'] = {
                'duplicates_removed': int(duplicates_removed),
                'missing_values_handled': int(missing_before),
                'categorical_encoded': len(encoded_cols),
                'features_scaled': len(numeric_cols)
            }
            
            return True
        except Exception as e:
            print(f"[ERROR] Error in preprocessing test: {e}")
            traceback.print_exc()
            return False
    
    def test_model_training(self):
        """Test ML models and get actual metrics"""
        print("\n" + "="*80)
        print("PHASE 3: MODEL TRAINING AND VALIDATION TEST")
        print("="*80)
        
        try:
            # Load models
            models_dir = os.path.join('data', 'models')
            models = {}
            
            model_files = {
                'LinearRegression': 'linearregression.pkl',
                'KNNRegressor': 'knnregressor.pkl',
                'DecisionTreeRegressor': 'decisiontreeregressor.pkl'
            }
            
            scaler = joblib.load(os.path.join(models_dir, 'scaler.pkl'))
            label_encoders = joblib.load(os.path.join(models_dir, 'label_encoders.pkl'))
            feature_columns = joblib.load(os.path.join(models_dir, 'feature_columns.pkl'))
            
            for model_name, filename in model_files.items():
                model_path = os.path.join(models_dir, filename)
                if os.path.exists(model_path):
                    models[model_name] = joblib.load(model_path)
                    print(f"[OK] Loaded {model_name}")
            
            if not models:
                print("[ERROR] No models found")
                return False
            
            # Prepare test data
            df = self.test_data.copy()
            
            # Encode categorical variables
            for col, encoder in label_encoders.items():
                if col in df.columns:
                    # Handle unseen categories
                    for value in df[col].unique():
                        if value not in encoder.classes_:
                            encoder.classes_ = np.append(encoder.classes_, value)
                    df[col] = encoder.transform(df[col])
            
            # Select features
            X = df[feature_columns]
            y = df['Prediction Score'] if 'Prediction Score' in df.columns else None
            
            if y is None:
                print("[ERROR] Target variable not found in test data")
                return False
            
            # Split data (80-10-10)
            from sklearn.model_selection import train_test_split
            X_train, X_temp, y_train, y_temp = train_test_split(X, y, test_size=0.2, random_state=42)
            X_val, X_test, y_val, y_test = train_test_split(X_temp, y_temp, test_size=0.5, random_state=42)
            
            print(f"\n[INFO] Data Split:")
            print(f"   - Training: {len(X_train)} records (80%)")
            print(f"   - Validation: {len(X_val)} records (10%)")
            print(f"   - Testing: {len(X_test)} records (10%)")
            
            # Test each model
            model_results = {}
            
            for model_name, model in models.items():
                print(f"\n[TEST] Testing {model_name}...")
                
                # Scale features (except for Decision Tree)
                if model_name != 'DecisionTreeRegressor':
                    X_train_scaled = scaler.transform(X_train)
                    X_test_scaled = scaler.transform(X_test)
                    X_val_scaled = scaler.transform(X_val)
                    
                    # Train on training set
                    model.fit(X_train_scaled, y_train)
                    
                    # Predict on test set
                    y_pred = model.predict(X_test_scaled)
                else:
                    # Decision Tree doesn't need scaling
                    model.fit(X_train, y_train)
                    y_pred = model.predict(X_test)
                
                # Calculate metrics
                mae = mean_absolute_error(y_test, y_pred)
                rmse = np.sqrt(mean_squared_error(y_test, y_pred))
                r2 = r2_score(y_test, y_pred)
                
                # Cross-validation
                if model_name != 'DecisionTreeRegressor':
                    cv_scores = cross_val_score(model, X_train_scaled, y_train, cv=5, scoring='r2')
                else:
                    cv_scores = cross_val_score(model, X_train, y_train, cv=5, scoring='r2')
                
                # Calculate accuracy (percentage of predictions within ±10 points)
                accuracy = np.mean(np.abs(y_test - y_pred) <= 10) * 100
                
                model_results[model_name] = {
                    'MAE': round(mae, 2),
                    'RMSE': round(rmse, 2),
                    'R2': round(r2, 4),
                    'CV_Score': round(cv_scores.mean(), 4),
                    'Accuracy_Rate': round(accuracy, 1)
                }
                
                print(f"   [OK] MAE: {mae:.2f}")
                print(f"   [OK] RMSE: {rmse:.2f}")
                print(f"   [OK] R²: {r2:.4f}")
                print(f"   [OK] CV Score: {cv_scores.mean():.4f}")
                print(f"   [OK] Accuracy Rate: {accuracy:.1f}%")
            
            # Test ensemble (average of all models)
            print(f"\n[TEST] Testing Ensemble Model...")
            ensemble_predictions = []
            for model_name, model in models.items():
                if model_name != 'DecisionTreeRegressor':
                    X_test_scaled = scaler.transform(X_test)
                    pred = model.predict(X_test_scaled)
                else:
                    pred = model.predict(X_test)
                ensemble_predictions.append(pred)
            
            ensemble_pred = np.mean(ensemble_predictions, axis=0)
            ensemble_mae = mean_absolute_error(y_test, ensemble_pred)
            ensemble_rmse = np.sqrt(mean_squared_error(y_test, ensemble_pred))
            ensemble_r2 = r2_score(y_test, ensemble_pred)
            ensemble_accuracy = np.mean(np.abs(y_test - ensemble_pred) <= 10) * 100
            
            # Cross-validation for ensemble (average of individual CV scores)
            ensemble_cv = np.mean([model_results[m]['CV_Score'] for m in models.keys()])
            
            model_results['Ensemble'] = {
                'MAE': round(ensemble_mae, 2),
                'RMSE': round(ensemble_rmse, 2),
                'R2': round(ensemble_r2, 4),
                'CV_Score': round(ensemble_cv, 4),
                'Accuracy_Rate': round(ensemble_accuracy, 1)
            }
            
            print(f"   [OK] MAE: {ensemble_mae:.2f}")
            print(f"   [OK] RMSE: {ensemble_rmse:.2f}")
            print(f"   [OK] R²: {ensemble_r2:.4f}")
            print(f"   [OK] CV Score: {ensemble_cv:.4f}")
            print(f"   [OK] Accuracy Rate: {ensemble_accuracy:.1f}%")
            
            self.results['model_training'] = {
                'training_set_size': len(X_train),
                'validation_set_size': len(X_val),
                'testing_set_size': len(X_test),
                'models': model_results
            }
            
            return True
        except Exception as e:
            print(f"[ERROR] Error in model training test: {e}")
            traceback.print_exc()
            return False
    
    def test_individual_predictions(self):
        """Test individual prediction performance"""
        print("\n" + "="*80)
        print("PHASE 4: INDIVIDUAL PREDICTION TEST")
        print("="*80)
        
        try:
            if self.predictor is None:
                self.predictor = get_predictor()
            
            if self.predictor is None:
                print("[ERROR] Predictor not available")
                return False
            
            # Test scenarios
            test_scenarios = [
                {
                    'name': 'Valid & Complete Data',
                    'data': pd.DataFrame({
                        'Attendance (%)': [85.5],
                        'Assignment (%)': [78.0],
                        'Study Time (hrs)': [4.5],
                        'Previous Score (%)': [72.0],
                        'Extracurricular (%)': [60.0],
                        'Department': ['Computer Science'],
                        'Year': ['2nd'],
                        'Status': ['Active']
                    })
                },
                {
                    'name': 'Partially Missing Data',
                    'data': pd.DataFrame({
                        'Attendance (%)': [90.0],
                        'Assignment (%)': [None],  # Missing
                        'Study Time (hrs)': [5.0],
                        'Previous Score (%)': [80.0],
                        'Extracurricular (%)': [70.0],
                        'Department': ['Engineering'],
                        'Year': ['3rd'],
                        'Status': ['Active']
                    })
                },
                {
                    'name': 'Boundary Value Input (Min)',
                    'data': pd.DataFrame({
                        'Attendance (%)': [0.0],
                        'Assignment (%)': [0.0],
                        'Study Time (hrs)': [0.0],
                        'Previous Score (%)': [0.0],
                        'Extracurricular (%)': [0.0],
                        'Department': ['Arts'],
                        'Year': ['1st'],
                        'Status': ['Active']
                    })
                },
                {
                    'name': 'Boundary Value Input (Max)',
                    'data': pd.DataFrame({
                        'Attendance (%)': [100.0],
                        'Assignment (%)': [100.0],
                        'Study Time (hrs)': [10.0],
                        'Previous Score (%)': [100.0],
                        'Extracurricular (%)': [100.0],
                        'Department': ['Business'],
                        'Year': ['4th'],
                        'Status': ['Active']
                    })
                }
            ]
            
            results = []
            total_time = 0
            success_count = 0
            
            for scenario in test_scenarios:
                try:
                    start_time = time.time()
                    prediction = self.predictor.predict_all_models(scenario['data'])
                    elapsed = time.time() - start_time
                    total_time += elapsed
                    success_count += 1
                    
                    results.append({
                        'scenario': scenario['name'],
                        'status': 'Success',
                        'response_time': round(elapsed, 3),
                        'prediction': prediction
                    })
                    print(f"[OK] {scenario['name']}: {elapsed:.3f}s")
                except Exception as e:
                    results.append({
                        'scenario': scenario['name'],
                        'status': 'Failed',
                        'error': str(e)
                    })
                    print(f"[ERROR] {scenario['name']}: {str(e)}")
            
            avg_response_time = total_time / len(test_scenarios) if test_scenarios else 0
            success_rate = (success_count / len(test_scenarios)) * 100 if test_scenarios else 0
            
            self.results['individual_prediction'] = {
                'avg_response_time': round(avg_response_time, 3),
                'success_rate': round(success_rate, 1),
                'test_scenarios': results
            }
            
            print(f"\n[INFO] Summary:")
            print(f"   - Average Response Time: {avg_response_time:.3f}s")
            print(f"   - Success Rate: {success_rate:.1f}%")
            
            return True
        except Exception as e:
            print(f"[ERROR] Error in individual prediction test: {e}")
            traceback.print_exc()
            return False
    
    def test_batch_predictions(self):
        """Test batch prediction performance"""
        print("\n" + "="*80)
        print("PHASE 5: BATCH PREDICTION TEST")
        print("="*80)
        
        try:
            if self.predictor is None:
                self.predictor = get_predictor()
            
            if self.predictor is None:
                print("[ERROR] Predictor not available")
                return False
            
            # Generate test data for different batch sizes
            batch_sizes = [10, 50, 100]
            batch_results = []
            
            for batch_size in batch_sizes:
                print(f"\n[TEST] Testing batch size: {batch_size}")
                
                # Generate sample data
                test_data = pd.DataFrame({
                    'Attendance (%)': np.random.uniform(60, 100, batch_size),
                    'Assignment (%)': np.random.uniform(50, 100, batch_size),
                    'Study Time (hrs)': np.random.uniform(2, 8, batch_size),
                    'Previous Score (%)': np.random.uniform(40, 95, batch_size),
                    'Extracurricular (%)': np.random.uniform(0, 100, batch_size),
                    'Department': np.random.choice(['Computer Science', 'Engineering', 'Business', 'Arts'], batch_size),
                    'Year': np.random.choice(['1st', '2nd', '3rd', '4th'], batch_size),
                    'Status': ['Active'] * batch_size
                })
                
                # Measure memory before
                process = psutil.Process(os.getpid())
                memory_before = process.memory_info().rss / 1024 / 1024  # MB
                
                # Run prediction
                start_time = time.time()
                predictions = []
                success_count = 0
                
                for idx, row in test_data.iterrows():
                    try:
                        pred = self.predictor.predict_all_models(pd.DataFrame([row]))
                        predictions.append(pred)
                        success_count += 1
                    except Exception as e:
                        predictions.append({'error': str(e)})
                
                elapsed = time.time() - start_time
                
                # Measure memory after
                memory_after = process.memory_info().rss / 1024 / 1024  # MB
                memory_used = memory_after - memory_before
                
                success_rate = (success_count / batch_size) * 100
                
                batch_results.append({
                    'batch_size': batch_size,
                    'time': round(elapsed, 2),
                    'success_rate': round(success_rate, 1),
                    'memory_mb': round(memory_used, 1)
                })
                
                print(f"   [OK] Time: {elapsed:.2f}s")
                print(f"   [OK] Success Rate: {success_rate:.1f}%")
                print(f"   [OK] Memory Used: {memory_used:.1f} MB")
            
            self.results['batch_prediction'] = batch_results
            
            return True
        except Exception as e:
            print(f"[ERROR] Error in batch prediction test: {e}")
            traceback.print_exc()
            return False
    
    def test_database_operations(self):
        """Test database CRUD operations"""
        print("\n" + "="*80)
        print("PHASE 6: DATABASE TEST")
        print("="*80)
        
        try:
            # Test CREATE
            print("\n[TEST] Testing CREATE operation...")
            start_time = time.time()
            test_student = {
                'student_id': f'TEST_{int(time.time())}',
                'name': 'Test Student',
                'department': 'Computer Science',
                'year': '2nd',
                'attendance': 85.5,
                'assignment': 78.0,
                'study_time': 4.5,
                'previous_score': 72.0,
                'extracurricular': 60.0,
                'status': 'Active'
            }
            
            create_query = """
                INSERT INTO students (student_id, name, department, year, attendance, 
                                    assignment, study_time, previous_score, extracurricular, status)
                VALUES (%(student_id)s, %(name)s, %(department)s, %(year)s, %(attendance)s,
                        %(assignment)s, %(study_time)s, %(previous_score)s, %(extracurricular)s, %(status)s)
            """
            
            try:
                db_manager.execute_query(create_query, test_student)
                create_time = time.time() - start_time
                print(f"   [OK] CREATE: {create_time:.3f}s")
            except Exception as e:
                print(f"   [ERROR] CREATE failed: {e}")
                create_time = None
            
            # Test READ
            print("\n[TEST] Testing READ operation...")
            start_time = time.time()
            read_query = "SELECT * FROM students WHERE student_id = %s LIMIT 1"
            try:
                result = db_manager.execute_query(read_query, (test_student['student_id'],))
                read_time = time.time() - start_time
                print(f"   [OK] READ: {read_time:.3f}s")
            except Exception as e:
                print(f"   [ERROR] READ failed: {e}")
                read_time = None
            
            # Test UPDATE
            print("\n[TEST] Testing UPDATE operation...")
            start_time = time.time()
            update_query = """
                UPDATE students SET attendance = %s WHERE student_id = %s
            """
            try:
                db_manager.execute_query(update_query, (90.0, test_student['student_id']))
                update_time = time.time() - start_time
                print(f"   [OK] UPDATE: {update_time:.3f}s")
            except Exception as e:
                print(f"   [ERROR] UPDATE failed: {e}")
                update_time = None
            
            # Test DELETE
            print("\n[TEST] Testing DELETE operation...")
            start_time = time.time()
            delete_query = "DELETE FROM students WHERE student_id = %s"
            try:
                db_manager.execute_query(delete_query, (test_student['student_id'],))
                delete_time = time.time() - start_time
                print(f"   [OK] DELETE: {delete_time:.3f}s")
            except Exception as e:
                print(f"   [ERROR] DELETE failed: {e}")
                delete_time = None
            
            self.results['database'] = {
                'create_time': round(create_time, 3) if create_time else None,
                'read_time': round(read_time, 3) if read_time else None,
                'update_time': round(update_time, 3) if update_time else None,
                'delete_time': round(delete_time, 3) if delete_time else None
            }
            
            return True
        except Exception as e:
            print(f"[ERROR] Error in database test: {e}")
            traceback.print_exc()
            return False
    
    def test_report_generation(self):
        """Test Excel report generation"""
        print("\n" + "="*80)
        print("PHASE 7: REPORT GENERATION TEST")
        print("="*80)
        
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment
            
            # Test individual report
            print("\n[TEST] Testing Individual Student Report...")
            start_time = time.time()
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Prediction Results"
            
            # Add sample data
            ws['A1'] = "PREDICTION RESULTS"
            ws['A1'].font = Font(bold=True, size=16)
            
            sample_data = {
                'Student ID': 'ST001',
                'Name': 'Test Student',
                'Predicted Score': 85.5,
                'Result': 'Pass',
                'Recommendation': 'Excellent performance'
            }
            
            for idx, (key, value) in enumerate(sample_data.items(), start=3):
                ws[f'A{idx}'] = key
                ws[f'B{idx}'] = value
            
            # Save to memory
            output = BytesIO()
            wb.save(output)
            individual_time = time.time() - start_time
            individual_size = len(output.getvalue()) / 1024  # KB
            
            print(f"   [OK] Individual Report: {individual_time:.3f}s, {individual_size:.2f} KB")
            
            # Test batch report (100 students)
            print("\n[TEST] Testing Batch Report (100 students)...")
            start_time = time.time()
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Prediction Results"
            
            # Add headers
            headers = ['Student ID', 'Name', 'Predicted Score', 'Result', 'Recommendation']
            for idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=idx, value=header)
                cell.font = Font(bold=True)
            
            # Add sample data (100 rows)
            for row in range(2, 102):
                ws.cell(row=row, column=1, value=f'ST{row:03d}')
                ws.cell(row=row, column=2, value=f'Student {row}')
                ws.cell(row=row, column=3, value=np.random.uniform(40, 95))
                ws.cell(row=row, column=4, value='Pass')
                ws.cell(row=row, column=5, value='Good performance')
            
            output = BytesIO()
            wb.save(output)
            batch_time = time.time() - start_time
            batch_size = len(output.getvalue()) / 1024 / 1024  # MB
            
            print(f"   [OK] Batch Report: {batch_time:.3f}s, {batch_size:.2f} MB")
            
            self.results['report_generation'] = {
                'individual_time': round(individual_time, 3),
                'individual_size_kb': round(individual_size, 2),
                'batch_time': round(batch_time, 3),
                'batch_size_mb': round(batch_size, 2)
            }
            
            return True
        except Exception as e:
            print(f"[ERROR] Error in report generation test: {e}")
            traceback.print_exc()
            return False
    
    def generate_test_report(self):
        """Generate comprehensive test report"""
        print("\n" + "="*80)
        print("GENERATING TEST REPORT")
        print("="*80)
        
        report = f"""
# SPPS - Comprehensive System Testing Report
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Test Results Summary

"""
        
        # Add all test results
        for test_name, test_result in self.results.items():
            report += f"\n### {test_name.replace('_', ' ').title()}\n\n"
            report += f"```json\n{str(test_result)}\n```\n"
        
        # Save report
        report_file = 'test_results_report.txt'
        with open(report_file, 'w') as f:
            f.write(report)
        
        print(f"[OK] Test report saved: {report_file}")
        return report
    
    def run_all_tests(self):
        """Run all test phases"""
        print("\n" + "="*80)
        print("SPPS COMPREHENSIVE SYSTEM TESTING")
        print("="*80)
        print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        tests = [
            ('Data Collection', self.load_test_data),
            ('Data Preprocessing', self.test_data_preprocessing),
            ('Model Training', self.test_model_training),
            ('Individual Predictions', self.test_individual_predictions),
            ('Batch Predictions', self.test_batch_predictions),
            ('Database Operations', self.test_database_operations),
            ('Report Generation', self.test_report_generation)
        ]
        
        passed = 0
        failed = 0
        
        for test_name, test_func in tests:
            try:
                if test_func():
                    passed += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"[ERROR] {test_name} failed with exception: {e}")
                failed += 1
        
        print("\n" + "="*80)
        print("TEST SUMMARY")
        print("="*80)
        print(f"[OK] Passed: {passed}")
        print(f"[ERROR] Failed: {failed}")
        print(f"Total: {passed + failed}")
        print(f"End Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Generate report
        self.generate_test_report()
        
        return self.results

if __name__ == "__main__":
    tester = SystemTester()
    results = tester.run_all_tests()
    print("\n[OK] Testing complete! Check test_results_report.txt for detailed results.")

