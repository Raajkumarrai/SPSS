import pandas as pd
import numpy as np
from predict_performance import StudentPerformancePredictor

def demo_realistic_dataset():
    """Demonstrate the realistic dataset with 42.8% failure rate"""
    
    print(" Realistic Dataset Demonstration - 42.8% Failure Rate")
    print("=" * 60)
    
    # Load the dataset
    df = pd.read_csv("students_dataset_with_predictions.csv")
    
    print(f"\n Dataset Overview:")
    print(f"   Total Students: {len(df):,}")
    print(f"   Pass Rate: {(df['Result'] == 'Pass').mean() * 100:.1f}%")
    print(f"   Failure Rate: {(df['Result'] == 'Fail').mean() * 100:.1f}%")
    print(f"   Average Prediction Score: {df['Prediction Score'].mean():.2f}%")
    
    print(f"\n Academic Metrics:")
    print(f"   Average Attendance: {df['Attendance (%)'].mean():.2f}%")
    print(f"   Average Assignment Score: {df['Assignment (%)'].mean():.2f}%")
    print(f"   Average Study Time: {df['Study Time (hrs)'].mean():.2f} hours")
    print(f"   Average Previous Score: {df['Previous Score (%)'].mean():.2f}%")
    
    print(f"\n Recommendation Distribution:")
    recommendation_counts = df['Recommendation'].value_counts()
    for rec, count in recommendation_counts.items():
        percentage = (count / len(df)) * 100
        print(f"   {rec}: {count:,} students ({percentage:.1f}%)")
    
    # Show some failing students
    failing_students = df[df['Result'] == 'Fail'].head(5)
    print(f"\n Sample Failing Students (42.8% of dataset):")
    print("-" * 50)
    for idx, student in failing_students.iterrows():
        print(f"   {student['Student Name']} - Score: {student['Prediction Score']:.1f}% - {student['Recommendation']}")
    
    # Show some passing students
    passing_students = df[df['Result'] == 'Pass'].head(5)
    print(f"\n Sample Passing Students (57.2% of dataset):")
    print("-" * 50)
    for idx, student in passing_students.iterrows():
        print(f"   {student['Student Name']} - Score: {student['Prediction Score']:.1f}% - {student['Recommendation']}")
    
    # Test prediction system
    print(f"\n Testing ML Prediction System:")
    print("-" * 50)
    
    try:
        predictor = StudentPerformancePredictor()
        
        # Test with a failing student profile
        failing_profile = pd.DataFrame({
            'Attendance (%)': [25.0],
            'Assignment (%)': [30.0],
            'Study Time (hrs)': [1.0],
            'Previous Score (%)': [20.0],
            'Extracurricular (%)': [5.0],
            'Department': ['Computer Science'],
            'Year': ['2nd'],
            'Status': ['Active']
        })
        
        print(f" Testing Failing Student Profile:")
        print(f"   Attendance: 25%, Assignment: 30%, Study Time: 1hr")
        print(f"   Previous Score: 20%, Extracurricular: 5%")
        
        predictions = predictor.predict_all_models(failing_profile)
        for model, pred in predictions.items():
            if isinstance(pred, (int, float)):
                result = predictor.get_result(pred)
                recommendation = predictor.get_recommendation(pred)
                print(f"   {model}: {pred:.1f}% - {result} - {recommendation}")
        
        # Test with a passing student profile
        passing_profile = pd.DataFrame({
            'Attendance (%)': [70.0],
            'Assignment (%)': [65.0],
            'Study Time (hrs)': [3.0],
            'Previous Score (%)': [55.0],
            'Extracurricular (%)': [15.0],
            'Department': ['Mathematics'],
            'Year': ['3rd'],
            'Status': ['Active']
        })
        
        print(f"\n Testing Passing Student Profile:")
        print(f"   Attendance: 70%, Assignment: 65%, Study Time: 3hr")
        print(f"   Previous Score: 55%, Extracurricular: 15%")
        
        predictions = predictor.predict_all_models(passing_profile)
        for model, pred in predictions.items():
            if isinstance(pred, (int, float)):
                result = predictor.get_result(pred)
                recommendation = predictor.get_recommendation(pred)
                print(f"   {model}: {pred:.1f}% - {result} - {recommendation}")
        
    except Exception as e:
        print(f"     Prediction system not available: {str(e)}")
    
    print(f"\n Key Insights:")
    print(f"   • 42.8% of students fail to meet the 40% passing threshold")
    print(f"   • 97% of students need additional study support")
    print(f"   • Average study time is only 1.66 hours per day")
    print(f"   • Low attendance (44.22%) indicates systemic issues")
    print(f"   • Poor assignment performance (37.58%) suggests curriculum challenges")
    
    print(f"\n This realistic dataset represents:")
    print(f"   • Challenging academic environment")
    print(f"   • Students struggling with multiple factors")
    print(f"   • Need for comprehensive intervention programs")
    print(f"   • Realistic testing of ML prediction systems")

if __name__ == "__main__":
    demo_realistic_dataset()
