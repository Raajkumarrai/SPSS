import pandas as pd
import random
import faker
import os

# Initialize Faker
fake = faker.Faker()

# Parameters
num_records = 50000

# Possible values
departments = [
    "Computer Science", "Mathematics", "Physics", "Chemistry", 
    "Biology", "Economics", "Engineering", "Literature", "History"
]
years = ["1st", "2nd", "3rd", "4th"]
statuses = ["Active", "Inactive"]

# Generate dataset
data = []
for i in range(1, num_records + 1):
    student_id = f"STU{i:05d}"
    student_name = fake.name()
    
    # More realistic and challenging score ranges with bias towards lower scores
    # Use triangular distribution to bias towards lower scores
    attendance = round(random.triangular(8, 75, 50), 2)  # Even more biased towards lower attendance
    assignment = round(random.triangular(3, 70, 40), 2)  # Even more biased towards lower assignments
    study_time = round(random.triangular(0.5, 3.5, 1), 2)  # Even more biased towards lower study time
    prev_score = round(random.triangular(3, 65, 35), 2)  # Even more biased towards lower previous scores
    extracurricular = round(random.triangular(0, 25, 8), 2)  # Even more biased towards lower extracurricular
    
    department = random.choice(departments)
    year = random.choice(years)
    status = random.choice(statuses)
    
    data.append([
        student_id, student_name, attendance, assignment,
        study_time, prev_score, extracurricular, department,
        year, status
    ])

# Create DataFrame
df = pd.DataFrame(data, columns=[
    "Student ID", "Student Name", "Attendance (%)", "Assignment (%)",
    "Study Time (hrs)", "Previous Score (%)", "Extracurricular (%)",
    "Department", "Year", "Status"
])

# Add Prediction Score with more challenging formula
df["Prediction Score"] = round(
    (df["Attendance (%)"] * 0.35) +
    (df["Assignment (%)"] * 0.45) +
    (df["Previous Score (%)"] * 0.15) +
    (df["Study Time (hrs)"] * 2) +  # Study time has less impact
    (df["Extracurricular (%)"] * 0.05), 2
)

# Add Recommendation
def get_recommendation(study_time, prediction_score):
    if study_time < 2.5 or prediction_score < 50:
        return "Needs more study hours"
    elif prediction_score >= 80:
        return "Keep up the good work"
    else:
        return "Maintain consistency"

df["Recommendation"] = df.apply(
    lambda x: get_recommendation(x["Study Time (hrs)"], x["Prediction Score"]),
    axis=1
)

# Add Result
df["Result"] = df["Prediction Score"].apply(lambda x: "Pass" if x >= 40 else "Fail")

# Save files
df.to_csv("students_dataset_with_predictions.csv", index=False)
try:
    df.to_excel("students_dataset_with_predictions.xlsx", index=False)
    print(" Excel file saved successfully")
except PermissionError:
    print("  Excel file could not be saved (file may be open)")
    print(" CSV file saved successfully")

print(" Dataset generated and saved as CSV and Excel")
print(f" Total records generated: {len(df)}")
print(f" Files saved in: {os.getcwd()}")
print("\n Dataset Statistics:")
print(f"   - Average Prediction Score: {df['Prediction Score'].mean():.2f}")
print(f"   - Pass Rate: {(df['Result'] == 'Pass').mean() * 100:.1f}%")
print(f"   - Departments: {len(df['Department'].unique())}")
print(f"   - Years: {len(df['Year'].unique())}")
