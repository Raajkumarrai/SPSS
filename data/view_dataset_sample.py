import pandas as pd

# Load the dataset
df = pd.read_csv("students_dataset_with_predictions.csv")

print(" Student Performance Dataset Sample")
print("=" * 80)
print(f"Total records: {len(df):,}")
print(f"Columns: {list(df.columns)}")
print("\n" + "=" * 80)

# Display first 5 records with better formatting
print("\n First 5 Records:")
print(df.head().to_string(index=False, max_colwidth=20))

print("\n" + "=" * 80)

# Display basic statistics
print("\n Dataset Statistics:")
print(f"Average Attendance: {df['Attendance (%)'].mean():.2f}%")
print(f"Average Assignment Score: {df['Assignment (%)'].mean():.2f}%")
print(f"Average Study Time: {df['Study Time (hrs)'].mean():.2f} hours")
print(f"Average Previous Score: {df['Previous Score (%)'].mean():.2f}%")
print(f"Average Prediction Score: {df['Prediction Score'].mean():.2f}%")
print(f"Pass Rate: {(df['Result'] == 'Pass').mean() * 100:.1f}%")

print("\n" + "=" * 80)

# Department distribution
print("\n Department Distribution:")
dept_counts = df['Department'].value_counts()
for dept, count in dept_counts.items():
    print(f"  {dept}: {count:,} students")

print("\n" + "=" * 80)

# Year distribution
print("\n Year Distribution:")
year_counts = df['Year'].value_counts()
for year, count in year_counts.items():
    print(f"  {year} Year: {count:,} students")

print("\n" + "=" * 80)

# Recommendation distribution
print("\n Recommendation Distribution:")
rec_counts = df['Recommendation'].value_counts()
for rec, count in rec_counts.items():
    print(f"  {rec}: {count:,} students")

print("\n" + "=" * 80)

# Status distribution
print("\n Status Distribution:")
status_counts = df['Status'].value_counts()
for status, count in status_counts.items():
    print(f"  {status}: {count:,} students")

print("\n" + "=" * 80)

# Show some sample records with key information
print("\n Sample Student Records:")
sample_df = df[['Student ID', 'Student Name', 'Department', 'Year', 'Prediction Score', 'Result', 'Recommendation']].head(10)
print(sample_df.to_string(index=False))
