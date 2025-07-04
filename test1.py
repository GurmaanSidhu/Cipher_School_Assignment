import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

url = 'https://docs.google.com/spreadsheets/d/1wrH9VMn3BB1t8Y5F7bmtC0z0-8QCRFRf/export?format=xlsx'
df = pd.read_excel(url)

df['JoinDate'] = pd.to_datetime(df['JoinDate'], dayfirst=True, errors='coerce')

df['Salary'] = pd.to_numeric(df['Salary'], errors='coerce')
df['PerformanceRating'] = pd.to_numeric(df['PerformanceRating'], errors='coerce')

df['Salary'] = df['Salary'].fillna(df['Salary'].median())
df['PerformanceRating'] = df['PerformanceRating'].fillna(df['PerformanceRating'].median())

df['Department'] = df['Department'].fillna(df['Department'].mode()[0])
df['Gender'] = df['Gender'].fillna(df['Gender'].mode()[0])
df['JoinDate'] = df['JoinDate'].fillna(df['JoinDate'].mode()[0])

print("Cleaned Data (first 5 rows):")
print(df.head())
print("\n")

df['Tenure'] = 2025 - df['JoinDate'].dt.year

def salary_category(salary):
    if salary < 50000:
        return 'Low'
    elif salary <= 90000:
        return 'Medium'
    else:
        return 'High'

df['SalaryCategory'] = df['Salary'].apply(salary_category)

print("Data with Tenure and SalaryCategory (first 5 rows):")
print(df[['Salary', 'Tenure', 'SalaryCategory']].head())
print("\n")

avg_salary_by_dept = df.groupby('Department')['Salary'].mean().reset_index()
print("Average Salary by Department:")
print(avg_salary_by_dept)
print("\n")

gender_count_by_dept = df.groupby(['Department', 'Gender']).size().unstack(fill_value=0)
print("Gender Count by Department:")
print(gender_count_by_dept)
print("\n")

avg_rating_by_dept = df.groupby('Department')['PerformanceRating'].mean().reset_index()
print("Average Performance Rating by Department:")
print(avg_rating_by_dept)
print("\n")

low_performers = df[df['PerformanceRating'] <= 2]
print("Low Performers (Performance Rating ≤ 2):")
print(low_performers)
print("\n")

with pd.ExcelWriter('employee_analysis_result.xlsx') as writer:
    df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
    avg_salary_by_dept.to_excel(writer, sheet_name='Avg_Salary_By_Dept', index=False)
    gender_count_by_dept.to_excel(writer, sheet_name='Gender_Count_By_Dept')
    avg_rating_by_dept.to_excel(writer, sheet_name='Avg_Rating_By_Dept', index=False)
    low_performers.to_excel(writer, sheet_name='Low_Performers', index=False)

print("All results have been saved to 'employee_analysis_result.xlsx'.")
print("\n")

plt.figure(figsize=(8,5))
plt.figure(figsize=(8,5))
sns.barplot(x='Department', y='Salary', data=avg_salary_by_dept, hue='Department', palette='Blues', legend=False)
plt.title('Average Salary by Department')
plt.xlabel('Department')
plt.ylabel('Average Salary')
plt.show()

gender_counts = marketing['Gender'].value_counts()

plt.figure(figsize=(6,6))
plt.pie(gender_counts, labels=gender_counts.index, autopct='%1.1f%%', colors=sns.color_palette('Set2'))
plt.title('Gender Distribution in Marketing (Seaborn Colors)')
plt.show()