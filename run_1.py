import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Load & Clean Data
df = pd.read_excel("employee_performance.xlsx", engine="openpyxl")
df['JoinDate'] = pd.to_datetime(df['JoinDate'], format='%d-%m-%Y', errors='coerce')
df['Salary'] = pd.to_numeric(df['Salary'], errors='coerce')
df['PerformanceRating'] = pd.to_numeric(df['PerformanceRating'], errors='coerce')

df.fillna({
    'Salary': df['Salary'].mean(),
    'PerformanceRating': df['PerformanceRating'].mode()[0],
    'JoinDate': pd.to_datetime('01-01-2020', format='%d-%m-%Y')
}, inplace=True)

print("✅ 1. Data Cleaning - First 5 rows:")
print(df.head(), "\n")

# 2. Feature Engineering
df['Tenure'] = 2025 - df['JoinDate'].dt.year

def get_salary_category(salary):
    if salary < 50000:
        return 'Low'
    elif 50000 <= salary <= 90000:
        return 'Medium'
    else:
        return 'High'

df['SalaryCategory'] = df['Salary'].apply(get_salary_category)

print("✅ 2. Feature Engineering - Sample:")
print(df[['Name', 'JoinDate', 'Tenure', 'Salary', 'SalaryCategory']].head(), "\n")

# 3. Aggregated Analysis
avg_salary_by_dept = df.groupby('Department')['Salary'].mean().reset_index()
gender_count_by_dept = df.groupby(['Department', 'Gender']).size().reset_index(name='Count')
avg_rating_by_dept = df.groupby('Department')['PerformanceRating'].mean().reset_index()
low_performers = df[df['PerformanceRating'] <= 2]

print("Aggregated Analysis Outputs:")
print("Average Salary by Department:")
print(avg_salary_by_dept, "\n")

print("Gender Count by Department:")
print(gender_count_by_dept, "\n")

print("Average Performance Rating by Department:")
print(avg_rating_by_dept, "\n")

print("Employees with Performance Rating ≤ 2:")
print(low_performers[['Name', 'Department', 'PerformanceRating']], "\n")

# 4. Save to Excel
with pd.ExcelWriter("employee_analysis_result.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
    avg_salary_by_dept.to_excel(writer, sheet_name='Avg_Salary_By_Dept', index=False)
    gender_count_by_dept.to_excel(writer, sheet_name='Gender_Count_By_Dept', index=False)
    avg_rating_by_dept.to_excel(writer, sheet_name='Avg_Rating_By_Dept', index=False)
    low_performers.to_excel(writer, sheet_name='Low_Performers', index=False)

# 5. Visualizations

plt.figure(figsize=(8, 6))
sns.barplot(data=avg_salary_by_dept, x='Department', y='Salary')
plt.title('Average Salary by Department')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('avg_salary_by_dept.png')
plt.show()

plt.figure(figsize=(6, 6))
salary_counts = df['SalaryCategory'].value_counts()
plt.pie(salary_counts, labels=salary_counts.index, autopct='%1.1f%%', startangle=140)
plt.title('Salary Category Distribution')
plt.tight_layout()
plt.savefig('salary_category_distribution.png')
plt.show()
