HR & Business Performance Dashboard 📊
An end-to-end HR analytics solution built with Python, Excel, and Power BI — covering attrition trends, department performance, salary analysis, and gender pay comparison across 3,000 employee records.

🛠️ Tech Stack

Python — Synthetic dataset generation (Pandas, NumPy, Faker)
Microsoft Excel — Data pipeline with 7 structured sheets and live KPI formulas
Power BI — 4-page interactive dashboard with DAX measures and slicers
DAX — Custom measures for attrition rate, avg salary, performance KPIs


📊 Dataset Overview
A synthetic dataset of 3,000 employees with 23 business-relevant features generated using Python:
CategoryColumnsEmployee InfoEmployeeID, Gender, Age, Education, Department, JobRole, LocationCompensationMonthlySalary, SalaryBand, BonusPercentPerformancePerformanceRating, TargetAchievementPct, TrainingHoursEngagementEngagementScore, JobSatisfaction, WorkLifeBalanceAttritionAttrition, TenureYears, LastPromotionDateBusinessBusinessUnit, ProjectsCompleted, RevenueContribution

📂 Excel Data Pipeline
The Excel file contains 7 structured sheets:
SheetContentsRaw Data3,000 employee records × 23 columnsKPI Summary8 live formula-driven KPIsDept SummaryHeadcount, attrition, salary by departmentSalary Band SummaryMin/max/avg salary + attrition by bandAge Group SummaryAttrition & performance by age groupBusiness Unit SummaryRevenue, projects, performance by regionData DictionaryColumn definitions and data types

📈 Power BI Dashboard Pages
Page 1 — Overview

Total Employees, Attrition Rate, Avg Salary, Avg Performance KPI cards
Headcount by Department, Gender Distribution, Attrition Rate by Department
Slicers: Department, Gender, Attrition

Page 2 — Attrition Analysis

Attrition by Department, Age Group, Salary Band
Gender-wise attrition and Engagement vs Attrition scatter plot

Page 3 — Department Performance

Avg Performance Rating, Target Achievement %, Training Hours by Department
Engagement vs Performance scatter plot

Page 4 — Salary & Compensation

Headcount by Salary Band, Avg Salary by Department
Salary vs Performance scatter and Gender Pay Comparison by Salary Band


⚙️ How to Run
Generate Dataset
bashpip install pandas numpy openpyxl faker
python generate_hr_dashboard.py
Open Dashboard

Open HR_Business_Performance_Dashboard.pbix in Power BI Desktop
Update data source path to your local Excel file
Click Refresh to load data


💡 Key Insights

Band 1 (Entry) employees show the highest attrition rate
31-40 age group has the highest attrition count
Higher engagement scores correlate with better performance ratings
Gender pay gap widens at senior salary bands


🙋‍♀️ Author
Tanishka Pagar

GitHub: https://github.com/TanishkaPagar
LinkedIn: https://www.linkedin.com/in/tanishka-pagar
