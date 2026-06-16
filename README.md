# HR & Business Performance Dashboard 📊
 
An end-to-end HR analytics solution built with **Python, Excel, and Power BI** — covering attrition trends, department performance, salary analysis, and gender pay comparison across 3,000 employee records.
 
---
 
## 🛠️ Tech Stack
 
- **Python** — Synthetic dataset generation (Pandas, NumPy, Faker)
- **Microsoft Excel** — Data pipeline with 7 structured sheets and live KPI formulas
- **Power BI** — 4-page interactive dashboard with DAX measures and slicers
- **DAX** — Custom measures for attrition rate, avg salary, performance KPIs
---
 
## 📁 Project Structure
 
```
hr-business-performance-dashboard/
│
├── HR_Business_Performance_Dashboard.xlsx   # Excel data pipeline
├── HR_Business_Performance_Dashboard.pbix   # Power BI dashboard
├── generate_hr_dashboard.py                 # Python script to generate dataset
└── README.md
```
 
---
 
## 📊 Dataset Overview
 
A synthetic dataset of **3,000 employees** with **23 business-relevant features** generated using Python:
 
| Category | Columns |
|---|---|
| Employee Info | EmployeeID, Gender, Age, Education, Department, JobRole, Location |
| Compensation | MonthlySalary, SalaryBand, BonusPercent |
| Performance | PerformanceRating, TargetAchievementPct, TrainingHours |
| Engagement | EngagementScore, JobSatisfaction, WorkLifeBalance |
| Attrition | Attrition, TenureYears, LastPromotionDate |
| Business | BusinessUnit, ProjectsCompleted, RevenueContribution |
 
---
 
## 📂 Excel Data Pipeline
 
The Excel file contains **7 structured sheets:**
 
| Sheet | Contents |
|---|---|
| Raw Data | 3,000 employee records × 23 columns |
| KPI Summary | 8 live formula-driven KPIs |
| Dept Summary | Headcount, attrition, salary by department |
| Salary Band Summary | Min/max/avg salary + attrition by band |
| Age Group Summary | Attrition & performance by age group |
| Business Unit Summary | Revenue, projects, performance by region |
| Data Dictionary | Column definitions and data types |
 
---
 
## 📈 Power BI Dashboard Pages
 
### Page 1 — Overview
- Total Employees, Attrition Rate, Avg Salary, Avg Performance KPI cards
- Headcount by Department (bar chart)
- Gender Distribution (donut chart)
- Attrition Rate by Department (bar chart)
- Slicers: Department, Gender, Attrition
### Page 2 — Attrition Analysis
- Attrition by Department, Age Group, Salary Band
- Gender-wise attrition (donut chart)
- Engagement vs Attrition scatter plot
### Page 3 — Department Performance
- Avg Performance Rating by Department
- Target Achievement % by Department
- Training Hours by Department
- Engagement vs Performance scatter plot
### Page 4 — Salary & Compensation
- Headcount by Salary Band
- Avg Salary by Department
- Salary vs Performance scatter plot
- Gender Pay Comparison by Salary Band
---
 
## ⚙️ How to Run
 
### Generate Dataset
```bash
pip install pandas numpy openpyxl faker
python generate_hr_dashboard.py
```
This generates the Excel file with all 7 sheets.
 
### Open Dashboard
1. Open `HR_Business_Performance_Dashboard.pbix` in **Power BI Desktop**
2. Update the data source path to point to your local Excel file
3. Click **Refresh** to load the data
---
 
## 💡 Key Insights
 
- **Band 1 (Entry)** employees show the highest attrition rate
- **31-40 age group** has the highest attrition count
- Departments with higher engagement scores show better performance ratings
- Gender pay gap is minimal at entry level but widens at senior bands
---
 
## 🙋‍♀️ Author
 
GitHub: https://github.com/TanishkaPagar
LinkedIn: https://www.linkedin.com/in/tanishka-pagar

## 👩‍💻 Developed By
Tanishka Pagar — LABTECH Internship 2026
