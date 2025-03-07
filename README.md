# Excel-Projects

This is my **Excel Projects** repository! It contains various projects built using Microsoft Excel, leveraging different functions and formulas. The projects focus on data analysis, data visualization, and reporting to support better decision-making. 

## Purpose  
This repository serves as a centralized space to store and track all my Excel projects, improving efficiency and accessibility for future enhancements and collaborations.  

## Projects Included  
### 1. **Budget Tracker**
   
This project is a Budget Tracking project which is a simple and efficient tool for managing income, expenses, and savings.  

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/income.JPG?raw=true)

The **Income Breakdown** chart shows how your salary is affected by deductions like taxes, insurance, and superannuation. These reduce your total income, but a bonus adds to it, resulting in your final net income. This Waterfall Chart is useful for tracking budget changes and understanding how different amounts increase or decrease over time.

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/expenses.JPG?raw=true)]

The **Expenses Breakdown** chart shows how your total spending is divided into different categories. The biggest costs are for housing and food, followed by smaller expenses like travel, fuel, gym, and phone bills. This Funnel Chart helps show how expenses decrease step by step. It is often used to track spending or see where people drop off in a process, like a customer journey.

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/donut.JPG?raw=true)

The **% Expenses chart** is a Donut Chart, which looks like a pie chart but has a hole in the middle. It shows that 94% of your income goes to expenses, leaving only 6% as savings. Donut charts are great for dashboard reports because they look clean and make it easy to compare parts of a whole. They are often used instead of pie charts when you need a modern look with space for extra labels.

You can download this project here:
[Download Budget Tracker File](https://github.com/Sofiya-Banmala/Excel-Project/raw/main/Budget%20Tracking%20Portfolio.xlsx)

### 2. **Essential Excel Functions for Data Analysis**

This is an **Excel file** where I practiced some **essential Excel functions and formulas**. It includes common formulas used for basic calculations, lookup and reference, conditional counting, summing, filtering, and data organization and visualization required for data analysis tasks.  

**Download the file here:** [Download Excel File](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/Essential_Excel_Functions_for_Data_Analysis_Practise.xlsx)  

Here are some of the business questions that i solved using these functions and formuas:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/excel%20func.JPG?raw=true)

This business questions helps to find the techniques covered that enables efficient calculation of salaries, headcounts, and averages, while also supporting dynamic data filtering, sorting, and gender-based analysis. They facilitate the generation of reports, advanced lookups, error handling, and statistical or time-based analysis, providing valuable insights for decision-making.

Below, we describe each business question in detail, along with the functions used to achieve the desired results.

**1. Total Salary and Headcount by Department**

### Description:
This analysis calculates the total number of employees (HeadCount) and their total salary across different departments. Additionally, it separately calculates the headcount and total salary of permanent employees within each department.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/1an%202.JPG?raw=true)


### Functions Used:

- `=COUNTIF(staff[Department], A4)`  
  - This function counts the number of employees in a given department.

- `=SUMIF(staff[Department], A5, staff[Salary])`  
  - This function calculates the total salary of all employees in a given department.

- `=COUNTIFS(staff[Department], A5, staff[Employee type], "Permanent")`  
  - This function counts the number of permanent employees in a given department.

- `=SUMIFS(staff[Salary], staff[Department], A4, staff[Employee type], "Permanent")`  
  - This function calculates the total salary of permanent employees in a given department.

**2. Average Salary by Department**

### Description:
This analysis calculates the average salary of employees in each department. The average salary provides insights into compensation distribution across departments.

### Function Used:

- `=AVERAGEIF(staff[Department], A4, staff[Salary])`  
  - This function calculates the average salary of employees in each department.

**3. All Employees with More Than $100K Salary**

### Description:
This analysis identifies employees earning more than $100,000 annually across different departments. By filtering employees based on salary, we can gain insights into high-income earners within the organization. This data is useful for workforce planning, salary benchmarking, and identifying top-earning employees across locations.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/3.JPG?raw=true)

### Functions Used:
- `=FILTER(staff, staff[Salary] > D2)`: This function is used to filter and display all employees whose salary exceeds $100,000. The `FILTER` function dynamically retrieves records based on the salary condition.
- `=staff[#Headers]`: This function is used to reference the column headers dynamically, ensuring that the extracted data includes appropriate labels.

**4. All Female Employees with More Than $100K Salary**

### Description:
This analysis identifies all female employees who earn more than $100,000 annually. The purpose of this analysis is to examine gender-based salary distribution and identify high-earning female employees within the organization.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/4.JPG?raw=true)

### Functions Used:
- `=CHOOSECOLS(FILTER(staff, staff[Gender] = "Female", staff[Salary] > 100000), 1,2,3,4,5,6)`:  
  - `FILTER(staff, staff[Gender] = "Female", staff[Salary] > 100000)`: Filters employees based on gender (Female) and salary greater than $100,000.
  - `CHOOSECOLS(..., 1,2,3,4,5,6)`: Selects specific columns (Emp ID, First Name, Last Name, Gender, Department, Salary) from the filtered data.
 
**5. All Female Employees with More Than $100K Salary Who Joined in 2020 or After**

### Description:
This analysis identifies all female employees who earn more than $100,000 annually and joined the organization in 2020 or later. The purpose of this analysis is to understand the impact of recent hires with high salaries and assess trends related to high-earning female employees in the company.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/5.JPG?raw=true)

### Functions Used:
- `=FILTER(staff, (staff[Gender]="Female") * (staff[Salary]>100000) * (YEAR(staff[Start Date])>=2020))`:
  - `FILTER(staff, ...)`: Filters the dataset based on the specified conditions.
  - `(staff[Gender]="Female")`: Filters for female employees.
  - `(staff[Salary]>100000)`: Filters for employees with a salary greater than $100,000.
  - `(YEAR(staff[Start Date])>=2020)`: Filters for employees who joined in 2020 or after.

**6 & 7. Salary Analysis: Lowest, Highest, and Top 5 Salary Values (Overall and by Gender)**

### Description:
This analysis identifies the lowest, highest, and top 5 salary values within the organization as well as by gender (Male and Female). The goal is to assess salary distribution across the workforce, highlight trends within different genders, and identify the range of salaries, including the highest earners. This will help evaluate overall salary equity and gender-based disparities, if any, in compensation within the company.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/67.JPG?raw=true)

### Functions Used:
- `=MIN(staff[Salary])`: Identifies the lowest salary across all employees.
- `=MAX(staff[Salary])`: Identifies the highest salary across all employees.
- `=LARGE(staff[Salary], E6)`: Returns the top 5 highest salary values from the dataset based on the rank number (from 1 for the highest, 2 for second highest, etc.).
- `=MINIFS(staff[Salary], staff[Gender], "Male")`: Identifies the lowest salary for male employees.
- `=MAXIFS(staff[Salary], staff[Gender], "Male")`: Identifies the highest salary for male employees.
- `=MINIFS(staff[Salary], staff[Gender], "Female")`: Identifies the lowest salary for female employees.
- `=MAXIFS(staff[Salary], staff[Gender], "Female")`: Identifies the highest salary for female employees.
- `=TAKE(SORT(staff[Salary], , -1), 5)`: Returns the top 5 highest salaries after sorting the dataset in descending order.

**8 & 9. Department List Analysis: All Departments and Comma-Separated List**

### Description:
This analysis generates a list of all unique departments within the organization, as well as a comma-separated list of these departments. The goal is to provide an overview of the organizational structure by highlighting the various departments and creating a consolidated, easily-readable list for reporting purposes.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/8.JPG?raw=true)

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/9.JPG?raw=true)

### Functions Used:
- `=UNIQUE(staff[Department])`: Extracts a unique list of all departments from the `Department` column, removing duplicates.
- `=TEXTJOIN(", ", TRUE, UNIQUE(staff[Department]))`: Combines all unique departments into a single cell, with each department name separated by a comma and a space. The `TRUE` argument ensures that any empty cells are ignored.

**10. Employee Details Lookup**

### Description:
This analysis provides a lookup of employee details based on specific identifiers such as Employee ID or Last Name. The purpose of this analysis is to retrieve and display information about a particular employee, including their first name, last name, department, and salary, based on a given search criterion. This can be useful for HR departments, payroll teams, or any role requiring quick access to employee-specific data.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/1%0.JPG?raw=true)

### Functions Used:
- `=VLOOKUP(B3, staff, 2, 0)`: This function looks up the Employee ID (provided in cell `B3`) in the employee data (`staff`), retrieving the corresponding value from the second column (First Name in this case). The `0` argument ensures an exact match.
- `=INDEX(staff[Emp ID], B15)`: This function uses the index to retrieve the Employee ID from the dataset based on the row number provided in `B15`.
- `=MATCH(B14, staff[Last Name], 0)`: This function searches for the Last Name (given in `B14`) in the `staff` table and returns the row number where the last name is found. The `0` ensures that only an exact match is returned.





**Functions Used in the File**

- **COUNTA** – Counts all non-empty rows in a given range.  
- **COUNT** – Counts only numeric values in a range.  
- **COUNTIF** – Counts how many times a specific value appears in a range.  
- **SUMIF** – Adds up values based on a condition.  
- **COUNTIFS** – Counts values based on multiple conditions.  
- **SUMIFS** – Adds up values based on multiple conditions.  
- **AVERAGEIF** – Calculates the average of values that meet a condition.  
- **MINIF / MAXIF** – Finds the smallest or largest value based on a condition.  
- **MINIFS / MAXIFS** – Finds the smallest or largest value based on multiple conditions.  
- **FILTER** – Filters data based on a condition.  
- **SORT & TAKE** – Sorts data and extracts a subset of rows.  
- **VLOOKUP** – Searches for a value in a column and returns a result from another column.  
- **XLOOKUP** – A more flexible version of VLOOKUP.  
- **CHOOSECOLS** – Selects specific columns from a dataset.  
- **MEDIAN** – Finds the middle value of a dataset.  

This file is useful for anyone learning **Data Analysis in Excel**.

### 3. **Excel Case Study Questions Solving**

Case Study 1

Case Study 2

Case Study 3

## How to Use?  
1. **Browse** the repository to find the desired Excel tool.  
2. **Download** the spreadsheet and open it in **Microsoft Excel** or **Google Sheets**.  
3. **Follow** the instructions in each file to input data and analyze results.  

## Contribution  
Feel free to **fork**, **suggest improvements**, or share ideas for new Excel-based tools! 🚀  

---
  
🔹 *Stay tuned for more Excel-based solutions!*  
