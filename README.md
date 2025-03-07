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


## 1. Total Salary and Headcount by Department

### Description:
This analysis calculates the total number of employees (HeadCount) and their total salary across different departments. Additionally, it separately calculates the headcount and total salary of permanent employees within each department.

### Table Structure:



### Functions Used:

- `=COUNTIF(staff[Department], A4)`  
  - This function counts the number of employees in a given department.

- `=SUMIF(staff[Department], A5, staff[Salary])`  
  - This function calculates the total salary of all employees in a given department.

- `=COUNTIFS(staff[Department], A5, staff[Employee type], "Permanent")`  
  - This function counts the number of permanent employees in a given department.

- `=SUMIFS(staff[Salary], staff[Department], A4, staff[Employee type], "Permanent")`  
  - This function calculates the total salary of permanent employees in a given department.

---

## 2. Average Salary by Department

### Description:
This analysis calculates the average salary of employees in each department. The average salary provides insights into compensation distribution across departments.

### Table Structure:



### Function Used:

- `=AVERAGEIF(staff[Department], A4, staff[Salary])`  
  - This function calculates the average salary of employees in each department.


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
