# Drywall Company Work Sheet Process

## About 

Designed a process for retrieving, manipulating and reporting company work sheets. More specifically, on a weekly basis, employees submit their work completed via a Google Form. On a specific day of the week, all entries by employees are manipulated in python to a form that allows the supervisor to make any deduction and/or adjustments to the excel sheets created. Lastly, using python, those revised sheets are converted to a pdf work sheet report for each employee that includes their work and pay. 

## Motivation

I was inspired to create this process from seeing an individual complete work sheets (better known as time sheets) for his/her drywall company on a weekly basis for all employees. Previously, the way the work completion per employee was gathered was by the employee sending an email to the supervisor at the end of the week. The retrieval of data came from an email which the supervisor would then format into a well-formatted excel sheet. This excel sheet was formatted to be printed for employees to keep track of their weeks work and pay.

## Noes 

- Tools used for this project are highly influenced by the supervisor. For instance, the idea to leverage google services strives from catering to the user
- For privacy reasons/restrictions, the files presented are from testing phases and no company information is shown
- This project is not completed, more is still being added

## File Content

- subcontractors.xlsx : excel sheet containing employee personal information as well as their rate of pay
- weekXentries.xlsx: excel sheet containing entries from employees google form submissions 
- manipulation.py: python script for manipulating weekXentries.xlsx to work_sheet_20210101.xlsx
- work_sheet_20210101.xlsx: excel sheet containing a different sheet per employee and 'all employee' sheet for a summary of all the work completed 
- manipulation.py: python script for manipulating work_sheet_20210101.xlsx to Mariana_Mariles.pdf
- Mariana_Mariles.pdf: final work sheet for employee

## Libraries Used

- pandas
- numpy
- openpyxl
- reportlab

  