# PayrollEntry
## All of these files are proprietary files and were created by me for use within the company I work for. Any sensitive or identifying data has been removed and replaced with dummy data.

This is a proprietary python script I created to take information from an excel spreadsheet, in conjunction with an ADP payroll report and create a CSV file to import into Sage 50 Accounting.

I found myself spending several hours every week cross referencing the reported hours of our workers and applying them to our ongoing jobs for job costing purposes. On top of that, the reported hours needed to be recorded in our accounting software, Sage 50, each week, with the wages being applied to the proper jobs. My responsibilities to the company are not strictly data entry, so these tasks related to job costing and payroll took me away from producing and procuring more work to generate more revenue for the company. In noticing that the requirements to fulfil this part of my job were repetative each week, I set out to design a program or a script to streamline and reduce a few hours of office work required, per week. This is where this python script was created.

PayrollEntry.py is the script that I created. The program prompts the user for the date of the payroll, the taxes required (reported from ADP), and the dollar amount required (reported from ADP). From there, using the date entered, reads from the spreadsheet, 'Job List 2021.xlsx', and steps through the 'TimeSheets' sheet, gathering the information for the job site each employee worked on each day of the payroll week. The script records each job the employee worked during the week, tallies up the hours worked on each job,  multiplies the hours by the wage of each employee, and stores this information in a python dictionary. Once the data is collected, the script formats the data from the spreadsheet, along with other hardcoded payroll deduction data, and outputs to PAYROLL.CSV, which can then be imported to Sage 50 Accounting, as a general journal entry.

<br>
Files included:
<br>
<b>PayrollEntry.py</b> - Script that reads Job List 2021.xlsx spreadsheet and formats data to CSV to import to Sage 50 Accounting
<br>
<b>Job List 2021.xlsx</b> - Excel Spreadsheet that tracks job costing and employee timesheets

<b>PAYROLL.CSV</b> - Ouput file used to import General Journal Entry to Sage 50 Accounting.
