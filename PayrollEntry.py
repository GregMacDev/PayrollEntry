import xlrd
import os
import calendar

class Employee:
    # Employee class that holds Payroll information for each Employee
    # Employee Name, Employee Number, Salary, Excel Column, 
    # Retirement Amount, Distribution Number (for Sage 50 formatting)
    def __init__(self, name = "", number = 0, salary = 0, col = 0, retirement = 0, numOfDist = 0):
        self.name = name
        self.number = number
        self.salary = salary
        self.retirement = retirement
        self.numOfDist = numOfDist
        self.weekly = {}
        self.jobs = {}
        self.col = col

# Basic employee Payroll data. To be moved to external file.
emp1 = Employee("Employee 1", 77560, 33.00, 30, 0)
emp2 = Employee("Employee 2", 77520, 33.00, 0, 0, 1)
emp3 = Employee("Employee 3", 77540, 40.00, 12, -48.00)
emp4 = Employee("Employee 4", 77530, 40.00, 3, -75.00)
emp5 = Employee("Employee 5", 77550, 33.00, 21, -75.00, 1)
employees = (emp1, emp2, emp3, emp4, emp5)

def payrollEntry(employeeList, date):
    ## Iterates through employee list menu for payroll entry
    
    for i in employeeList:
        if i.name != "Employee 2":
            getPayroll(i, date)
            
def getPayroll(employee,date):
    ## Takes in an employee and runs through the days of the week
    ## to input a job for each day
    
    days = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    
    weekdayRow = timeSheetDate(date)
    for i in days:            
        day = setDay(employee, weekdayRow)
        if day:
            employee.weekly[i] = day
        weekdayRow += 1
    setWeeklyJobs(employee)
    return 

def timeSheetDate(date):
    # Reads the check date, and returns the row of the payroll week's start date.
      
    #sets weekStartDate, in string format, as the beginning of the payroll week
    weekStartDate = payrollWeekBegin(date)             
    row = 2
    
    # Linear search through spreadsheet dates until week start date is found, or
    # until the list ends.    
    while True:
        try:
            newDate = readDate(timeSheet.cell(row,1).value)
            if weekStartDate == newDate:
                return row
            row += 1
        except:
            print("no date match")
            input()
            return -1

def setDay(employee, row):
    ## Creates a list of 2 elements. (JobName, Hours)
    ## based on user input for each day's job and hours worked.
    
    day = []
    jobName = timeSheet.cell_value(row, employee.col)
    if (jobName == 0.0 or jobName == ""):
        return
    jobName = jobName.capitalize()
    day.append(jobName)
    if (jobName == "Out" or jobName == "out"):
        hours = "0"
        day.append(hours)
    elif jobName == "Off" or jobName == "Holiday" or jobName == "Sick" or jobName == "Vacation" or jobName == "Office" or jobName == "Running" or jobName == "":
        hours = "8"
        day.append(hours)
    else:
        hours = timeSheet.cell_value(row, (employee.col + 4))
        day.append(hours)
    return day

def readDate(excelDate):
    # Retrieves the date from Spreadsheet and returns it as the string
    # formatted date.
    
    date = xlrd.xldate.xldate_as_tuple(excelDate, 0)
    newDate = str(date[1]) + "/" + str(date[2]) + "/" + str(date[0]-2000)
    return newDate
            
def payrollWeekBegin(date):
    # Takes in the check date, and returns the string formatted date 
    # of the beginning of payroll week
    
    date_list = date.split("/")
    month = date_list[0]
    day = date_list[1]
    year = date_list[2]
    return(getWeekStartDay(month, day, year))

def getWeekStartDay(month, day, year):
    # Takes in the values of the check date and returns the string formatted 
    # date of the start of the Payroll week
    
    working_year = int("20" + year)
    new_day = int(day) - 10
    new_month = int(month)
    if new_day < 1:
        new_month = int(month) - 1
        if new_month == 0:
            new_month = 12
            working_year = working_year - 1
        days_in_month = calendar.monthrange(working_year, new_month)[1]
        new_day = new_day + days_in_month
    return (str(new_month) +"/" + str(new_day) + "/" + str(working_year-2000))



def setWeeklyJobs(employee):
    ## Generates class dictionary to represent the 
    ## number of hours worked per job that week.
    
    for i in employee.weekly.values():
        if not i:
            continue
        else:
            wage = getWage(employee, i[0])
            jobNumber = getJobNumber(i[0])
            hours = float(i[1])            
            if jobNumber in employee.jobs:
                tempInfo = employee.jobs[jobNumber]
                hours += tempInfo[0]
            else:
                employee.numOfDist += 1
            jobInfo = (hours, wage)
            employee.jobs[jobNumber] = jobInfo
    print(employee.name, ":", employee.jobs)
                
def getWage(employee, jobName):
    ## Sets a wage for each job, if different from employee's
    ## base wage.
    
    sheets = wb.sheet_names()
    job = ""
    for i in sheets:
        if i.lower() == jobName.lower():
            job = i
            break
    if job == "":
        return float(employee.salary)
    else:
        jobSheet = wb.sheet_by_name(job)
        wage = jobSheet.cell_value(4,8)
        if wage == "":
            return float(employee.salary)
        else:
            return float(wage)
        
def getJobNumber(jobName):
    ## Cross references the job name and returns the Job number
    ## based on the job list Excel file.
    
    if jobName.lower() == "office" or jobName.lower() == "running" or jobName.lower() == "holiday" or jobName.lower() == "sick" or jobName.lower() == "vacation" or jobName.lower() == "out":
        return ""
    number = ""
    rows = 1
    while(number == ""):
        cellValue = jobs.cell_value(rows, 1)
        if(cellValue == "" or cellValue == 0.0):
            number = getJobNumberFromSheet(jobName)
            return number
        elif(cellValue.lower() == jobName.lower()):
            number = jobs.cell_value(rows, 0)
            return number
        rows += 1

def totalNumOfDist(employees):
    ## Adds up the total number of distributions for use
    ## in the .csv file for import
    
    numOfDist = 2
    for i in employees:
        # Employee 2 is salaried, thus processed differently.
        if i.name == "Employee 2":
            numOfDist += i.numOfDist
        if i.weekly:
            if i.retirement != 0:
                i.numOfDist += 1
            numOfDist += i.numOfDist
    return numOfDist

def getJobNumberFromSheet(jobName):
    ## Sets a job number from a sheet, if not found
    ## from the job list
    
    sheets = wb.sheet_names()
    job = ""
    for i in sheets:
        if i.lower() == jobName.lower():
            job = i
            break
    if job == "":
        return ""
    else:
        jobSheet = wb.sheet_by_name(job)
        return jobSheet.cell_value(1,8)

def printPayroll(date, taxes, employees):
    ## Creates the .csv file based on the entered payroll info
    
    csvFile = open("PAYROLL.CSV", "w")
    numOfDist = totalNumOfDist(employees)
    
    # Writes payroll taxes line in CSV
    csvFile.write(date + "," + str(numOfDist) + ",72000,," + taxes 
                  + ",,FALSE,37,1,0,0\n")
    
    #Iterates through each employee to print Payroll info to CSV file
    for i in employees:
        
        # Employee 2 is Salaried and the payroll entry never changes.
        if(i.name == "Employee 2"):
                csvFile.write(date + "," + str(numOfDist) + "," 
                              + str(i.number) + "," + i.name 
                              + "," + str(700) + ",,FALSE,37,1,0,0\n")
        else:
            for j,k in i.jobs.items():
                if(j == "Office" or j == "Running"):
                    jobNumber = ""
                    wagesPaid = float(k[0]) * float(k[1])
                else:
                    jobNumber = j
                    wagesPaid = float(k[0]) * float(k[1])
                csvFile.write(date + "," + str(numOfDist) + "," 
                              + str(i.number) + "," + i.name + "," 
                              + str(wagesPaid) + "," + jobNumber 
                              + ",FALSE,37,1,0,0\n")
    for i in employees:
        if i.weekly:
            if(i.retirement != 0):
                csvFile.write(date + "," + str(numOfDist) + ",24500," 
                              + i.name + " Retirement," + str(i.retirement) 
                              + ",,FALSE,37,1,0,0\n")
    if emp5.weekly:
        csvFile.write(date + "," + str(numOfDist) + "," + str(emp5.number) 
                      + ",Child Support,-165.00,,FALSE,37,1,0,0\n")
    csvFile.write(date + "," + str(numOfDist) + ",10300,ADP Payroll,-" 
                  + str(weeklyPayroll) + ",,FALSE,37,1,0,0,")
    csvFile.close()

## Main Program Functionality and user data input

try:
    wb = xlrd.open_workbook('Job List 2021.xlsx')
except:
    print("Failed to Open File")
    input()
    
jobs = wb.sheet_by_name('All Jobs')
timeSheet = wb.sheet_by_name("TimeSheets")
numOfDist = 2
correct = "n"

while correct.lower() == "n":
    checkDate = -1
    os.system('cls')
    print("\n\n\t\t*****Sage 50 Import File ADP Payroll entry*****")    
    while checkDate == -1:
        checkDate = input("\nCheck Date (mm/dd/yy): ")
        #checkDate = getCheckDate(date)
    taxes = input("\nPayroll taxes: ")
    weeklyPayroll = input("\nADP Payroll deduction: ")
    print("\n\nYou entered:\nCheck Date:", checkDate, "\n     Taxes: " 
          + "${:,.2f}".format(float(taxes)), "\n Deduction: " 
          + "${:,.2f}".format(float(weeklyPayroll)))
    correct = input("\nIs this info correct? y/n: ")
payrollEntry(employees, checkDate)
printPayroll(checkDate, taxes, employees)
input("\n\n\nPress any key to quit.....")
