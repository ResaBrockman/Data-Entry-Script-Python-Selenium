# Data-Entry-Script-Python-Selenium
A simple, recursive script that uses Selenium to enter data from a spreadsheet into a reporting website. This significantly saved time on a monthly data entry task, however, this website is no longer in use.

The first few lines of code start a Chrome browser through Selenium, and logged into the reporting site. Login information was taken out of this file to prevent identifying the business despite the fact that the website is no longer in use (and therefore login information is otherwise useless.)

File format for the excel file with applicable data contained information in the following column order: Reporting Class (if the first row of a new class, otherwise blank), Name, SSN, Occupation, SSN, Hours Worked, Earnings. SSN was included twice as the way the data came.

The next segments of code, clearly marked, are the methods followed by the actual data entry run.

PROCESS:
1) The code would find the applicable reporting class on the website when Column 1 of the associated spreadsheet was not blank.
2) Delete worker numbers from last month (these autopopulate from previously entered data - it was much easier to delete and reenter applicable lines than to match existing website lines with new months spreadsheet lines.) 
3) Enter the information for each line of the spreadsheet until it reached a new reporting class, Save.
4) Find the new reporting class. 
5) Continue until it hit blank lines on the spreadsheet.

This was a task that took hours of data entry, reduced to the time it took the script to run in the background. Totals were always confirmed after running to make sure the program had not malfunctioned.

This reporting was moved to a new website and format after 8 months of use of this script.
