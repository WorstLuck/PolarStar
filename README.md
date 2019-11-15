# Current interface
![Image of interface 2019/09/08](https://github.com/WorstLuck/PolarStar/blob/master/Current%20Interface.png)

# Summary
This is an application that automatically generates invoices associated with a particular index by taking a dataset in the form of an "xlsx" or "xls" file as well as two reference datasets that are used to match each index and associate the appropriate fees to based on their key/value pairs. The user then has an option to merge invoices as well as create pivot tables based on their chosen index. 

The invoice generated has two sheets, one is the respective rows sampled from the original dataset and the other is an aggregate calculation summing up two types of fees involved.

These invoices can be generated for any index chosen by the user. 

Note: Results.txt is just there to load previously saved form inputs.

# Monthly generation steps:
Note: The first step is to make sure Results.txt as well atleast one excel file with names "ltd" or "qlhf" in the dist folder.

# Section 1.
1) Executable can be found in SheetGen > dist launch it and wait for GUI to pop up.
2) Pick from the menu the Admin File name as well as the reference files for both Advisors and Key files.
3) Refer to Section 2 to create an RMB pivot before you continue if needed
4) Chose the advisor from the menu that you wish to generate the invoice for
5) Pick a date from the dropdown menu 
6) Make sure the advanced options tab has the correct inputs (Refer to Advanced Options Inputs section at the end for explanations)
6) Press "Receive file"

# Section 2.
1) Press Make RMB pivot
2) Choose your Admin and RMB files from the drop-down menu
3) Refresh 
4) Pick the values and Broker Corporate column names and press Make Pivot Table
5) The program should then write to a file with the name "RMB_split_rmb table.xlsx"
6) Continue from step 3 in Section 1.

Following the Receive File command the program will either:
1) Generate a master file for which it advises you to fill in unassigned advisors - during which you can rerun the "Receive file" command once you edit the file to generate the invoice.
2) Ask you again to fill in the unassigned advisors in the master excel file if you haven't done so correctly.
3) Generate the invoice

# Warnings to watch out for:
There are several warnings that are generated:
1) WARNING: You created x invoice but Admin File indicates another Date\
This warning occurs when one tries to generate an invoice using an admin file from the previous month instead of the month specified by the date chosen.
2) ERROR: "Toplevel entry!" Restart the application to fix this problem.
2) ERROR: One of Management Fee or Performance Fee are spelled wrong\
This indicates that either one or both of those two are spelled wrong, or, the data could not be written because of invalid entries in the advanced options tab.

# Quarterly generation steps / Invoice joining :
Press the Join Invoices button and pick your 3 invoices followed by "Merge File" to generate a file with the suffix "Full"

# Terminology:
Admin File refers to the admin file that is received\
Advisor Reference File refers to the reference file to match Investor/series combinations\
Key refers to the key file (has to have columns (Mgnt Fee and Perf. Fee))\

# Advanced Options Inputs:
Admin/Advisor/Key sheet names - Specify the sheet names for all \
Admin/Advisor Investor column names - Specify the name of the "Investor" column for both sheets\
Admin/Advisor Series column names - Specify the name of the "Series" column for both sheets\
Admin Management Fee/Performance fee column names - Specify the names of both of those that are in the admin file\
Admin/Advisor/Key columns start row - Specify the starting row (on excel e.g 2 or 4) where the column headers begin.\
Column range - Specify the column range in excel to read the data until\
