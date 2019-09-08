# PolarStar

Sheet Generator.

Note: Results.txt is just there to load previously saved form inputs.

# Monthly generation steps:
Note: The first step is to make sure Results.txt as well atleast one excel file with names "ltd" or "qlhf" in the dist folder.

1) Executable can be found in SheetGen > dist launch it and wait for GUI to pop up.
2) Pick from the menu the Admin File name as well as the reference files for both Advisors and Key files.
3) Write down the advisor name you wish to generate the invoice for
4) Pick a date from the dropdown menu for the invoice
5) Make sure the advanced options tab has the correct inputs (Refer to Advanced Options Inputs section at the end for explanations)
6) Press "Receive file"

Following the Receive File command the program will either:
1) Generate a master file for which it advises you to fill in unassigned advisors - during which you can rerun the "Receive file" command once you edit the file to generate the invoice.
2) Ask you again to fill in the unassigned advisors in the master excel file if you haven't done so correctly.
3) Generate the invoice

# Warnings to watch out for:
There are several warnings that are generated:
1) WARNING: You created x invoice but Admin File indicates another Date\
This warning occurs when one tries to generate an invoice using an admin file from the previous month instead of the month specified by the date chosen.
2) ERROR: One of Management Fee or Performance Fee are spelled wrong\
This indicates that either one or both of those two are spelled wrong, or, the data could not be written because of invalid entries in the advanced options tab.

# Quarterly generation steps:
Note: The first step is to make sure that there exists atleast one advisor invoice containing "31st" in the file name.

Press the "Quarterly" button and input all three monthly sheets names that you generated, as well as the advisor name and date range.

Press merge files and the excel file should be created in dist.

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
