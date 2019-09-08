# PolarStar

Sheet Generator.

Note: Results.txt is just there to load previously saved form inputs.

# Monthly generation steps:
Executable can be found in SheetGen > dist launch it and wait for GUI to pop up, then press load previous and run the exe by pressing "Receive file" -  make sure that the excel files loaded are in the dist folder.

Following this, the program will either:
1) Create an invoice for the chosen advisor
2) Return an excel file prefixed "J-" where it advises you to fill in the unmatched advisors

If 2) occurs, then fill in each unmatched respective advisor in the generated file and rerun "receive file" to obtain the invoice.

# Quarterly generation steps:
Press Quarterly? button and Input all three monthly sheets names that you generated, as well as the advisor name and date range.

Press merge files and the excel file should be created in dist.

# Terminology:
Admin refers to the admin file that is received\
Advisor refers to the reference file to match Investor/series combinations\
Key refers to the key file (has to have columns (Mgnt Fee and Perf. Fee))\

# Inputs:
Admin/Advisor/Key file names - Specify File names for all (including xlsx/xls extension)\
Admin/Advisor/Key sheet names - Specify the sheet names for all \
Admin/Advisor Investor column names - Specify the name of the "Investor" column for both sheets\
Admin/Advisor Series column names - Specify the name of the "Series" column for both sheets\
Admin Management Fee/Performance fee column names - Specify the names of both of those that are in the admin file\
Admin/Advisor/Key columns start row - Specify the starting row (on excel e.g 2 or 4) where the column headers begin.\
Advisor Name - Specify the advisor to generate the sheet for\
Date - Specify the date to name the file appropriatly\
Column range - Specify the column range in excel to read the data until\
