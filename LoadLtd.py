#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import tkinter as tk
from tkinter import *
import os

# Initialise variables
def initialvalues(**d):
    print('--------------------------------')
    print("Starting initialisation process")
    # Advanced options labels
    global sheet_name_admin
    global Investor_admin
    global Series_admin
    global skiprows_admin
    global sheet_name_advisor
    global Investor_advisor
    global Series_advisor
    global Advisor_advisor
    global skiprows_advisor
    global sheet_name_key
    global skiprows_key
    global Range
    global file_1
    global file_2
    global file_3
    global Advisor
    global Monthname
    global Mgnt_admin
    global Perf_admin
    global currencytype
    global RefFiles
    global MasterFiles
    global Months
    global Dates
    global Keys
    
    # Quarterly Globals
    global FirstMonth
    global SecondMonth
    global ThirdMonth
    global Advisornew
    global AdvisorFiles
    
    # Merge globals
    global file1
    global file2
    global file3

    #Pivot globals
    global adminfile
    global admin
    global file_admin
    global Values
    global Broker
    global RmbFileFinal
    global openalloc
    
    #Stupid excel cant name sheets with slashes...
    global fee_sheet
    global AdvisorClean
    
    #Bank to pivot
    global Bank
    
    # Dropdown menu variables
    if '.xlsx' in ' '.join(os.listdir()) or 'xls' in ' '.join(os.listdir()):
        RefFiles = [element for element in os.listdir() if ('xlsx' in element.lower() or 'xls' in element.lower())]
        MasterFiles = [element for element in os.listdir() if ('ltd' in element.lower() or 'qihf' in element.lower() 
                 or 'master' in element.lower()) and ('xlsx' in element.lower() or 'xls' in element.lower())]
    else:
        print("WARNING: No xls/xlsx files found in application directory or file names do not contain qihf/ltd ")
        RefFiles = ['NO SHEETS FOUND']
        MasterFiles = ['NO SHEETS FOUND']
        
    Months = {'01':'January','02':'February','03':'March','04':'April','05':'May','06':'June','07':'July','08':
              'August','09':'September','10':'October','11':'November','12':'December'}
    Dates = ['31st ' + element for element in Months.values()]
    AdvisorFiles = [element for element in os.listdir() if '31st' in element.lower()] + ['None']
    
    
    #Option menus for front page
    option1 = OptionMenu(front,adminvar,*RefFiles).grid(row = 1, column = 0)
    option2 = OptionMenu(front,advisorvar,*MasterFiles).grid(row = 1, column = 1)
    option3 = OptionMenu(front,keyvar,*RefFiles).grid(row = 1, column = 2)
    Month = OptionMenu(front,Date,*Dates).grid(row=3, column=1)
    RMBo = OptionMenu(front,RMBSplit,*RefFiles).grid(row=3, column=2)
    
    #Pivot menus
    adminfile = OptionMenu(Piv,adminvar,*RefFiles).grid(row = 4, column = 0)
    RMBfile = OptionMenu(Piv,RMB,*RefFiles).grid(row = 2, column = 0)
    file_admin = RMB.get()
    admin = adminvar.get()

# # # # Set initial values (To reduce time)
#     adminvar.set('Polar Star Ltd_NAV Workbook_31-10-2019.xlsx')
#     advisorvar.set('Master 31st October Ltd.xlsx')
#     keyvar.set('Polar_Star_Fund_Ltd-Rebate(Advisor)-Abdullah-August-2019.xlsx')
# #     Date.set('31st September')

    # File name
    file_1 = adminvar.get()
    file_2 = advisorvar.get()
    file_3 = keyvar.get()
    # Date
    Monthname = Date.get()
    # RMB
    RmbFileFinal = RMBSplit.get()
    
    # Option Menu to show available advisors by using Key File index
    try:
        Keyfile = pd.read_excel(file_3,sheet_name=sheet_name_key,skiprows = skiprows_key-1)
        Keys = Keyfile.index.tolist()
        Advisor = OptionMenu(front,Advisorname,*Keys).grid(row=3, column=0)
    except:
        Keys = ['NO ADVISORS FOUND - CHOOSE KEY FILE AND REFRESH!']
        Advisor = OptionMenu(front,Advisorname,*Keys).grid(row=3, column=0)
        
    #Pivot trys
    try:
        dfadmin = pd.read_excel(file_admin)
        colist = [element for element in dfadmin.columns]
        Values = OptionMenu(Piv,ValuesName,*colist).grid(row = 2, column = 1)
        Broker = OptionMenu(Piv,BrokerCol,*colist).grid(row = 2, column = 2)
    except:
        K = ['None']
        Values = OptionMenu(Piv,ValuesName,*K).grid(row= 2,column=1)
        Broker = OptionMenu(Piv,BrokerCol,*K).grid(row = 2, column = 2)
        
    #Makepivot variables
    Values = ValuesName.get()
    Broker = BrokerCol.get()
    Bank = str(d['e17'].get())
    openalloc = str(d['e18'].get())
    
    # Quarterly variables
    file1 = FirstMonth.get()
    file2 = SecondMonth.get()
    file3 = ThirdMonth.get()
    Advisor = Advisorname.get()
    
    #Admin_File
    sheet_name_admin = str(d['e4'].get())
    Investor_admin = str(d['e7'].get())
    Series_admin = str(d['e9'].get())
    Mgnt_admin = str(d['e11'].get())
    Perf_admin = str(d['e12'].get())
    skiprows_admin = int(d['e13'].get())

    #Advisor_file
    sheet_name_advisor = str(d['e5'].get())
    Investor_advisor = str(d['e8'].get())
    Series_advisor =  str(d['e10'].get())
    Advisor_advisor = 'Fee'
    skiprows_advisor = int(d['e14'].get())

    #Key_file
    sheet_name_key = str(d['e6'].get())
    skiprows_key = int(d['e15'].get())

    #Read_range
    Range = str(d['e16'].get())
    
    #Currency type based on admin file name
    if 'ltd' in file_1.lower():
        currencytype = 'Ltd'
    else:
        currencytype = 'QIHF'
    print("Variables initialised")
    
    if '/' in Advisor:
        fee_sheet = Advisor.replace("/","and") + ' Fees'
    else:
        fee_sheet = Advisor + ' Fees'
    
    AdvisorClean = fee_sheet.split(' Fees')[0]

#Colours headings and formattings 
def colour(df,worksheet,row,workbook,color):
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': color,
    'border': 1, 'font_color':'white'})
    header_format.set_center_across()
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(row, col_num + 1, value, header_format) 

#Return the appropriate formatting based on currency type
def Whatformat(file_1):
    RandFormat = 'R#,##0.00'
    DollarFormat = '$#,##0.00'
    if "ltd" in file_1.lower():
        return DollarFormat
    else:
        return RandFormat
    
#Make pivot function
def makepivot(*args):   
    #RMB file
    dfadmin = pd.read_excel(file_admin)
    adminfile = pd.read_excel(admin,sheet_name=sheet_name_admin,usecols = Range,skiprows=skiprows_admin-1)
    adminfile = adminfile[adminfile[Investor_admin].notnull() & adminfile[Series_admin].notnull()].reset_index(drop=True)
    adminsum = adminfile[adminfile[Investor_admin] == Bank].reset_index(drop = True)
    adminsumtot = adminsum[openalloc].sum()
    
    PivotData = pd.DataFrame(columns = ['Advisor','Total Sum','Percentage RMB'])
    for element in list(dfadmin[Broker].unique()):
        Investor_frame = dfadmin[dfadmin[Broker] == element]
        InvestorSum = Investor_frame[Values].sum()
        TotalSum = dfadmin[Values].sum()
        PivotData = PivotData.append(pd.DataFrame(data = [[element,float(InvestorSum),'{}%'.format(float(InvestorSum/TotalSum * 100))]],
                                        columns = ['Advisor','Total Sum','Percentage RMB']))
    PivotData = PivotData.reset_index(drop = True)
    writer_admin = pd.ExcelWriter('RMB_split_{}.xlsx'.format(file_admin.lower().split('RMB')[0]), engine='xlsxwriter')
    
    PivotData.to_excel(writer_admin,sheet_name = 'RMB split')
    adminsum.to_excel(writer_admin,sheet_name='Admin_RMB')
    
    workbook_admin = writer_admin.book
    worksheet_admin= writer_admin.sheets['RMB split']
    worksheet2_admin = writer_admin.sheets['Admin_RMB']
    perc_format = workbook_admin.add_format({'num_format':'#,##0.00%'})
    formatwrap = workbook_admin.add_format({'text_wrap':True})
    
    worksheet_admin.set_column('A:D',30)
    worksheet_admin.set_column('D:D',30,perc_format)
    worksheet2_admin.set_column('A:Z',30, formatwrap)
    header_format_admin = workbook_admin.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#122057',
    'border': 1, 'font_color':'white'})
    worksheet_admin.write(PivotData.shape[0] + 2, 1,'Previous Total',header_format_admin)
    worksheet_admin.write(PivotData.shape[0] + 2, 2 , PivotData['Total Sum'].sum(),workbook_admin.add_format({'bold':True}))
    worksheet_admin.write(PivotData.shape[0] + 3, 1,'Current Total',header_format_admin)
    worksheet_admin.write(PivotData.shape[0] + 3, 2 , adminsumtot ,workbook_admin.add_format({'bold':True}))
    writer_admin.save()   
    
#Function that merges invoices together (requires editing)
def Merge():  
    print(file1,file2,file3)
    # Instantiate writer 
    writer2 = pd.ExcelWriter(file1.split('.xlsx')[0] + ' Full' + '.xlsx', engine='xlsxwriter')
    # Instantiate book
    workbook2  = writer2.book
    # Formatting of entire book
    formatf = workbook2.add_format({'num_format': Whatformat(file1)}) 
    formatblue = workbook2.add_format({'fg_color': '#122057','font_size':30})
    format_zero = workbook2.add_format({'font_color': 'red'})
    formatperc = workbook2.add_format({'num_format': '#,##0.00%'})
    try:
        df_f1 = pd.read_excel(file1)
        df_tf1 = pd.read_excel(file1, sheet_name = file1.split('31st')[0] + 'Fees')
        df_f1.to_excel(writer2, sheet_name = file1.split('31st')[0])
        worksheet_1 = writer2.sheets[file1.split('31st')[0]]
        worksheet_1.set_column('A:D',35)
        worksheet_1.set_column('E:E',35,formatperc)
        worksheet_1.set_column('F:U', 35, formatf)
        colour(df_f1,worksheet_1,0,workbook2,'#122057')
        worksheet_1.insert_image('A1', 'ps_logo_home.png',{'x_offset':40,'y_offset':5,'x_scale': 0.25,'y_scale': 0.15})
        worksheet_1.conditional_format(1, 4, df_f1.shape[0],19 , {'type':     'cell',
                                          'criteria': '==',
                                          'value':    0,
                                          'format':   format_zero})
        worksheet_1.write('A1','',formatblue)
        df_joined = abs(df_tf1.iloc[[0]])
        MgtFee = [float(df_tf1.iloc[3][0].split('(')[1].split('%')[0])/100]
        PerfFee = [float(df_tf1.iloc[3][1].split('(')[1].split('%')[0])/100]
        print('{} Added as {}'.format(file1 + ' Sheet: ' + Monthname , str(file1).split('31st')[0]))
    except:
        print("ERROR: Please enter an invoice file!")
    
    try:
        df_f2 = pd.read_excel(file2)
        df_tf2 = pd.read_excel(file2, sheet_name = file2.split('31st')[0] + 'Fees')
        df_f2.to_excel(writer2, sheet_name= file2.split('31st')[0])
        worksheet_2 = writer2.sheets[file2.split('31st')[0]]
        colour(df_f2,worksheet_2,0,workbook2,'#122057')
        worksheet_2.set_column('A:D',35)
        worksheet_2.set_column('E:E',35,formatperc)
        worksheet_2.set_column('F:U', 35, formatf)
        worksheet_2.insert_image('A1', 'ps_logo_home.png',{'x_offset':40,'y_offset':5,'x_scale': 0.25,'y_scale': 0.15})
        worksheet_2.conditional_format(1, 4, df_f2.shape[0],19 , {'type':     'cell',
                                          'criteria': '==',
                                          'value':    0,
                                          'format':   format_zermo})
        worksheet_2.write('A1','',formatblue)
        df_joined = pd.concat([abs(df_tf1.iloc[[0]]),abs(df_tf2.iloc[[0]])])
        MgtFee.append(float(df_tf2.iloc[3][0].split('(')[1].split('%')[0])/100)
        PerfFee.append(float(df_tf2.iloc[3][1].split('(')[1].split('%')[0])/100)
        print('{} Added as {}'.format(file2 + ' Sheet: ' + Monthname , str(file2).split('31st')[0]))  
    except:
        print('WARNING: Invoice 2 not loaded')
    
    try: 
        df_f3 = pd.read_excel(file3)
        df_tf3 = pd.read_excel(file3, sheet_name = file3.split('31st')[0] + 'Fees')
        df_f3.to_excel(writer2, sheet_name= file3.split('31st')[0])
        worksheet_3 = writer2.sheets[file3.split('31st')[0]]
        colour(df_f3,worksheet_3,0,workbook2,'#122057')
        worksheet_3.set_column('A:D',35)        
        worksheet_3.set_column('E:E',35,formatperc)
        worksheet_3.set_column('F:U', 35, formatf)
        worksheet_3.insert_image('A1', 'ps_logo_home.png',{'x_offset':40,'y_offset':5,'x_scale': 0.25,'y_scale': 0.15})
        worksheet_3.conditional_format(1, 4, df_f3.shape[0],19 , {'type':     'cell',
                                          'criteria': '==',
                                          'value':    0,
                                          'format':   format_zero})

        worksheet_3.write('A1',None,formatblue)
        df_joined = pd.concat([abs(df_tf1.iloc[[0]]),abs(df_tf2.iloc[[0]]),abs(df_tf3.iloc[[0]])])
        MgtFee.append(float(df_tf3.iloc[3][0].split('(')[1].split('%')[0])/100)
        PerfFee.append(float(df_tf3.iloc[3][1].split('(')[1].split('%')[0])/100)
        print('{} Added as {}'.format(file3 + ' Sheet: ' + Monthname , str(file3).split('31st')[0]))
    except:
        print("WARNING: Invoice 3 not loaded")     
    dfsum = pd.DataFrame([abs(df_joined.sum())],index=['Total'])
    Totals = [abs(df_joined.iloc[: , 2].sum()),abs(df_joined.iloc[: , 3].sum())]
    TotalsFirst = [abs(df_joined.iloc[: , 0].sum()),abs(df_joined.iloc[: , 1].sum())]
    df_portion = pd.DataFrame(data = [[abs(df_joined.iloc[: , 2].sum()),abs(df_joined.iloc[: , 3].sum())]],columns =
                              [df_joined.columns[0] + '\n' + '(' + " , ".join(str(x*100) + '%' for x in MgtFee) + ')', 
                               df_joined.columns[1] + '\n' + '(' + " , ".join(str(x*100) + '%' for x in PerfFee) + ')' ],
                                  index = ['Total payable'])

    df_joined = pd.concat([df_joined,dfsum])
    df_joined.to_excel(writer2, sheet_name = file1.split('31st')[0] + 'Fees')
    df_portion.to_excel(writer2, sheet_name = file1.split('31st')[0] + 'Fees' ,startrow = df_joined.shape[0] + 3)
    worksheet_final = writer2.sheets[file1.split('31st')[0] + 'Fees']
    worksheet_final.set_column('A:E',35,formatf)
    
    #Formattings
    total_format = workbook2.add_format({'bold': True, 'bg_color':'#FF5A34','border': 1,'text_wrap': True,
                                         'num_format': Whatformat(file1)})
    total_format.set_center_across()
    hench_format2 = workbook2.add_format({'bold': True,'text_wrap': True,'border':1,'num_format': Whatformat(file1)})
    
    worksheet_final.write(df_joined.shape[0],1,TotalsFirst[0],hench_format2)
    worksheet_final.write(df_joined.shape[0],2,TotalsFirst[1],hench_format2)
    worksheet_final.write(df_joined.shape[0],3,Totals[0],hench_format2)
    worksheet_final.write(df_joined.shape[0],4,Totals[1],hench_format2)
    worksheet_final.write(df_joined.shape[0] + 4,1, Totals[0], total_format)
    worksheet_final.write(df_joined.shape[0] + 4,2, Totals[1], total_format)
    colour(df_joined,worksheet_final,0,workbook2,'#122057')
    colour(df_portion,worksheet_final,df_joined.shape[0] + 3,workbook2,'#122057')
    worksheet_final.conditional_format(1, 1, df_joined.shape[0]-1,10 , {'type':     'cell',
                                          'criteria': '==',
                                          'value':    0,
                                          'format':   format_zero})
    worksheet_final.write('A1','',formatblue)
    worksheet_final.insert_image('A1', 'ps_logo_home.png',{'x_offset':40,'y_offset':5,'x_scale': 0.25,'y_scale': 0.15})
    writer2.save()
    
# Main function that writes monthly invoices
def Main(*args):
    # If the month name chosen does not match that of the admin file name then print warning
    try:
        if Monthname.split()[1].lower() not in file_1.lower() and Months[file_1.lower().split('31')[1].split('-')[1]].lower()!= Monthname.split()[1].lower():
            print("WARNING: You created {} invoice but Admin File indicates another date".format(Monthname))
    except:
        print('WARNING: File name does not contain a date in the format (d-m-y) or month names!')
    #Attempt to load a master file of that month name else continue to except
    try:
        print("Checking if Master exists for that admin file")
        df_join = pd.read_excel('Master ' + Monthname + ' ' + currencytype + ".xlsx")
        df_join.name = 'Master ' + Monthname + ' ' + currencytype + ".xlsx"
        print("Advisor Reference File {} has been loaded".format(df_join.name))
        # Load Key sheet into dataframe 
        df_key = pd.read_excel(file_3,sheet_name=sheet_name_key,skiprows = skiprows_key-1)
        # Check dataframe of unassigned advisors
        df_fill = df_join[df_join['Advisor'].isnull()].reset_index(drop=True)
        # If it is empty then write the invoice else give warning and don't write
        if df_fill.empty:
            # Try to write else give error that inputs are wrong
            try:
                print('Trying to write invoice for {}'.format(Advisor))
                write(df_key,df_join,Advisor,Mgnt_admin,Perf_admin,file_1,Monthname,RMBSplit,fee_sheet)
            except:
                print(Advisor,Mgnt_admin,Perf_admin,file_1,Monthname,RMBSplit,fee_sheet)
                print("ERROR: Some inputs are invalid in Advanced options or File is not closed!")
        else:
            print("WARNING: Not all advisors have been assigned in {} file!".format(df_join.name))
    except:
        print("Master file does not exist.. created one. Please make sure all advisors have been assigned!")
        
        # Load all three sheets.
        df_admin = pd.read_excel(file_1,sheet_name=sheet_name_admin,usecols = Range,skiprows=skiprows_admin-1)
        df_advisor = pd.read_excel(file_2,sheet_name=sheet_name_advisor,usecols = Range,skiprows = skiprows_advisor-1)
        df_key = pd.read_excel(file_3,sheet_name=sheet_name_key,skiprows = skiprows_key-1)
        
        # Load admin dataframe where investor and series columns are not null
        df_admin = df_admin[df_admin[Investor_admin].notnull() & df_admin[Series_admin].notnull()].reset_index(drop=True)
        df_admin = df_admin.loc[df_admin[Investor_admin]!= str(Investor_admin)]
        
        # Drop duplicate cols
        df_admin.dropna(axis=1,how='all',inplace=True)
        # Rename advisor columns to match those of admin
        df_advisor.rename(columns={Advisor_advisor: 'Advisor',Investor_advisor:Investor_admin,Series_advisor:Series_admin},
                          inplace=True)
        # Select only three columns
        df_advisor = df_advisor[['Advisor',Investor_admin,Series_admin]]

        # Left join: Take items from left table (admin) and (only) matching items from right table (advisor)
        # In this case we take all the columns in admin and join to right table (advisor) on investor,series 
        df_join = pd.merge(df_admin,df_advisor, on=[Investor_admin,Series_admin],how='left',suffixes=(' ',' '))

        # Move Advisor column to left
        df_join = df_join[['Advisor'] + [col for col in df_join.columns if col != 'Advisor']]

        # Remove spaces at beginning and end of column names
        df_join.columns = df_join.columns.str.strip()
        # Multiply percentages by 100
        try:
            df_join['%'] = df_join['%'].apply(lambda x: x)
        except:
            print('ERROR: No percentage column (%) present in admin file!')
        df_key.index = df_key.index.str.strip()
        # Create master writer instance 
        master_writer = pd.ExcelWriter('Master ' + Monthname + ' ' + currencytype + ".xlsx",engine='xlsxwriter')
        # Write the resulting dataframe into master writer with sheet name Advisor Split
        df_join.to_excel(master_writer, sheet_name='Advisor Split')
        # Create workbook 
        master_workbook = master_writer.book
        # Call sheets instance
        master_worksheet = master_writer.sheets['Advisor Split']
        # Dictate workbook format
        master_format = master_workbook.add_format({'num_format': Whatformat(file_1)})
        # Dictate Advisor Split sheet column format (Dollar or Rand)
        master_worksheet.set_column('F:U', 18, master_format)
        # Dictate Advisor Split sheet beginning columns format (Size)
        master_worksheet.set_column('A:F', 30)
        # Colour the workbook in red
        colour(df_join, master_worksheet, 0, master_workbook,'#FF0000') 
        # Save the Master file
        master_writer.save()

def write(df_key,df_join,Advisor,Mgnt_admin,Perf_admin,file_1,Monthname,RMBSplit,fee_sheet):
    # Calculate management and performance fees based on Mgnt Fee and Perf. Fee in Key file
    MngFee = round((1 - 0.5*df_key.loc[Advisor]['Mgnt Fee']),5)
    PerfFee = round((1 - 0.05*df_key.loc[Advisor]['Perf. Fee']),5)
    
    # First dataframe is selecting the advisor you chose
    df1 = df_join[df_join['Advisor'] == Advisor].reset_index(drop = True)

    # RMB Dataframes
    try:
        df4 = pd.read_excel(RmbFileFinal,sheet_name = "RMB split")
        df4 = df4[df4['Advisor'].str.contains(Advisor,na=False)]
        Percentage = float(str(list(df4['Percentage RMB'])[0]).split('%')[0])/100
        df5 = pd.read_excel(RmbFileFinal,sheet_name = "Admin_RMB")
        df5['Advisor'] = Advisor
        df5= df5[['Advisor'] + [col for col in df5.columns if col != 'Advisor']]
        df5[df5.select_dtypes(include=['number']).columns]*=Percentage
        #Rename first dataframe
        df1 = pd.concat([df1,df5],ignore_index=True,sort=False)
    except:
        print("WARNING: No RMB file selected ")
       
    # Second dataframe is summing the management fee and performance fee and putting them as a dataframe with index being
    # The date
    df2 = pd.DataFrame(data = [[abs(df1[Mgnt_admin].sum()),abs(df1[Perf_admin].sum()), abs(MngFee*df1[Mgnt_admin].sum()),
                                abs(PerfFee*df1[Perf_admin].sum())]],
                       columns =['Management Fee Total (excl Vat)','Performance Fee Total (excl Vat)', 
                                 'Management Fee payable (excl Vat)', 
                                 'Performance Fee payable (excl Vat)'],
                       index = [Advisor + ' (' + Monthname + ')'])
    # Third dataframe is multiplying and calculating the percentage payable for the respective advisor with date index too
    df3 = pd.DataFrame(data = [[abs(df2.iloc[: , 2].sum()),
                               abs(df2.iloc[: , 3].sum())]],
                                columns = ['Management Fee payable (' + str(MngFee*100) + '%)',
                                          'Performance Fee payable (' + str(PerfFee*100) + '%)'],index=df2.index)
    

    # Create writer instance e.g "Rosebank 31st August LTD.xlsx" 
    writer = pd.ExcelWriter(AdvisorClean + ' ' + Monthname + ' ' + currencytype + ".xlsx", engine='xlsxwriter')
    
    # Write all previous dataframes to excel, 1 being on date sheet, and 2,3 being on Advisor Fees sheet    
    df1.to_excel(writer, sheet_name= Monthname)
    df2.to_excel(writer, sheet_name= fee_sheet)
    df3.to_excel(writer, sheet_name= fee_sheet, startrow = df2.shape[0] + 3)
    
    # Create workbook instance to allow sheet manipulation
    workbook  = writer.book
    # Call sheets instance
    worksheet1 = writer.sheets[Monthname]
    worksheet2 = writer.sheets[fee_sheet]
    # Formatting of entire workbook 
    format1 = workbook.add_format({'num_format': Whatformat(file_1)})
    totalformat = workbook.add_format({'bold': True, 'bg_color':'#122057','border': 1,'text_wrap': True,
                                       'font_color':'white'})
    totalformat.set_center_across()
    finalformat = workbook.add_format({'bold': True, 'num_format': Whatformat(file_1),
                                       'font_color':'black','bg_color':'#B35909','border': 1})
    finalformat.set_center_across()
    format7 = workbook.add_format({'fg_color': '#122057','font_size':30})
    format8 = workbook.add_format({'num_format': '#,##0.00%'})
    formatzero = workbook.add_format({'font_color': 'red'})
    worksheet2.set_column('A:E',35,format1)
    
    # Prev
    hench_format2 = workbook.add_format({'bold': True,'text_wrap': True,'num_format': Whatformat(file_1),'border':1,
                                         'bg_color':'#FF5A34'})
    hench_format2.set_center_across()
    hench_format3 = workbook.add_format({'bold': True,'text_wrap': True,'num_format': Whatformat(file_1)})
    
    worksheet2.write(df2.shape[0],1,df2.iloc[: , 0],hench_format3)
    worksheet2.write(df2.shape[0],2,df2.iloc[: , 1],hench_format3)
    worksheet2.write(df2.shape[0],3,df2.iloc[: , 2],hench_format3)
    worksheet2.write(df2.shape[0],4,df2.iloc[: , 3],hench_format3)
    worksheet2.write(df2.shape[0]+4,1,df2.iloc[: , 2].sum(), hench_format2)
    worksheet2.write(df2.shape[0]+4,2,df2.iloc[: , 3].sum(), hench_format2)
    
    # Set columns of first date sheet and second Advisor Fees sheet to be (Dollar or Rand and Size)\
    worksheet1.set_column('A:D', 35)
    worksheet1.set_column('E:E', 35,format8)
    worksheet1.set_column('F:U', 35, format1)
    #Insert image
    worksheet1.insert_image('A1', 'ps_logo_home.png',{'x_offset':40,'y_offset':5,'x_scale': 0.25,'y_scale': 0.15})
    worksheet2.insert_image('A1', 'ps_logo_home.png',{'x_offset':40,'y_offset':5,'x_scale': 0.25,'y_scale': 0.15})
    worksheet1.write('A1',None,format7)
    worksheet2.write('A1',None,format7)
    worksheet1.conditional_format(1, 4, df1.shape[0],19 , {'type':     'cell',
                                          'criteria': '==',
                                          'value':    0,
                                          'format':   formatzero})
    worksheet2.conditional_format(1, 1, df2.shape[0]-1,10 , {'type':     'cell',
                                      'criteria': '==',
                                      'value':    0,
                                      'format':   formatzero})
    try:
        worksheet2.write(2*df2.shape[0] + 5, 2,Percentage,format8)
        worksheet2.write(2*df2.shape[0] + 5, 1,'Percentage',totalformat)
    except:
        print("ERROR: No percentage from RMB written")
    # Colour all three dataframes using colour function where wraps, bold, aligns, border, colour can be changed
    colour(df1,worksheet1,0,workbook,'#122057')
    colour(df2,worksheet2,0,workbook,'#122057')
    colour(df3,worksheet2,df2.shape[0] + 3,workbook,'#122057')
    # Save the file
    writer.save()
    # Tell user that invoice has been written
    print("SUCCESS: Invoice written as {}".format(Advisor + ' ' + Monthname + ".xlsx"))
    
if __name__ == '__main__':
    d = {}   
    # Tkinter instance for front GUI
    front = Tk()
    # Size of GUI
    front.minsize(width=50, height=50)
    # Title of GUI
    front.title('Sheet Gen')
    # Allow to not be resizable
    front.resizable(0,0)
    
    # Instantiate master instance which has the title "Advanced options tab"
    master = tk.Toplevel(front)
    master.title('Advanced options tab')
    
    # Instantiate new instance which has title "Quarterly Invoice generator"
    new = tk.Toplevel(front)
    new.title("Invoice Joiner")
    
    # Create front GUI labels and position them appropriately
    Labels_front = ["Admin File","Advisor Reference File","Key Reference File","Advisor Name","Date","RMB Split File"]
    Label(front, text = Labels_front[0]).grid(row = 0, column = 0)
    Label(front, text = Labels_front[1]).grid(row = 0, column = 1)
    Label(front, text = Labels_front[2]).grid(row = 0, column = 2)
    Label(front, text = Labels_front[3]).grid(row = 2, column = 0)
    Label(front, text = Labels_front[4]).grid(row = 2, column = 1)
    Label(front, text = Labels_front[5]).grid(row = 2, column = 2)
    
    # Quarterly dropdown menu bariables
    FirstMonth = StringVar(new)
    SecondMonth = StringVar(new)
    ThirdMonth = StringVar(new)
    Advisornew = StringVar(new)
    
    
    # Pivot instantiate
    Piv = tk.Toplevel(front)
    Piv.title('Pivot Table Maker')
    Piv.minsize(width=50, height=50)
    # Allow to not be resizable
    Piv.resizable(0,0)
    
    # Front dropdown menu variables (Admin, advisor, and key files as well as Date and Advisor name)
    adminvar = StringVar(front)
    advisorvar = StringVar(front)
    keyvar = StringVar(front)
    Date = StringVar(front)
    Advisorname = StringVar(front)
    RMBSplit = StringVar(front)
    
    Labels_pivot = ["RMB File","Values Column name","Broker Corporate column name","Admin file"]
    Label(Piv, text = Labels_pivot[0]).grid(row = 0, column = 0)
    Label(Piv, text = Labels_pivot[1]).grid(row = 0, column = 1)
    Label(Piv, text = Labels_pivot[2]).grid(row = 0, column = 2)
    Label(Piv, text = Labels_pivot[3]).grid(row = 3, column = 0)

    RMB = StringVar(Piv)
    ValuesName = StringVar(Piv)
    BrokerCol = StringVar(Piv)
                                                  
    # Create Advanced Options tab Labels
    Labels = ["Admin Sheet Name","Advisor Sheet Name",
             "Key Sheet Name","Admin investor column name","Advisor investor column name","Admin series column name",
              "Advisor series column name","Admin Management Fee column name","Admin Performance Fee column name",
              "Admin columns start row","Advisor columns start row","Key columns start row","Column range (e.g: A:T)",
             "Make Pivot Investor Name","Opening Allocation Value column name"]
    
    # Create Quarterly tab labels
    Q_Labels = ["Invoice 1","Invoice 2","Invoice 3"]
    
    # Advanced options tab labels and grid placements
    Label(master, text = Labels[0]).grid(row = 2, column = 0)
    Label(master, text = Labels[1]).grid(row = 2, column = 1)
    Label(master, text = Labels[2]).grid(row = 2, column = 2)
    Label(master, text = Labels[3]).grid(row = 4, column = 0)
    Label(master, text = Labels[4]).grid(row = 4, column = 1)
    Label(master, text = Labels[5]).grid(row = 6, column = 0)
    Label(master, text = Labels[6]).grid(row = 6, column = 1)
    Label(master, text = Labels[7]).grid(row = 8, column = 0)
    Label(master, text = Labels[8]).grid(row = 8, column = 1)
    Label(master, text = Labels[9]).grid(row = 10, column = 0)
    Label(master, text = Labels[10]).grid(row = 10, column = 1)
    Label(master, text = Labels[11]).grid(row = 10, column = 2)
    Label(master, text = Labels[12]).grid(row = 12, column = 2)
    Label(master, text = Labels[13]).grid(row = 13, column = 0)
    Label(master, text = Labels[14]).grid(row = 13, column = 1)
    
    # Make dictionary with keys and values being entries
    for i in range (4,19):
        d["e{0}".format(i)] = Entry(master,width = 60)

    d["e4"].grid(row = 3, column = 0)
    d["e5"].grid(row = 3, column = 1)
    d["e6"].grid(row = 3, column = 2)
    d["e7"].grid(row = 5, column = 0)
    d["e8"].grid(row = 5, column = 1)
    d["e9"].grid(row = 7, column = 0)
    d["e10"].grid(row = 7, column = 1)
    d["e11"].grid(row = 9, column = 0)
    d["e12"].grid(row = 9, column = 1)
    d["e13"].grid(row = 11, column = 0)
    d["e14"].grid(row = 11, column = 1)
    d["e15"].grid(row = 11, column = 2)
    d["e16"].grid(row = 13, column = 2)
    d['e17'].grid(row = 14, column = 0)
    d['e18'].grid(row = 14, column = 1)
    
    # Define variable that disables or enables grid in Advanced options tab
    var1 = IntVar()
    
    # Save function that writes to results.txt file the present inputs
    def save():
        f = open('results.txt','w')
        for element in d:
            print(d[element].get())
            f.write(d[element].get() + '\n')
        print("SUCCESS: Results have been saved in results.txt file")
        f.close() 
        
    # Load function that deletes entries that you put in and inserts the ones from results.txt file
    def load(d,state):
        f = open('results.txt','r')
        prev = [line.strip('\n') for line in f]
        counter = 0 
        for element in d:
            d[element].delete(0,END)
            d[element].insert(0,prev[counter])
            # If unticked then state is disabled else enabled
            if var1.get() == 0:
                d[element].config(state=state)
            else:
                d[element].config(state='normal')
            counter+=1
        f.close()
    
    # Will always try and load into the entries the results file when you open the program or else it will return error
    try:
        load(d,'disabled')
    except:
        print('ERROR: results.txt File cannot be found. Please paste it into the current directory!')
    
    # Hide the Advanced options tab 
    master.withdraw()
    # Hide Quarterly tab
    new.withdraw()
    # Hide pivot tab
    Piv.withdraw()
    
    # Create quarterly invoices - Function contains GUI variables, labels, and positions.
    def quarterly():
        # Position the labels
        for i in range (0,3): 
            Label(new, text = Q_Labels[i]).grid(row=0,column=i)
        # Month option menus which take advisorfiles array that is any file with "31st" in it 
        Month1 = OptionMenu(new, FirstMonth,*AdvisorFiles).grid(row=1,column=0)
        Month2 = OptionMenu(new, SecondMonth,*AdvisorFiles).grid(row=1,column=1)
        Month3 = OptionMenu(new, ThirdMonth,*AdvisorFiles).grid(row=1,column=2)
        FirstMonth.set(AdvisorFiles[-1])
        SecondMonth.set(AdvisorFiles[-1])
        ThirdMonth.set(AdvisorFiles[-1])
        # Buttons to call merge function, destroy, and refresh
        b8 = Button(new, text='Merge files',command=lambda: 
                    (initialvalues(**d),Merge()),bg = 'green').grid(row=4,column =1)
        b10 = Button(new, text='Hide Menu',command=lambda: new.withdraw(),bg = 'IndianRed4').grid(row=4,column = 0,sticky=W)
        b11 = Button(new, text = 'Refresh',command = lambda: initialvalues(**d),
                     bg='White').grid(row=4,column=2,sticky=W)
    
    # Advanced options tab buttons
    b0 = Button(master,text="Save inputs", command=lambda: save(),bg = 'DeepSkyBlue3').grid(row=19,column = 1)
    b6 = Button(master, text='Hide Menu', command=lambda: (master.withdraw()),bg='IndianRed2').grid(row=19, column=0, sticky=W, pady=4)
    b2 = Button(master,text='Load previous',command=lambda: load(d,'disabled'),bg='thistle1').grid(row=19,column=2)
    b7 = Checkbutton(master,text='Edit Values?',variable=var1,command=lambda: load(d,'disabled')).grid(row=20,column=1)
    # Front tab buttons
    b1 = Button(front,text='Receive file',command=lambda: (initialvalues(**d),Main(file_1,file_2,file_3,sheet_name_admin,
                                                                                   sheet_name_advisor,sheet_name_key,
                                                                                   Investor_admin,Investor_advisor,Series_admin,
                                                                                   Series_advisor,Mgnt_admin, Perf_admin,skiprows_admin,
                                                                                   skiprows_advisor,skiprows_key,Advisor,Monthname,
                                                                                   Range,currencytype))
                                                                        ,bg='green').grid(row=6,column=1)
    b3 = Button(front, text='Join Invoices',command=lambda: (initialvalues(**d),new.deiconify(),quarterly()),
                bg='Orange').grid(row=8,column=1)
    
    b4 = Button(front, text='Advanced Options',command=lambda: master.deiconify(),bg='Cyan').grid(row=7,column=2)
    b5 = Button(front, text='Quit', command=lambda: (master.destroy(),new.destroy(),front.destroy()),bg ='IndianRed4').grid(row=6, column=0, sticky=W, pady=4)
    b9 = Button(front, text = 'Refresh',command = lambda: initialvalues(**d),bg='White').grid(row=7,column=0,sticky=W)
    b10 = Button(front, text='Make RMB Pivot',command=lambda: Piv.deiconify(),bg='Purple').grid(row=7,column=1)
    
    #Piv buttons
    Generate = Button(Piv, text ="Make Pivot Table",command = lambda:(initialvalues(**d),
                                                                        makepivot(file_admin,Values,Broker,Investor_admin
                                                                                 ,Series_admin,skiprows_admin,Range,Bank))
                      ,bg = 'orange').grid(row = 6, column = 1)
    Ref = Button(Piv, text = "Refresh",command = lambda:(initialvalues(**d)), bg = 'white').grid(row = 6, column = 2)
    Quit = Button(Piv, text = 'Hide Menu',command = lambda: Piv.withdraw(),bg = 'IndianRed4').grid(row = 6, column = 0)


    # Initialise variables before running mainloop
    initialvalues(**d)
    # Run main loop
    mainloop()


# In[ ]:




