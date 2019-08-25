#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
from tkinter import *
import tkinter as tk

#Variables
def initialvalues(**d):
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
    global Date
    global Mgnt_admin
    global Perf_admin
    
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
    Range = str(d['e18'].get())

    #File name
    file_1 = str(d['e1'].get())
    file_2 = str(d['e2'].get())
    file_3 = str(d['e3'].get())
    
    #Advisor and date 
    Advisor = str(d['e16'].get())
    Date = str(d['e17'].get())
#Colours headings
def colour(df,worksheet,row,workbook):
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(row, col_num + 1, value, header_format)   
        
def Merge(**d2):
    global file1
    global file2
    global file3
    global Advisor
    global DateRange
    
    file1 = str(d2['file1'].get())
    file2 = str(d2['file2'].get())
    file3 = str(d2['file3'].get())
    Advisor = str(d2['file4'].get())
    DateRange = str(d2['file5'].get())

    df_f1 = pd.read_excel(file1)
    df_f2 = pd.read_excel(file2)
    df_f3 = pd.read_excel(file3)
    
    df_tf1 = pd.read_excel(file1, sheet_name = Advisor + ' Fees')
    df_tf2 = pd.read_excel(file2, sheet_name = Advisor + ' Fees')
    df_tf3 = pd.read_excel(file3, sheet_name = Advisor + ' Fees')

    writer2 = pd.ExcelWriter(Advisor + ' ' + DateRange + ".xlsx", engine='xlsxwriter')
    workbook2  = writer2.book
    
    formatf = workbook2.add_format({'num_format': '$#,##0.00'})
    
    df_f1.to_excel(writer2, sheet_name=str(file1).strip('.xlsx'))
    df_f2.to_excel(writer2, sheet_name= str(file2).strip('.xlsx'))
    df_f3.to_excel(writer2, sheet_name= str(file3).strip('.xlsx'))
    
    worksheet_1 = writer2.sheets[str(file1).strip('.xlsx')]
    worksheet_2 = writer2.sheets[str(file2).strip('.xlsx')]
    worksheet_3 = writer2.sheets[str(file3).strip('.xlsx')]
    
    colour(df_f1,worksheet_1,0,workbook2)
    colour(df_f2,worksheet_2,0,workbook2)
    colour(df_f3,worksheet_3,0,workbook2)
    
    MgtFee = float(df_tf1.iloc[3][0].split('(')[1].split('%')[0])/100
    PerfFee = float(df_tf1.iloc[3][1].split('(')[1].split('%')[0])/100
        
    df_joined = pd.concat([df_tf1.iloc[[0]],df_tf2.iloc[[0]],df_tf3.iloc[[0]]])
    
    
    worksheet_1.set_column('B:U', 18, formatf)
    worksheet_2.set_column('B:U', 18, formatf)
    worksheet_3.set_column('B:U', 18, formatf)
    dfsum = pd.DataFrame(data = [[df_joined[df_joined.columns[0]].sum(),df_joined[df_joined.columns[1]].sum()]
                                ],columns = df_joined.columns,
                                  index=['Total'])
    df_joined = pd.concat([df_joined,dfsum])
    df_portion = pd.DataFrame(data = [[MgtFee * df_joined[df_joined.columns[0]].sum() , 
                                  PerfFee * df_joined[df_joined.columns[1]].sum()]],columns = [df_joined.columns[0] +
                              ' ' + '(' + str(MgtFee*100) + '%' + ')', df_joined.columns[1] + ' ' + '(' + str(PerfFee*100) + '%' + ')'],
                              index = ['Total payable'])
    df_joined.to_excel(writer2, sheet_name = Advisor + ' Fees')
    df_portion.to_excel(writer2, sheet_name = Advisor + ' Fees',startrow = df_joined.shape[0] + 3)
    worksheet_final = writer2.sheets[Advisor + ' Fees']
    worksheet_final.set_column('A:B',18,formatf)
    colour(df_joined,worksheet_final,0,workbook2)
    colour(df_portion,worksheet_final,df_joined.shape[0] + 3,workbook2)
    writer2.save()
     
def Main(file_1,file_2,file_3,sheet_name_admin,sheet_name_advisor,sheet_name_key,Investor_admin,Investor_advisor,
        Series_admin,Series_advisor,Mgnt_admin, Perf_admin,skiprows_admin,skiprows_advisor,skiprows_key,
         Advisor,Date,Range):
    
    # Load all three sheets.
    df_admin = pd.read_excel(file_1,sheet_name=sheet_name_admin,usecols = Range,skiprows=skiprows_admin-1)
    df_advisor = pd.read_excel(file_2,sheet_name=sheet_name_advisor,usecols = Range,skiprows = skiprows_advisor-1)
    df_key = pd.read_excel(file_3,sheet_name=sheet_name_key,skiprows = skiprows_key-1)

    df_admin = df_admin[df_admin[Investor_admin].notnull() & df_admin[Series_admin].notnull()].reset_index(drop=True)
    df_admin.dropna(axis=1,how='all',inplace=True)
    df_advisor.rename(columns={Advisor_advisor: 'Advisor',Investor_advisor:Investor_admin,Series_advisor:Series_admin},
                      inplace=True)

    #Select cols
    df_advisor = df_advisor[['Advisor',Investor_admin,Series_admin]]

    # Left join: Take items from left table (admin) and (only) matching items from rght table (advisor)
    # In this case we take all the columns in admin and join to right table (advisor) on investor,series 
    df_join = pd.merge(df_admin,df_advisor, on=[Investor_admin,Series_admin],how='left',suffixes=(' ',' '))

    #Move to left
    df_join = df_join[['Advisor'] + [col for col in df_join.columns if col != 'Advisor']]

    # #Remove spaces at beggining and end of column names
    df_join.columns = df_join.columns.str.strip()
    df_key.index = df_key.index.str.strip()

    df_fill = df_join[df_join['Advisor'].isnull()].reset_index(drop=True)

    #Display database
    MngFee = round((1 - 0.5*df_key.loc[Advisor]['Mgnt Fee']),5)
    PerfFee = round((1 - 0.05*df_key.loc[Advisor]['Perf. Fee']),5)

    df1 = df_join[df_join['Advisor'] == Advisor].reset_index(drop = True)

    df2 = pd.DataFrame(data = [[df1[Mgnt_admin].sum(),df1[Perf_admin].sum()]],
                       columns =['Management Fee Total (excl Vat)','Performance Fee Total (excl Vat)'],index = [Date])

    df3 = pd.DataFrame(data = [[MngFee*df1[Mgnt_admin].sum(),
                               PerfFee*df1[Perf_admin].sum()]],
                                columns = ['Management Fee payable (' + str(MngFee*100) +'%) excl Vat', 
                                           'Performance Fee payable (' + str(PerfFee*100) + '%) excl Vat'],index=df2.index)

    writer = pd.ExcelWriter(Advisor + ' ' + Date + ".xlsx", engine='xlsxwriter')

    df1.to_excel(writer, sheet_name=Date)
    df2.to_excel(writer, sheet_name= Advisor + ' Fees')
    df3.to_excel(writer, sheet_name= Advisor + ' Fees', startrow = df2.shape[0] + 3)

    workbook  = writer.book
    worksheet1 = writer.sheets[Date]
    worksheet2 = writer.sheets[Advisor + ' Fees']

    #formatting
    format1 = workbook.add_format({'num_format': '$#,##0.00'})

    worksheet1.set_column('B:U', 18, format1)
    worksheet2.set_column('B:C', 18, format1)

    colour(df1,worksheet1,0,workbook)
    colour(df2,worksheet2,0,workbook)
    colour(df3,worksheet2,df2.shape[0] + 3,workbook)
    writer.save()
    
if __name__ == '__main__':
    master = Tk()
    master.minsize(width=400, height=100)
    master.title('Advisor monthly sheet generator')
    
    Labels = ["Admin File Name","Advisor File Name","Key File Name","Admin Sheet Name","Advisor Sheet Name",
             "Key Sheet Name","Admin investor column name","Advisor investor column name","Admin series column name",
              "Advisor series column name","Admin Management Fee column name","Admin Performance Fee column name",
              "Admin columns start row","Advisor columns start row","Key columns start row","Advisor name",
              "Date","Column range (e.g: A:T)"]
    
    Labels2 = ["File1","File2","File3","Advisor","DateRange"]
    d = {}
    d2 = {}
 
    Label(master, text = Labels[0]).grid(row = 0, column = 0)
    Label(master, text = Labels[1]).grid(row = 0, column = 1)
    Label(master, text = Labels[2]).grid(row = 0, column = 2)
    Label(master, text = Labels[3]).grid(row = 2, column = 0)
    Label(master, text = Labels[4]).grid(row = 2, column = 1)
    Label(master, text = Labels[5]).grid(row = 2, column = 2)
    Label(master, text = Labels[6]).grid(row = 4, column = 0)
    Label(master, text = Labels[7]).grid(row = 4, column = 1)
    Label(master, text = Labels[8]).grid(row = 6, column = 0)
    Label(master, text = Labels[9]).grid(row = 6, column = 1)
    Label(master, text = Labels[10]).grid(row = 8, column = 0)
    Label(master, text = Labels[11]).grid(row = 8, column = 1)
    Label(master, text = Labels[12]).grid(row = 10, column = 0)
    Label(master, text = Labels[13]).grid(row = 10, column = 1)
    Label(master, text = Labels[14]).grid(row = 10, column = 2)
    Label(master, text = Labels[15]).grid(row = 12, column = 0)
    Label(master, text = Labels[16]).grid(row = 12, column = 1)
    Label(master, text = Labels[17]).grid(row = 12, column = 2)
    for i in range (1,19):
        d["e{0}".format(i)] = Entry(master,width = 60)
    
    d["e1"].grid(row = 1, column = 0)
    d["e2"].grid(row = 1, column = 1)
    d["e3"].grid(row = 1, column = 2)
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
    d["e16"].grid(row = 13, column = 0)
    d["e17"].grid(row = 13, column = 1)
    d["e18"].grid(row = 13, column = 2)
            
    var1 = IntVar()
    def save():
        if var1.get()==1:
            f = open('results.txt','w')
            for element in d:
                print(d[element].get())
                f.write(d[element].get() + '\n')
            f.close() 
        else:
            for element in d:
                d[element].insert(0,"")        
    def load():
        f = open('results.txt','r')
        prev = [line.strip('\n') for line in f]
        counter = 0 
        for element in d:
            d[element].delete(0,END)
            d[element].insert(0,prev[counter])
            counter+=1
        f.close()
        
    def window():
        new = tk.Toplevel(master)
        for i in range (0,len(Labels2)):
            Label(new, text = Labels2[i]).grid(row=i+1)
        for i in range (1,len(Labels2)+1):
            d2["file{0}".format(i)] = Entry(new,width = 100)
        for i in range (1,len(Labels2)+1):
            d2["file{0}".format(i)].grid(row=i,column=1)
        d2["file1"].insert(0,"Rosebank 31st August.xlsx")
        d2["file2"].insert(0,"Rosebank 31st July.xlsx")
        d2["file3"].insert(0,"Rosebank 31st Sep.xlsx")
        d2["file4"].insert(0,"Rosebank")
        d2["file5"].insert(0,"July-Aug")
        b4 = Button(new, text='Merge files',command=lambda: Merge(**d2)).grid(row=7,column =1)
        
    b0 = Checkbutton(master,text="Save inputs?",variable=var1,command=lambda: save()).grid(row=20,column = 1)
    b1 = Button(master,text='Receive file',command=lambda: (initialvalues(**d)
                                                            ,Main(file_1,file_2,file_3,sheet_name_admin,
                                                                  sheet_name_advisor,sheet_name_key,Investor_admin,
                                                                  Investor_advisor,Series_admin,Series_advisor,
                                                                  Mgnt_admin,Perf_admin,
                                                                  skiprows_admin,skiprows_advisor,
                                                                  skiprows_key,Advisor,Date,Range))).grid(row=19,column=1)
    b2 = Button(master,text='Load previous',command=lambda: load()).grid(row=19,column=2)
    b3 = Button(master, text='Quarterly?',command=lambda: window()).grid(row=19,column=3)
    
    Button(master, text='Quit', command=lambda: (master.destroy())).grid(row=19, column=0, sticky=W, pady=4)
    mainloop()

