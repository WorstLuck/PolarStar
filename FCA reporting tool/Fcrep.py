#!/usr/bin/env python
# coding: utf-8

# In[11]:


import pandas as pd
import tkinter as tk
from tkinter import *
import os
import numpy as np
import requests
from bs4 import BeautifulSoup
import sys
import datetime
import io


# In[103]:



# Initialized variables
def initialvalues():
    global date
    global File
    global Sheet
    global key_file
    global NAV
    global CashU
    global OTC_Margin
    global keys
    global Cd
    global Banks
    global key
    
    Files = [element for element in os.listdir() if 'xlsx' in element.lower() or 'xls' in element.lower()]
    Op1 = OptionMenu(front,Dump,*Files).grid(row=1,column=0)
    Op2 = OptionMenu(front,Keys,*Files).grid(row=3,column=0)
    
    File = Dump.get()
    key_file = Keys.get()
    Sheet = str(d['e1'].get())
    date = str(d['e2'].get())
    NAV = float(d['e3'].get().replace(',',''))
    CashU = float(d['e4'].get().replace(',',''))
    OTC_Margin = float(d['e5'].get().replace(',',''))

    # Key dataset
    keys ={'Safex':'XSAF','CME':'XCME','Cbot':'XCBT','Nymex':'NYMX','Nybot':'XNYF','ICE':'IEPA','Liffe':'XLIF'}
    key = pd.DataFrame(list(keys.items()),columns=['Exchange','Code'])
    key.set_index(['Exchange'],inplace=True)

    #Codes to banks
    Cd = {'0PSF1':'SG','0PSF2':'SG','FKW642':'RMB','FUF999':'RMB','PSFL01':'Macq','PSMSEZCFGC':'Macq','RJO 40020':'RJO'}
    Banks = ['RJO','JPM','Macq','SG','RMB','ABSA']


# Return EUR to Dollar rates
def GrabRate(date):
    global rate
    
    # Building blocks for the URL
    entrypoint = 'https://sdw-wsrest.ecb.europa.eu/service/' # Using protocol 'https'
    resource = 'data'           # The resource for data queries is always'data'
    flowRef ='EXR'              # Dataflow describing the data that needs to be returned, exchange rates in this case
    key = 'D.USD.EUR.SP00.A'    # Defining the dimension values, explained below
    
    request_url = entrypoint + resource + '/'+ flowRef + '/' + key
    
    # Define the parameters
    parameters = {'startPeriod': date, 'endPeriod': date}
    
    # Make the HTTP request again, now requesting for CSV format
    response = requests.get(request_url, params=parameters, headers={'Accept': 'text/csv'})

    # Read the response as a file into a Pandas DataFrame
    df = pd.read_csv(io.StringIO(response.text))
    
    rate = float(df['OBS_VALUE'].values[0])
    return print(rate)

def main():
    GrabRate(date)

    Sheet1 = pd.read_excel(File,sheet_name = Sheet)
    codes = pd.read_excel(key_file)

    codes.set_index(['Entity','Name'],inplace=True)

    Ex1 = Sheet1.groupby(['Exchange','Commodity']).sum().round(1)
    Mg1 = Sheet1.groupby(['Manager Category','Commodity']).sum().round(1)
    Fd1 = Sheet1.groupby(['Contract']).sum().round(1)
    Fd2 = Sheet1.groupby(['Contract','Commodity','Exchange','Manager Category','Description','Last Trade Date'],as_index=False).sum().round(1)
    Fd3 = Sheet1.groupby(['Geo Region','Commodity']).sum().round(1)
    Fd4 = Sheet1.groupby(['Fund Category','Commodity']).sum().round(1)
    Fd6 = Sheet1.groupby(['Description','Commodity']).sum().round(1)
    Fd7 = Sheet1.groupby(['Account Code','Commodity']).sum().round(1)
    Fd8 = Sheet1.groupby(['Commodity']).sum().round(1)
    Fd12 = Sheet1.groupby(['Manager Category','Commodity']).sum().reset_index()


    Fd12 = Fd12[['Manager Category','Commodity','Nominal USD']]
    Fd12['ABS'] = Fd12['Nominal USD'].apply(lambda x: abs(x))
    Fd12['LONG'] = Fd12['Nominal USD'].apply(lambda x: None if x < 0 else x)
    Fd12['SHORT'] = Fd12['Nominal USD'].apply(lambda x: None if x > 0 else x)
    Fd12.set_index(['Manager Category','Commodity'],inplace=True)

    Indices = pd.DataFrame(0,columns=['Nominal USD','ABS','LONG','SHORT'],index=['Other Commodities/ Agricultural Products','Other Commodities/ Livestock',
               'Other Structured/securitised products','Other cash and cash equivalents (excluding governement securities)'
              ,'Other commodities/Industrial metals','Other Commodities/ Other','Precious metals/Gold','Precious Metals/Other'
              ,'Energy/ Crude Oil','Energy/Natural gas'])

    Indices.index.name='Manager Category'
    Indices = Indices.reset_index()

    Fd12Sum = Fd12.groupby(['Manager Category']).sum()
    Fd12Sum = Fd12Sum.reset_index()

    Fd12Final = pd.merge(Indices,Fd12Sum,on='Manager Category',how='outer',suffixes=('_',''))
    Fd12Final.drop([x for x in Fd12Final if x.endswith('_')],axis=1,inplace=True)

    Fd12Final.set_index(['Manager Category'],inplace=True)
    Fd12Sum.set_index(['Manager Category'],inplace=True)


    #Some variables for INDEXP
    TbillL = Fd12.xs('TBILL',level=1, drop_level=False)['LONG'].sum()
    TbillS = abs(Fd12.xs('TBILL',level=1, drop_level=False)['SHORT'].sum())
    MACQL = Fd12.xs('MACQ_RPI',level=1, drop_level=False)['LONG'].sum()
    MACQS = abs(Fd12.xs('MACQ_RPI',level=1, drop_level=False)['SHORT'].sum())
    ok = Fd12Sum.loc['Other cash and cash equivalents (excluding governement securities)']['ABS'] - TbillL
    E129L = Fd12Sum.loc['Other Commodities/ Livestock']['LONG']
    E129S = abs(Fd12Sum.loc['Other Commodities/ Livestock']['SHORT'])
    E130L = Fd12Sum.loc['Other Commodities/ Agricultural Products']['LONG']
    E130S = abs(Fd12Sum.loc['Other Commodities/ Agricultural Products']['SHORT'])

    MASTER = pd.DataFrame(data = [[TbillL,TbillS],[MACQL,MACQS],[ok,None],[E129L,E129S],[E130L,E130S]],
                          columns = ['LONG','SHORT'],index=[77,107,120,129,130])

    #END 
    Fd11 = Sheet1.loc[:, ['Geo Region','Nominal USD']]
    Fd11H = Sheet1.loc[:, ['CONCA','Geo Region','Nominal USD']]
    Fd11H['CONCA'] = Fd11H['CONCA'].apply(lambda x: x if 'hedge' in x.lower() else None)
    Fd11H = Fd11H.dropna(axis=0)
    Fd11H['LONG (USD)'] = Fd11H['Nominal USD'].apply(lambda x: None if x < 0 else x)
    Fd11H['SHORT (USD)'] = Fd11H['Nominal USD'].apply(lambda x: None if x > 0 else x)
    Fd11H.set_index(['Geo Region','CONCA'],inplace=True)
    Fd11H = Fd11H.sort_values(by='Geo Region')

    Fd11HF = Fd11H.groupby(['Geo Region']).sum()
    Fd11['LONG (USD)'] = Fd11['Nominal USD'].apply(lambda x: None if x < 0 else x)
    Fd11['SHORT (USD)'] = Fd11['Nominal USD'].apply(lambda x: None if x > 0 else x)
    Fd11 = Fd11.groupby(['Geo Region']).sum().round(1)

    Fd11 = Fd11.drop(['Nominal USD'],axis=1)
    Fd11H = Fd11H.drop(['Nominal USD'],axis=1)
    Fd11HF = Fd11HF.drop(['Nominal USD'],axis=1)

    Fd10 = pd.DataFrame(0,index=Banks,columns=['Value','%'])

    Fd9 = Sheet1[['Identifier']]
    Fd9 = Fd9.Identifier.unique().tolist()
    Fd9 = pd.DataFrame(Fd9,columns=['Identifier'])
    Fd9 = Fd9.sort_values(by=['Identifier'],ascending=True).reset_index(drop=True)


    Fd8 = Fd8[['Nominal USD']]
    Fd8['ABS'] = Fd8['Nominal USD'].apply(lambda x: abs(x))
    Fd8.drop('TBILL',inplace=True)
    Fd8Sum = Fd8['ABS'].sum()

    Fd7 = Fd7[['Nominal USD']]
    Fd7['ABS'] = Fd7['Nominal USD'].apply(lambda x: abs(x))
    Fd7Sum = Fd7.groupby(['Account Code']).sum().round(1).reset_index()
    Fd7Sum['Bank'] = Fd7Sum['Account Code'].apply(lambda x: Cd[x])
    Fd7Final = Fd7Sum.groupby(['Bank']).sum().round(1).sort_values(by=['ABS'],ascending=False)
    Fd7Final['Cash at Bank'] = 0
    Fd7Final['%'] = 0
    Fd7Final['Difference'] = 0

    Fd7Sum.set_index(['Bank','Account Code'],inplace=True)

    # Conc princip

    Fd4 = Fd4[['Nominal USD']]
    Fd4['ABS'] = Fd4['Nominal USD'].apply(lambda x: abs(x))
    Fd4['LONGS'] = Fd4['Nominal USD'].apply(lambda x: None if x < 0 else x)
    Fd4['% LONGS'] = Fd4['LONGS']/Fd4['ABS'].sum()
    Fd4['SHORTS'] = Fd4['Nominal USD'].apply(lambda x: None if x > 0 else x)
    Fd4['% SHORTS'] = abs(Fd4['SHORTS']/Fd4['ABS'].sum())

    #RHS UP data set
    Fd4Sum = Fd4.reset_index()
    Fd4Sum = Fd4Sum.groupby(['Fund Category']).sum().sort_values(by=['ABS'],ascending=False)


    # Reindex for RHS down
    Fd4Final = Fd4Sum.reset_index()

    #RHS down dataset
    Fd4Final = Fd4Final.melt(id_vars=["Fund Category","Nominal USD","ABS","LONGS",'SHORTS'], 
            var_name="Type", 
            value_name="%").sort_values(by='%',ascending=False)
    Fd4Final['Type'] = Fd4Final['Type'].apply(lambda x: x.split('%')[1])

    Fd4Final.set_index(['Fund Category'],inplace=True)
    Fd4Final = Fd4Final[['%','Type']]
    Fd4Final['NET USD'] = Fd4Final['%'] * Fd4['ABS'].sum()

    # END 

    Fd3 = Fd3[['Nominal USD']]
    Fd3['ABS'] = Fd3['Nominal USD'].apply(lambda x: abs(x))
    Fd3Sum = Fd3.groupby(['Geo Region']).sum().round(1)
    Fd3Sum['% Aggregate Asset Value'] = Fd3Sum['ABS']/Fd3Sum['ABS'].sum()
    Fd3Sum['Net'] = Fd3Sum['Nominal USD']
    Fd3Sum['% Portion of NAV'] = Fd3Sum['Net']/NAV
    Fd3Sum['% of NAV to 100%'] = Fd3Sum['% Portion of NAV']/Fd3Sum['% Portion of NAV'].sum()

    Fd2['Type'] = Fd2['Nominal USD'].apply(lambda x: 'SHORT' if x < 0 else 'LONG')
    Fd2['ABS'] = Fd2['Nominal USD'].apply(lambda x: abs(x))

    Fd2 = Fd2[['Contract','Commodity','Exchange','Manager Category','Description','Last Trade Date','ABS','Type']]
    Fd2 = Fd2.sort_values(by=['ABS'],ascending=False)
    Fd2.set_index(['Contract'],inplace=True)

    Fd1 = Fd1[['Nominal USD']]
    Fd1['ABS'] = Fd1['Nominal USD'].apply(lambda x: abs(x))

    #Borrow R GrossExp sheet:
    Fd6 = Fd6[['Nominal USD']]
    Fd6['ABS'] = Fd6['Nominal USD'].apply(lambda x: abs(x))
    Fd6Sum = Fd6.groupby(['Description']).sum().round(1)
    Fd6Com = Fd6.reset_index(True)

    OTC = Fd6Com[(Fd6Com['Commodity'] == 'QCOTC')|(Fd6Com['Commodity'] == 'CCOTC') 
                 |(Fd6Com['Commodity'] == 'SYOTC')]

    Fd6.drop('TBILL',level=1,inplace=True)
    Fd6.drop('QCOTC',level=1,inplace=True)
    Fd6.drop('CCOTC',level=1,inplace=True)
    Fd6.drop('SYOTC',level=1,inplace=True)

    Fd6FinalSum = Fd6.groupby(['Description'])
    DervGross = Fd6['ABS'].sum()
    Margin = NAV*CashU - OTC_Margin
    Difference = DervGross - Margin
    Percentage = Difference*100/NAV
    OTC = OTC['ABS'].sum()
    Difference_2 = OTC - OTC_Margin
    Percentage_2 = Difference_2*100/NAV
    Gross2 = (DervGross + OTC)*100/NAV

    ETD = pd.DataFrame([DervGross,Margin,Difference,NAV,Percentage,OTC,OTC_Margin,Difference_2,Percentage_2,Gross2],columns = ['Value']
                      ,index = ['DervGross','Margin','Difference',
                                'NAV','Percentage','OTC','OTC Margin','Difference','Percentage','Gross'])

    Mg1Sum = Mg1[['Nominal USD']]
    Mg1Sum = Mg1Sum.rename(columns = {'Nominal USD':'Total'})
    Mg1Sum['ABS'] = Mg1Sum['Total'].apply(lambda x: abs(x))
    Mg1FinalSum = Mg1Sum.groupby(['Manager Category']).sum().round(1)
    Mg1FinalSum = Mg1FinalSum.rename(columns = {'ABS':'NET USD'})
    Mg1FinalSum['NET EUR'] = Mg1FinalSum['NET USD']*(1/rate)


    FinalSum = Ex1[['Nominal USD']]

    Ex1.drop('MACQ_RPI',level=1,inplace=True)
    Ex1.drop('TBILL',level=1,inplace=True)
    Ex1.drop('QCOTC',level=1,inplace=True)
    Ex1.drop('CCOTC',level=1,inplace=True)
    Ex1.drop('SYOTC',level=1,inplace=True)

    #OTC's
    Ex2 = Sheet1.groupby(['Commodity'],as_index=False).sum().round(1)
    Ex2 = Ex2[Ex2['Commodity'].str.contains('OTC')]
    Ex2['Exchange'] = 'OTC'
    Ex2.set_index(['Exchange','Commodity'],inplace=True)

    Ex = pd.concat([Ex1,Ex2],ignore_index=False)

    Final = Ex[['Nominal USD']]

    FinalPlus = Final[Ex.groupby('Nominal USD')['Nominal USD'].transform(lambda x: x > 0)].reset_index(True)
    FinalMinus = Final[Ex.groupby('Nominal USD')['Nominal USD'].transform(lambda x: x < 0)].reset_index(True)

    Final3 = pd.merge(FinalPlus,FinalMinus, on=['Exchange','Commodity'],how='outer',suffixes = (' +',' -'))
    Final3 = Final3.fillna(0)
    Final4 = Final3.groupby(['Exchange']).sum().round(1)
    Final4['Nominal USD -'] = Final4['Nominal USD -'].apply(lambda x: abs(x))

    Tot = Final4['Nominal USD +'].sum() + Final4['Nominal USD -'].sum()

    Final4['+%'] = Final4['Nominal USD +']*100/(Tot)
    Final4['-%'] = Final4['Nominal USD -']*100/(Tot)

    Final4['NET USD'] = Final4['Nominal USD +'] + Final4['Nominal USD -']

    # Conmost NB
    FdNB = Final4.reset_index()
    FdNB = FdNB.rename(columns = {'Nominal USD +':'LONGS','Nominal USD -': 'SHORTS','+%':'LONG','-%':'SHORT'})


    FdNB['Max USD'] = FdNB[["LONGS", "SHORTS"]].max(axis=1)

    FdNB = FdNB.melt(id_vars=["Exchange","LONGS","SHORTS","NET USD",'Max USD'], 
            var_name="Type", 
            value_name="Sum").sort_values('Sum',ascending=False)

    FdNB['Sum'] = FdNB['Sum']/100

    FdNB.set_index(['Exchange'],inplace=True)


    #END

    Final4['NET EUR'] = Final4['NET USD'] * (1/rate)

    FinalSum = FinalSum.rename(columns = {'Nominal USD':'Total'})
    FinalSum['ABS'] = FinalSum['Total'].apply(lambda x: abs(x))


    # To excel
    Writer = pd.ExcelWriter('Split.xlsx',engine='xlsxwriter')

    FinalSum.to_excel(Writer,sheet_name = 'Mg Rep - Ex BDown')
    Final4.to_excel(Writer,sheet_name = 'Mg Rep - Ex BDown',startcol= FinalSum.shape[1] + 3)

    Mg1Sum.to_excel(Writer,sheet_name = 'Mg Rep - Instr')
    Mg1FinalSum.to_excel(Writer,sheet_name = 'Mg Rep - Instr',startcol= Mg1Sum.shape[1] + 3)

    codes.to_excel(Writer,sheet_name = 'Fund Rep - Prime Broker')

    Fd1.to_excel(Writer,sheet_name = 'FundRep- Concentra',startrow = 10)
    Fd2.to_excel(Writer,sheet_name = 'FundRep- Concentra',startrow = 10,startcol= Fd1.shape[1] + 2)
    key.to_excel(Writer,sheet_name = 'FundRep- Concentra')

    Fd3.to_excel(Writer,sheet_name = 'FundRep-ConcenGeog')
    Fd3Sum.to_excel(Writer,sheet_name = 'FundRep-ConcenGeog',startcol= Fd3.shape[1] + 3)

    Fd4.to_excel(Writer,sheet_name = 'FundRep-ConcPrincipExp')
    Fd4Sum.to_excel(Writer,sheet_name = 'FundRep-ConcPrincipExp',startrow= Fd4.shape[0] + 3,startcol = 1)
    Fd4Final.to_excel(Writer,sheet_name = 'FundRep-ConcPrincipExp',startcol = Fd4.shape[1] + 3)

    FdNB.to_excel(Writer,sheet_name = 'FundRep-ConMost NB',startrow = 10)
    key.to_excel(Writer,sheet_name = 'FundRep-ConMost NB')

    Fd12.to_excel(Writer,sheet_name='FundRep-IndExp')
    Fd12Final.to_excel(Writer,sheet_name='FundRep-IndExp',startrow=Fd12.shape[0]+3,startcol=1)
    MASTER.to_excel(Writer,sheet_name='FundRep-IndExp',startcol=Fd12.shape[1]+3)

    Fd11.to_excel(Writer,sheet_name = 'Fund Rep - Currency Exp')
    Fd11H.to_excel(Writer,sheet_name = 'Fund Rep - Currency Exp',startcol=Fd11.shape[1]+3)
    Fd11HF.to_excel(Writer,sheet_name = 'Fund Rep - Currency Exp',startrow = Fd11H.shape[0]+3,startcol=Fd11.shape[1]+3)

    Fd10.to_excel(Writer, sheet_name='Fund Rep - Counter Risk',startrow=2)
    codes.to_excel(Writer, sheet_name='Fund Rep - Counter Risk',startcol=Fd10.shape[1] + 3)

    Fd6.to_excel(Writer,sheet_name = 'FundRep-BorrowR GrossExp')
    Fd6Sum.to_excel(Writer,sheet_name = 'FundRep-BorrowR GrossExp',startcol= Fd6.shape[1] + 3)
    ETD.to_excel(Writer,sheet_name = 'FundRep-BorrowR GrossExp',startcol= Fd6Sum.shape[1] + Fd6.shape[1] + 6 )

    Fd8Final = pd.DataFrame(data = [Fd8Sum.round(2),NAV,Gross2],columns=['Values'],index=['Commitment Exposure of AIF',
                                                                                          'NAV','Gross'])
    Fd8Final.to_excel(Writer,sheet_name = 'FundRep-BorrowR CommitExp',startrow=2,startcol=Fd8.shape[1] + 3) 
    Fd8.to_excel(Writer,sheet_name = 'FundRep-BorrowR CommitExp')

    Fd9.to_excel(Writer, sheet_name='FundRep-Op R#ofPosi')

    Fd7.to_excel(Writer,sheet_name = 'FundRep-BorrowSource')
    Fd7Sum.to_excel(Writer,sheet_name = 'FundRep-BorrowSource',startcol= Fd7.shape[1] + 3)
    Fd7Final.to_excel(Writer,sheet_name = 'FundRep-BorrowSource',startrow = Fd7Sum.shape[0]+3,startcol= Fd7.shape[1] + 3)
    codes.to_excel(Writer, sheet_name='FundRep-BorrowSource', startrow = 
                   Fd7Sum.shape[0] + Fd7Final.shape[0] + 6, startcol = Fd7.shape[1] + 3)

    #Formatting
    Workbook = Writer.book
    Worksheet = Writer.sheets['Mg Rep - Ex BDown']
    Worksheet2 = Writer.sheets['Mg Rep - Instr']
    Worksheet3 = Writer.sheets['FundRep- Concentra']
    Worksheet4 = Writer.sheets['FundRep-ConcenGeog']
    Worksheet5 = Writer.sheets['FundRep-ConcPrincipExp']
    Worksheet6 = Writer.sheets['FundRep-ConMost NB']
    Worksheet7 = Writer.sheets['Fund Rep - Currency Exp']
    Worksheet8 = Writer.sheets['FundRep-BorrowR GrossExp']
    Worksheet9 = Writer.sheets['FundRep-BorrowSource']
    Worksheet10 = Writer.sheets['FundRep-BorrowR CommitExp']
    Worksheet11 = Writer.sheets['FundRep-Op R#ofPosi']
    Worksheet12 = Writer.sheets['Fund Rep - Counter Risk']
    Worksheet13 = Writer.sheets['FundRep-IndExp']

    codes = Writer.sheets['Fund Rep - Prime Broker']

    Format = Workbook.add_format({'num_format':'#,##0.00'})
    Format2 = Workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#122057',
        'border': 1, 'font_color':'white','num_format':'#,##0.00'})
    Format3 = Workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#ff0000',
        'border': 1, 'font_color':'white','num_format':'#,##0.00'})
    Format3.set_center_across()

    PercForm = Workbook.add_format({'num_format':'0.00%'})
    PercForm2 = Workbook.add_format({'bold':True, 'num_format':'0.00%'})
    PercForm2.set_center_across()
    Dollars = Workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#122057',
        'border': 1, 'font_color':'white','num_format':'$#,##0.00'})

    Worksheet.set_column('A:L',20,Format)
    Worksheet.set_tab_color('#00B050')

    Worksheet2.set_column('B:I',20,Format)
    Worksheet2.set_column('A:A',60,Format)
    Worksheet2.set_column('F:F',60,Format)
    Worksheet2.set_tab_color('#00B050')

    Worksheet3.set_column('A:L',20,Format)
    Worksheet3.set_column('H:H',60,Format)
    Worksheet3.set_tab_color('#00B050')
    Worksheet3.set_zoom(85)

    Worksheet4.set_column('A:L',20,Format)
    Worksheet4.set_column('I:I',30,PercForm)
    Worksheet4.set_column('K:L',30,PercForm)
    Worksheet4.set_zoom(85)
    Worksheet4.set_tab_color('#00B050')

    Worksheet5.set_column('A:B',40,Format)
    Worksheet5.set_column('C:H',20,Format)
    Worksheet5.set_column('J:J',40,Format)
    Worksheet5.set_column('L:M',20,Format)
    Worksheet5.set_column('F:F',20,PercForm)
    Worksheet5.set_column('H:H',20,PercForm)
    Worksheet5.set_column('K:K',15,PercForm)
    Worksheet5.set_zoom(85)
    Worksheet5.set_tab_color('#00B050')

    Worksheet6.set_column('A:L',20,Format)
    Worksheet6.set_column('G:G',20,PercForm)
    Worksheet6.set_tab_color('#00B050')

    Worksheet7.set_column('A:I',20,Format)
    Worksheet7.set_column('G:G',35,Format)
    Worksheet7.set_tab_color('#00B050')

    Worksheet8.set_column('A:I',20,Format)
    Worksheet8.set_column('K:M',20,Format)
    Worksheet8.set_tab_color('#00B050')
    Worksheet8.write(5,Fd6Sum.shape[1] + Fd3.shape[1] + 7,str(Percentage.round(2)) + '%',PercForm2)
    Worksheet8.write(9,Fd6Sum.shape[1] + Fd3.shape[1] + 7,str(Percentage_2.round(2)) + '%',PercForm2)
    Worksheet8.write(10,Fd6Sum.shape[1] + Fd3.shape[1] + 7,str(Gross2.round(2)) + '%',PercForm2)

    Worksheet9.set_column('A:D',20,Format)
    Worksheet9.set_column('F:H',35,Format)
    Worksheet9.set_column('I:K',20, Format)

    Worksheet9.set_tab_color('#00B050')

    Worksheet10.set_column('A:H',20,Format)
    Worksheet10.set_column('F:F',35,Format)
    Worksheet10.set_tab_color('#00B050')

    Worksheet11.set_column('A:H',20,Format)
    Worksheet11.set_tab_color('#00B050')

    Worksheet12.set_column('A:C',20,Format)
    Worksheet12.set_tab_color('#00B050')
    Worksheet12.set_column('F:I',35,Format)

    Worksheet13.set_column('A:B',60,Format)
    Worksheet13.set_column('C:F',20,Format)
    Worksheet13.set_column('H:J',20,Format)
    Worksheet13.set_zoom(85)
    Worksheet13.set_tab_color('#ff0000')

    codes.set_column('A:L',35,Format)
    codes.set_tab_color('#00B050')

    Worksheet.write(Final4.shape[0] + 3,FinalSum.shape[1] + Final4.shape[1] + 3,'1:' + str(rate) + ' or ' + str(np.round(1/rate,4)) + ':1',Format2)
    Worksheet.write(Final4.shape[0] + 2,FinalSum.shape[1] + Final4.shape[1] + 3,'Current EUR to USD:',Format2)

    Worksheet2.write(Mg1FinalSum.shape[0] + 3,Mg1Sum.shape[1] + Mg1FinalSum.shape[1] + 3,'1:' + str(rate) + ' or ' + str(np.round(1/rate,4)) + ':1',Format2)
    Worksheet2.write(Mg1FinalSum.shape[0] + 2,Mg1Sum.shape[1] + Mg1FinalSum.shape[1] + 3,'Current EUR to USD:',Format2)

    Worksheet3.write(0,0,'Exchange',Format3)
    Worksheet3.write(0,1,'Code',Format3)

    Worksheet6.write(0,0,'Exchange',Format3)
    Worksheet6.write(0,1,'Code',Format3)

    Worksheet8.write(0,Fd6Sum.shape[1] + Fd6.shape[1] + 6 ,'Cash U: {}'.format(str(CashU*100) + '%'),Format2)

    Worksheet9.write_comment(Fd7Sum.shape[0]+3,Fd7.shape[1] + 6, 'PLEASE FILL IN CASH AT BANK BELOW', {'x_scale': 1.2, 'y_scale': 0.8})
    Worksheet12.write(0,0,'NAV:',Format2)
    Worksheet12.write(0,1,NAV,Dollars)
    Worksheet12.write(0,2,'PLEASE FILL IN % COL',Dollars)

    for i in range(1,Fd10.shape[0]+1):
        Worksheet12.write(2+i,1,'=C{}/100*$B$1'.format(3+i))
    for i in range(1,Fd7Final.shape[0]+1):
        Worksheet9.write(Fd7Sum.shape[0]+3+i,Fd7.shape[1] + 8, '=H{}-I{}'.format(Fd7Sum.shape[0]+4+i,Fd7Sum.shape[0]+4+i)) 
        Worksheet9.write(Fd7Sum.shape[0]+3+i,Fd7.shape[1] + 7, '=H{}*100/I{} & "%"'.format(Fd7Sum.shape[0]+4+i,Fd7Sum.shape[0]+4+i)) 
    Writer.save()
    
if __name__ == '__main__':
    d = {}
    # Tkinter instance for front GUI
    front = Tk()
    # Size of GUI
    front.minsize(width=300, height=80)
    # Title of GUI
    front.title('FCB reporting')
    # Allow to not be resizable
    front.resizable(0,0)
    
    Labels = ["Dump File:","Sheet Name:","Keys File:","Date","NAV","CashU","OTC_Margin"]
    Label(front, text = Labels[0]).grid(row = 0, column = 0)
    Label(front, text = Labels[1]).grid(row = 0, column = 2)
    Label(front, text = Labels[2]).grid(row = 2, column = 0)
    Label(front, text = Labels[3]).grid(row = 2, column = 2)
    Label(front, text = Labels[4]).grid(row = 4, column = 2)
    Label(front, text = Labels[5]).grid(row = 6, column = 2)
    Label(front, text = Labels[6]).grid(row = 8, column = 2)
    for i in range (1,6):
        d["e{}".format(i)] = Entry(front,width=20)
    
    d["e1"].grid(row = 1, column = 2)
    d["e2"].grid(row = 3, column = 2)
    d["e3"].grid(row = 5, column = 2)
    d["e4"].grid(row = 7, column = 2)
    d["e5"].grid(row = 9, column = 2)
    Dump = StringVar(front)
    Keys = StringVar(front)
    
    var1 = IntVar()
    # Save function that writes to results.txt file the present inputs
    def save():
        f = open('entries.txt','w')
        for element in d:
            print(d[element].get())
            f.write(d[element].get() + '\n')
        print("SUCCESS: Results have been saved in results.txt file")
        f.close() 

    # Load function that deletes entries that you put in and inserts the ones from results.txt file
    def load(d,state):
        f = open('entries.txt','r')
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
    
    
    b1 = Button(front,text='Make',command=lambda:(initialvalues(),main()),bg='red').grid(row=10,column=2)
    b2 = Button(front,text='Load Previous',command=lambda:(load(d,'disabled')),bg='orange').grid(row=10,column=1)
    b3 = Button(front,text="Save inputs", command=lambda: save(),bg = 'DeepSkyBlue3').grid(row=11,column = 1)
    b4 = Checkbutton(front,text='Edit Values?',variable=var1,command=lambda: load(d,'disabled')).grid(row=12,column=1)
    b5 = Button(front, text='Quit', command=lambda: (front.destroy()),bg ='IndianRed4').grid(row=10, column=0, sticky=W, pady=4)
    initialvalues()
    mainloop()


# In[ ]:


df = pd.DataFrame([['Qwotc',1,2,3],['Qwotc2',4,5,6],['NTC',7,8,3]],columns = ['Ind','A','B','C'])

u = {'Qwotc':'RMB','Qwotc2':'RMB','NTC':'kay'}
df['Pop'] = df['Ind'].apply(lambda x: u[x])
display(df)


# In[ ]:


d = {'ok':'ata'}

ok = 'ok'
print(d[ok])

