import pandas as pd
import numpy as np
from pandas import DataFrame
import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SquareTransactionPath = os.path.join(BASE_DIR,'transaction.xlsx')

#This is the first Sales Receipt Number to import into Quickbooks. Go to Quickbooks and find the latest sale receipt number, increase that number by one
#to determine the START_SALESRECEIPT_NUMBER

START_SALESRECEIPT_NUMBER = int(input("Enter sales receipt number: "))


#load all rows in Excel to a raw dataframe
raw_df = pd.read_excel(SquareTransactionPath)

#drop out unneccessary columns
raw_df.drop(columns=["Time","Time Zone","Source","Transaction ID","Payment ID","Card Brand","PAN Suffix","Device Name","Staff Name","Staff ID","Details","Description","Event Type","Location","Dining Option","Customer ID","Customer Name","Customer Reference ID","Device Nickname","Deposit Details","Fee Percentage Rate","Fee Fixed Rate"], axis = 1, inplace = True)

#credit card dataframe is the row whose Deposit ID has value
creditcard_df = raw_df.dropna(subset =['Deposit ID'])


#cash dataframe is the row whose Deposit ID has no value
cash_df = raw_df[raw_df['Deposit ID'].isnull()]

#connect to Sqlite and create tables from those dataframes
con = sqlite3.connect("transaction.db")
raw_df.to_sql('raw',con,index = False, if_exists = 'replace')
creditcard_df.to_sql('creditcard',con,index = False, if_exists = 'replace')
cash_df.to_sql('cash',con,index = False, if_exists = 'replace')


cursor = con.cursor()

#Drop CreditCard Pivot Table if it exists to make sure the newly created table is clean
cursor.execute('''DROP TABLE IF EXISTS CreditCard_Pivot_Table''')


#create a new CreditCard_Pivot_Table from credit card table. This calculates sum of net sales, sales taxes, tips, fee, Taxable Sales Value, Non Taxable Sales value
#this groups by the Deposit ID as the deposit should match what Square deposits hit the bank account
cursor.execute(''' CREATE TABLE CreditCard_Pivot_Table AS SELECT Date, Sum("Net Sales") As "Total_Net_Sales", Sum("Tax") As "Total_Taxes", 
round(Sum("Tip"),2) As "Total_Tips",(-Sum("Fees")) As "Total_Fees", sum("Total Collected") As "Total_Collected",Sum("Net Total") AS "Net_Deposit", round((Sum(Tax)/0.0825),2) As Total_Taxable_Sales, 
round((Sum("Net Sales") - (Sum(Tax)/0.0825)),2) As Total_Non_Taxable_Sales, 
(round((Sum(Tax)/0.0825),2) + round((Sum("Net Sales") - (Sum(Tax)/0.0825)),2)=Sum("Net Sales")) As "Check_Sum"
FROM creditcard
GROUP BY "Deposit ID"
ORDER BY Date ''')
con.commit()

cursor.execute('''SELECT Date, Total_Taxable_Sales, Total_Non_Taxable_Sales, Total_Tips FROM CreditCard_Pivot_Table''')
rows = cursor.fetchall()


Sales_Date = []
Taxable_Item = []
Non_Taxable_Item = []
Taxable_Sales_Values = []
Non_Taxable_Sales_Values = []
Tip_Item = []
Tip_Values = []

SalesReceiptRefNumber = []



for row in rows:
    
    SalesReceiptRefNumber.append(START_SALESRECEIPT_NUMBER)
    Sales_Date.append(row[0])
    
    Taxable_Item.append("Food Sales")
    Taxable_Sales_Values.append(row[1])
    
    Non_Taxable_Item.append("Food Sales Non Tax")
    Non_Taxable_Sales_Values.append(row[2])
    
    Tip_Item.append("New Tip")
    Tip_Values.append(row[3])
    
    START_SALESRECEIPT_NUMBER = START_SALESRECEIPT_NUMBER + 1
    

# create 3 dataframes from 3 lists: taxable, non taxable, and tip
df1 = pd.DataFrame(list(zip(SalesReceiptRefNumber,Sales_Date,Taxable_Item,Taxable_Sales_Values)),columns = ['SalesReceiptRefNumber','Sales_Date','Item','Amount'])
df2 = pd.DataFrame(list(zip(SalesReceiptRefNumber,Sales_Date,Non_Taxable_Item,Non_Taxable_Sales_Values)),columns = ['SalesReceiptRefNumber','Sales_Date','Item','Amount'])
df3 = pd.DataFrame(list(zip(SalesReceiptRefNumber,Sales_Date,Tip_Item,Tip_Values)),columns = ['SalesReceiptRefNumber','Sales_Date','Item','Amount'])

#combine 3 dataframes in to one final dataframe
final_df = pd.concat([df1,df2,df3])

#add a new column called "Customer Name" and fill with "General Customer" value across the board
final_df['Customer_Name'] = "General Customer"

#add a new column called "Payment Method" and fill with "Visa" value across the board
final_df['Payment_Method'] = "Visa"

#sort the final dataframe by sales receipt number. This sorting is to keep all line items of the same sales receipt stay close together
final_df.sort_values(['SalesReceiptRefNumber'],ascending = [True],inplace = True)

#load the final dataframe to sqlite table, call it import table
final_df.to_sql('import',con,index = False, if_exists = 'replace')

#create an excel sheet from the final dataframe, ready to import to Quickbooks 
writer = pd.ExcelWriter("import.xlsx",engine='xlsxwriter')
final_df.to_excel(writer,sheet_name='Sheet1',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')
writer.save()

#close connection to the sqlite
con.close()