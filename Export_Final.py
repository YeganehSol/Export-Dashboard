#The last Script for export dashboard which gathers data from all sources (Jame'e mali,txt,excel and factsaleiframe) ,
#transfer anomlaie to DW 172.31.31.29 and finally it sends emails 




import os
import pandas as pd
import numpy as np
import pyodbc
from datetime import datetime, timedelta
import os
import json
from Functions import return_jalali_date, user_pyodbc_connection, pyodbc_connection
from Functions import Node1va2_Username, Node1va2_Password
import jdatetime
from datetime import timedelta
import jdatetime
from Export_Functions import *
import requests




print(Node1va2_Username)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
os.chdir(r'D:\Projects\Export Dashboard')




#Reading txt / xlsx Data
#NewExport = pd.read_excel(r'''\\172.31.50.113\Sale & Distribution\NewExport.xlsx''' , sheet_name=  'CurrencyParityRate')
nonsys = pd.read_csv(r'''\\172.31.50.113\Sale & Distribution\NonSysyemiForPowerBI.txt''', sep=',')
FactExcelAggr = pd.read_excel(r'\\172.31.50.113\Sale & Distribution\FactExcelAggr.xlsx') #FactExcelAggr + Khazane = nonsys



nonsys = nonsys[nonsys['TableInfo'] == 'JameMali'] #Beacuse of difference between nonsys and xlsx file
nonsys['NetAmount'] = nonsys['NetAmount']/1000000 #Turn into Toman
FactExcelAggr['Amount_Net'] = FactExcelAggr['Amount_Net']/1000000 #Turn into Toman


##############################################
#Preprocess and maindates

nonsys['MainDate'].astype('str')
nonsys['CurrencyCode'] = nonsys['CurrencyDesc'].map(CurrencyToCode)
nonsys['MainDate'] = nonsys['MainDate'].astype('str')
nonsys['YearMonth'] = nonsys['MainDate'].str[:6]
nonsys['YearMonth']=nonsys['YearMonth'].astype('str')



FactExcelAggr['CurrencyCode'] = FactExcelAggr['CurrencyDesc'].map(CurrencyToCode)
FactExcelAggr['Main_Date'] = FactExcelAggr['Main_Date'].astype('str')
FactExcelAggr['YearMonth'] = FactExcelAggr['Main_Date'].str[:6]
FactExcelAggr = FactExcelAggr[FactExcelAggr['Flag_InOut'] == 2]


FactExcelAggr = FactExcelAggr[FactExcelAggr['YearMonth'] >= '140300']
nonsys = nonsys[nonsys['YearMonth'] >= '140300']



# DataQuality Connection
#connection, cursor = user_pyodbc_connection("172.31.31.29", "DataQuality")

# Connection string to DataWarehouse DataBase / Node2
execute_connection, execute_cursor = user_pyodbc_connection(server="172.31.50.98",
                                                            database_name="SaleIntegratedModel",
                                                            username=Node1va2_Username,
                                                            password=Node1va2_Password)


today_date = datetime.now()
current_time = str(today_date.time())
persian_today_date = return_jalali_date(today_date)[0]
year = 1403



#Gathering Data from SaleIntegratedModel..factsalealk


ALK_rows = []

Alk_query = """
select CompanyCode ,  round((SUM(case when InvoceType = 2 then NetAmount else 0 end) - SUM(case when InvoceType = 3 then NetAmount else 0 end))/ 1000000.0, 0) as SumQatE
from SaleIntegratedModel..factsalealk
where SaleType = 2 and left(maindate , 6) >= 140300
group by CompanyCode"""



try:
    execute_cursor.execute(Alk_query)
    ALK_total_rows = execute_cursor.fetchall()
    
    for row in ALK_total_rows:
        if not row[0]:
            row[0] = "Null"
        if not row[1]:
            row[1] = 0
            
        CompanyCode, ALK_total = row
        ALK_rows.append(
            {
                "CompanyCode": CompanyCode,
                "ALK_total": ALK_total
            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)

ALK_df = pd.DataFrame(ALK_rows)

ALK_df = ALK_df[[ "CompanyCode", 'ALK_total']]


#Gathering Data from SaleIntegratedModel..FactSaleNonDistribut



Nondistribute_rows = []

NondistributeCompanies_query = """
select CompanyCode ,  round((SUM(case when InvoceType = 2 then NetAmount else 0 end) - SUM(case when InvoceType = 3 then NetAmount else 0 end))/ 1000000.0, 0) as SumQatE
from SaleIntegratedModel..FactSaleNonDistribut  where SaleType = 2 and left(maindate , 6) >= 140300
group by CompanyCode """



try:
    execute_cursor.execute(NondistributeCompanies_query)
    Nondistribute_total_rows = execute_cursor.fetchall()
    
    for row in Nondistribute_total_rows:
        if not row[0]:
            row[0] = "Null"
        if not row[1]:
            row[1] = 0
            
        CompanyCode, Nondistribute_Total = row
        Nondistribute_rows.append(
            {
                "CompanyCode": CompanyCode,
                "Nondistribute_Total": Nondistribute_Total
            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)

Nondistribute_df = pd.DataFrame(Nondistribute_rows)

Nondistribute_df = Nondistribute_df[[ "CompanyCode", 'Nondistribute_Total']]



#Merge All sources Data

#preparing nonsys data
 
TotalSale_Mali_and_excel_df = nonsys[['CompanyCode','NetAmount']].groupby(by = 'CompanyCode').sum().reset_index()
FactExcelAggr_df = FactExcelAggr[['CompanyCode' , 'Amount_Net']].groupby(by = 'CompanyCode').sum().reset_index()




Nondistribute_df.columns = ['CompanyCode' , 'N_NetAmount']
TotalSale_Mali_and_excel_df.columns = ['CompanyCode' , 'O_NetAmount']
FactExcelAggr_df.columns = ['CompanyCode' , 'X_NetAmount']
ALK_df.columns = ['CompanyCode' , 'ALK_NetAmount']


#Merge Nondistribute / Excel and khazane

Total_rows = pd.concat([Nondistribute_df , TotalSale_Mali_and_excel_df , FactExcelAggr_df] , axis = 0) 
Total_rows.TotalSale = Total_rows.N_NetAmount.astype(float)
Total_rows = Total_rows.groupby(by = 'CompanyCode').sum().reset_index()

print(f'''the total_rows is {len(Total_rows)} and the number of Companies is {Total_rows.CompanyCode.nunique()}''')

nonsys_df_Net = TotalSale_Mali_and_excel_df[['CompanyCode' , 'O_NetAmount']].groupby(by = ['CompanyCode']).sum().reset_index()


Total_rows_Net_tot = pd.merge(Nondistribute_df , nonsys_df_Net ,how = 'outer' , right_on= ['CompanyCode'] , left_on =  ['CompanyCode']) #merge nonsys and nondistribute
Total_rows_Net_tot_final = pd.merge(Total_rows_Net_tot , FactExcelAggr_df , how = 'outer' ,right_on= ['CompanyCode'] , left_on =  ['CompanyCode'])
Total_rows_Net_tot_final = Total_rows_Net_tot_final[['N_NetAmount', 'O_NetAmount', 'X_NetAmount','CompanyCode']].groupby(by = 'CompanyCode').sum().reset_index()




#FactIframe


# -------------------------------------------------- Iframe Amount per Companay

iframe_rows = []

iframe_sql = """
SELECT CompanyCode, dim.Company_PName,
            ROUND((SUM(CASE WHEN Sale_State = 2 THEN ROUND(Amount_Net, 0) ELSE 0 END) -
                   SUM(CASE WHEN Sale_State = 3 THEN ROUND(Amount_Net, 0) ELSE 0 END)) / 1000000.0, 0) AS Iframe_Total
FROM Qlikview.SSAS_Aggr.FactSaleIFrame as fact WITH (NOLOCK)
LEFT JOIN QlikView.SSAS_Aggr.Company as dim WITH (NOLOCK) ON fact.CompanyCode = dim.[Master Code]
WHERE Flag_InOut = 2 and left(Main_Date , 6) >=140300
GROUP BY CompanyCode, dim.Company_PName;
"""

try:
    execute_cursor.execute(iframe_sql)
    iframe_total_rows = execute_cursor.fetchall()
    
    for row in iframe_total_rows:
        if not row[1]:
            row[1] = "Null"
        if not row[0]:
            row[0] = "Null"
        if not row[2]:
            row[2] = 0
            
        CompanyCode, CompanyName, Iframe_Total = row
        iframe_rows.append(
            {
                "CompanyName": CompanyName,
                "CompanyCode": CompanyCode,
                #"Main_Date" : Main_Date,
                "Iframe_Total": Iframe_Total

            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)

iframe_df = pd.DataFrame(iframe_rows)



iframe_df = iframe_df[["CompanyName", "CompanyCode",  "Iframe_Total"]]


iframe_df_Net = iframe_df[['CompanyName' , 'CompanyCode' , 'Iframe_Total']]
iframe_df_Net = iframe_df_Net.groupby(by = ['CompanyName'	,'CompanyCode']).sum().reset_index()

###########################################################################


CurrencyRate_rows = []

CurrencyRatesql = '''SELECT YearMonth
      ,CurrencyParityRateDolar
	  ,b.CurrencyID
	  ,b.CurrencyName
  FROM [Qlikview].[dbo].[FactCurrencyRate] a
  inner join [Qlikview].[dbo].ExportCurrency b
  on a.CurrencyID = b.CurrencyID'''



try:
    execute_cursor.execute(CurrencyRatesql)
    CurrencyRate_total_rows = execute_cursor.fetchall()
    
    for row in CurrencyRate_total_rows:
            
        YearMonth, CurrencyParityRateDolar, CurrencyID,CurrencyName = row
        CurrencyRate_rows.append(
            {
                "YearMonth": YearMonth,
                "CurrencyParityRateDolar": CurrencyParityRateDolar,
                "CurrencyID" : CurrencyID,
                "CurrencyName": CurrencyName

            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)

CurrencyRate_df = pd.DataFrame(CurrencyRate_rows)



########PreProccess##########

CurrencyRate_df['YearMonth'] = CurrencyRate_df['YearMonth'].astype(str)
CurrencyRate_df['YearMonth'] = CurrencyRate_df['YearMonth'].astype("string")
CurrencyRate_df['CurrencyID'] = CurrencyRate_df['CurrencyID'].astype("string")


################################################

whole_final = pd.merge(iframe_df_Net , Total_rows_Net_tot_final ,how = 'outer' , right_on= ['CompanyCode'] , left_on =  ['CompanyCode'])
whole_final = pd.merge(whole_final , ALK_df ,how = 'outer' , right_on= ['CompanyCode'] , left_on =  ['CompanyCode'])



whole_final.N_NetAmount = whole_final.N_NetAmount.fillna(0)
whole_final.O_NetAmount = whole_final.O_NetAmount.fillna(0)
whole_final.X_NetAmount = whole_final.X_NetAmount.fillna(0)
whole_final.ALK_NetAmount = whole_final.ALK_NetAmount.fillna(0)













################################################################################################

for col in ['Iframe_Total','N_NetAmount','X_NetAmount','O_NetAmount','ALK_NetAmount']:
    whole_final[col] = (
        whole_final[col]
        .replace({',': ''}, regex=True)
        .replace(r'[^\d\.\-]+', '', regex=True)
        .apply(pd.to_numeric, errors='coerce')
        .fillna(0)
        .round(0)          
        .astype('Int64')   
    )


print(whole_final.info())


whole_final['Iframe_Total'] = pd.to_numeric(whole_final['Iframe_Total'], errors='coerce').fillna(0).astype('Int64')
whole_final['N_NetAmount'] = pd.to_numeric(whole_final['N_NetAmount'], errors='coerce').fillna(0).astype('Int64')
whole_final['X_NetAmount'] = pd.to_numeric(whole_final['X_NetAmount'], errors='coerce').fillna(0).astype('Int64')
whole_final['O_NetAmount'] = pd.to_numeric(whole_final['O_NetAmount'], errors='coerce').fillna(0).astype('Int64')
whole_final['ALK_NetAmount'] = pd.to_numeric(whole_final['ALK_NetAmount'], errors='coerce').fillna(0).astype('Int64')



################################################################################################

whole_final['TotalSale_Whole'] = whole_final['N_NetAmount'] + whole_final['O_NetAmount'] + whole_final['X_NetAmount'] + whole_final['ALK_NetAmount']
whole_final.drop(['N_NetAmount' , 'O_NetAmount' ,'X_NetAmount' ] , axis = 1 , inplace = True)



whole_final['Differenece'] = whole_final['Iframe_Total'] - whole_final['TotalSale_Whole']
df_anomalies = whole_final[whole_final['Differenece'] > 10]
df_anomalies.drop(['TotalSale_Whole' , 'ALK_NetAmount' , 'Iframe_Total'] ,axis = 1 ,  inplace = True)
df_anomalies.columns = ['نام شرکت' , 'کد مستر شرکت' , 'اختلاف'] #Creating df of anomalies






################################## Roham's Table ####################################


FactExport_rows = []

FactExportCompanies_query = """
SELECT 
    YearMonth,
    CompanyCode,
    SUM(RealAmount * DollarRate) AS SumInUSD
FROM [QlikView].[dbo].[FactExport]
where YearMonth > 140300
GROUP BY 
    YearMonth, 
    CompanyCode;"""



try:
    execute_cursor.execute(FactExportCompanies_query)
    FactExport_total_rows = execute_cursor.fetchall()
    
    for row in FactExport_total_rows:
        if not row[0]:
            row[0] = "Null"
        if not row[1]:
            row[1] = 0
            
        YearMonth , CompanyCode, SumInUSD = row
        FactExport_rows.append(
            {
                "YearMonth":YearMonth,
                "CompanyCode": CompanyCode,
                "SumInUSD": SumInUSD
            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)

FactExport_df = pd.DataFrame(FactExport_rows)

FactExport_df = FactExport_df[["YearMonth", "CompanyCode", 'SumInUSD']]



#################################################################################################################
########################### Gathering RealAmount Data from NonDistribute ########################################

Nondistribute_rows = []

NondistributeCompanies_query = """
select   sum(RealAmount) SumOfRealAmount , CompanyCode , left(MainDate , 6) YearMonth , CurrencyDesc , InvoceType
 from SaleIntegratedModel..FactSaleNonDistribut
where RealAmount is not null and MainDate >14030000 
group by CompanyCode ,left(MainDate , 6) , CurrencyDesc ,InvoceType
"""



try:
    execute_cursor.execute(NondistributeCompanies_query)
    Nondistribute_total_rows = execute_cursor.fetchall()

    for row in Nondistribute_total_rows:
            
        SumOfRealAmount , CompanyCode ,  YearMonth , CurrencyDesc , InvoceType  = row
        Nondistribute_rows.append(
            {
                "CompanyCode": CompanyCode,
                "SumOfRealAmount": SumOfRealAmount,
                "YearMonth" :YearMonth,
                'CurrencyDesc' : CurrencyDesc,
                "InvoceType" : InvoceType

            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)


Nondistribute_df = pd.DataFrame(Nondistribute_rows)
Nondistribute_df = Nondistribute_df[['CompanyCode','SumOfRealAmount', 'YearMonth','CurrencyDesc','InvoceType']]


########################### Data Cleansing ##################################

Nondistribute_df.CurrencyDesc= Nondistribute_df.CurrencyDesc.str.strip()
Nondistribute_df['CurrencyCode'] = Nondistribute_df.CurrencyDesc.map(CurrencyToCode)




########################### Convert Currency to dollar #########################



mergedFactExport=pd.merge(Nondistribute_df, CurrencyRate_df[['CurrencyParityRateDolar' , 'YearMonth' , 'CurrencyID']] , right_on= ['YearMonth' , 'CurrencyID'] , left_on=['YearMonth' , 'CurrencyCode'] , how = 'left')
mergedFactExport.drop(['CurrencyID'] , axis = 1 , inplace = True)


mergedFactExport['RealAmountDollar'] = mergedFactExport.SumOfRealAmount.astype('float') * mergedFactExport.CurrencyParityRateDolar.astype('float')

mergedFactExportInv = mergedFactExport.groupby(['YearMonth' ,'CompanyCode' , 'InvoceType']).sum('RealAmountDollar').reset_index()


grouped = (
    mergedFactExportInv.groupby(['YearMonth', 'CompanyCode', 'InvoceType'])['RealAmountDollar']
      .sum()
      .unstack(fill_value=0)   # ستون جدا برای 2 و 3
      .reset_index())

grouped["RealAmount_nonDis"] = grouped.get(2, 0) - grouped.get(3, 0)




############################### Add Alkowsar Data ############################

#ALK_RealAMount


ALK_rows_R = []

Alk_query_R = """
select left(maindate , 6) MainDate , CompanyCode , CurrencyDesc ,  round(SUM(case when InvoceType = 2 then RealAmount else 0 end) - SUM(case when InvoceType = 3 then RealAmount else 0 end) , 0) as SumQatE
from SaleIntegratedModel..factsalealk
where  left(maindate , 6) >= 140300
group by MainDate , CompanyCode , CurrencyDesc
"""



try:
    execute_cursor.execute(Alk_query_R)
    ALK_total_rows_R = execute_cursor.fetchall()
    
    for row in ALK_total_rows_R:
        if not row[0]:
            row[0] = "Null"
        if not row[1]:
            row[1] = 0
            
        MainDate , CompanyCode, CurrencyDesc, ALK_total_R = row
        ALK_rows_R.append(
            {
                "YearMonth": MainDate,
                "CompanyCode": CompanyCode,
                'CurrencyDesc' : CurrencyDesc,
                "ALK_total": ALK_total_R
            }
        )
except Exception as e:
     print("Executing iframe_sql query failed", e)

ALK_df_R = pd.DataFrame(ALK_rows_R)

ALK_df_RealAmount = ALK_df_R[[ 'YearMonth', "CompanyCode", "CurrencyDesc" , 'ALK_total']]




ALK_df_R['CurrencyCode'] = ALK_df_R['CurrencyDesc'].map(CurrencyToCode)



############################ Preproccessing ################################

merged_FactExcel = pd.merge(FactExcelAggr , CurrencyRate_df , left_on = ['YearMonth' , 'CurrencyCode'] , right_on = ['YearMonth' , 'CurrencyID'] , how = 'left')
merged_nonsys = pd.merge(nonsys , CurrencyRate_df , left_on = ['YearMonth' , 'CurrencyCode'] , right_on = ['YearMonth' , 'CurrencyID'] , how = 'left')
merged_ALK = pd.merge(ALK_df_R , CurrencyRate_df[['CurrencyParityRateDolar' ,'YearMonth', 'CurrencyID']] , left_on = ['YearMonth' , 'CurrencyCode'] , right_on = ['YearMonth' , 'CurrencyID'] , how = 'left' )





##Check Marjuei

merged_FactExcel['RealAmount_Dollar'] = merged_FactExcel['CurrencyParityRateDolar'].astype('float') * merged_FactExcel['RealAmount']
merged_nonsys['RealAmount_Dollar_sys'] =  merged_nonsys['CurrencyParityRateDolar'].astype('float') * merged_nonsys['RealAmount']
merged_ALK['RealAmount_Dollar_ALK'] = merged_ALK['CurrencyParityRateDolar'].astype('float') * merged_nonsys['RealAmount']



merged_All_Temp = pd.merge(merged_FactExcel[['CompanyCode' ,'RealAmount_Dollar' ,  'YearMonth']] , merged_nonsys[['CompanyCode' ,'RealAmount_Dollar_sys' ,  'YearMonth']] , how = 'outer' , left_on = 'CompanyCode' , right_on = 'CompanyCode' )
merged_All_Tempp = pd.merge(merged_All_Temp , grouped , right_on = 'CompanyCode' , left_on = 'CompanyCode' , how = 'outer')


#YearMonth handling 

merged_All_Tempp.YearMonth = merged_All_Tempp.YearMonth.fillna(merged_All_Tempp.YearMonth_y)
merged_All_Tempp.YearMonth = merged_All_Tempp.YearMonth.fillna(merged_All_Tempp.YearMonth_x)

#Drop Extera Columns
merged_All_2 = merged_All_Tempp.drop(['YearMonth_x' ,'YearMonth_y' , 2 , 3 ] , axis = 1 )


#Merge ALkowsar with main data
merged_All = pd.merge(merged_All_2 , merged_ALK , how = 'outer' , left_on = 'CompanyCode' , right_on = 'CompanyCode' )

#YearMonth Handling

merged_All['YearMonth_x'] = merged_All.YearMonth_x.fillna(merged_All.YearMonth_y)

merged_All.drop('YearMonth_y' , axis = 1 , inplace = True)

merged_All.rename(columns = {'YearMonth_x' : 'YearMonth'} , inplace = True)





###################### Fill Na ##########################


#Merge jameMali and 

merged_All['RealAmount_Dollar'] = merged_All.RealAmount_Dollar.fillna(0)
merged_All['RealAmount_Dollar_sys'] = merged_All.RealAmount_Dollar_sys.fillna(0)
merged_All['RealAmount_nonDis'] = merged_All.RealAmount_nonDis.fillna(0)


merged_All = merged_All.groupby(by = ['CompanyCode' , 'YearMonth']).sum().reset_index()

merged_All['RealAmount_Final'] = merged_All['RealAmount_Dollar'] + merged_All['RealAmount_Dollar_sys'] + merged_All['RealAmount_nonDis'] + merged_All['RealAmount_Dollar_ALK']


merged_All_Final = pd.merge(merged_All , FactExport_df , left_on = ['CompanyCode' , 'YearMonth']  , right_on =['CompanyCode' , 'YearMonth'] )
 
merged_All_Final['Difference'] = merged_All_Final['SumInUSD'].astype('float') - merged_All_Final['RealAmount_Final'].astype('float')



########################### Determine The threshold #########################

merged_All_Final_Anomalies = merged_All_Final[merged_All_Final['Difference'] > 100] # determine threshold for anomaly detection


######################### Rename Columns in order to insert to table #######################

merged_All_Final_Anomalies.drop(['CurrencyDesc' , 'CurrencyParityRateDolar' , 'ALK_total' , 'CurrencyID' , 'CurrencyCode'] , axis = 1 , inplace = True)


######################### Add Jalali Date #########################


#Add Date
from persiantools.jdatetime import JalaliDate

merged_All_Final_Anomalies['TodayDate'] = [JalaliDate.today().strftime('%Y%m%d')] * len(merged_All_Final_Anomalies)
merged_All_Final_Anomalies.columns = ['CompanyCode' , 'Date' , 'Excel' , 'non_sys' , 'non_Distribute', 'ALK' , 'Final_Result' , 'FactExport' , 'Difference' , 'TodayDate']


#Save Merged all final as a csv
merged_All_Final_Anomalies.to_csv(f'''D:\Projects\Export Dashboard\Anomalies\Export_Anomalies_{JalaliDate.today().strftime('%Y%m%d')}.csv''', encoding = 'utf-8-sig')



######################### Insert Data to Table ############################
# Find the lastest Date in CSV which saved on folder and Check duplicate dates

import re

folder_path = r"D:\Projects\Export Dashboard\Anomalies"  

csv_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) 
             if file.startswith("Export_Anomalies_") and file.endswith(".csv")]

if not csv_files:
    print("No matching files found.")
else:
    last_modified_file = max(csv_files)
    last_modified_files = csv_files[-2:]

    dfToday = pd.read_csv(last_modified_file)
    dfYesterday = pd.read_csv(min(last_modified_files))

    print(f"dfToday: {last_modified_file} ,\n, dfYesterday: {last_modified_files[0]}")

pathToday = last_modified_file
pathYesterday = min(last_modified_files)

matchToday = re.search(r'_(\d{8})\.csv$', pathToday)
matchYesterday = re.search(r'_(\d{8})\.csv$', pathYesterday)


if matchToday:
    extracted_dateToday = matchToday.group(1)  
    extracted_dateYesterday = matchYesterday.group(1)
    print(f"Extracted date Today: {extracted_dateToday}")
    print(f"Extracted date Yesterday: {extracted_dateYesterday}")
else:
    print("No date found in the string.")



def insert_to_sql(data, servername, Database):
    ServerName = servername
    Database = Database
    conn = pyodbc.connect(""" 
    Driver={{SQL Server}};
    server={0};
    Database={1};
    Trusted_Connection=yes;
""".format(ServerName, Database))
    cursor = conn.cursor()
# Check if today's date already exists in the database
    query = "SELECT MAX(TodayDate) AS MaxDate FROM [DataQuality].[Ex].[FactExport]"
    DataBaseDate = pd.read_sql_query(query, conn)
    max_date_db = str(DataBaseDate['MaxDate'].values[0])
    if extracted_dateToday == max_date_db:
        print("Date already exists in database. Skipping insert.")
        status = 0
    else:
        for index, row in merged_All_Final_Anomalies.iterrows():
            cursor.execute("""
                INSERT INTO [DataQuality].[Ex].[FactExport] (
                           CompanyCode , Date , Excel , non_sys , non_Distribute, ALK , Final_Result , FactExport , Difference , TodayDate)
                           Values (?,?,?,?,?,?,?,?,?,?)""" , row.CompanyCode ,  row.Date , row.Excel , row.non_sys , row.non_Distribute, row.ALK , row.Final_Result , row.FactExport , row.Difference , row.TodayDate)                        


        conn.commit()
        print("Data inserted.")
        status = 1

    cursor.close()
    return status
    

# Call and print result
status = insert_to_sql(merged_All_Final_Anomalies, '172.31.31.29', 'DataQuality')
print("Insert status:", status)


########################## Calculate Number of days #################################

###### Connection to 172.31.31.29 ########
execute_connection, execute_cursor = user_pyodbc_connection(
    server="172.31.31.29",
    database_name="DataQuality"
)

Anomalies_rows = []

AnomaliesRows = """
SELECT [Date]
      ,[CompanyCode]
      ,[Excel]
      ,[non_sys]
      ,[ALK]
      ,[non_Distribute]
      ,[Final_Result]
      ,[FactExport]
      ,[Difference]
      ,[TodayDate]
  FROM [DataQuality].[Ex].[FactExport]
  where date > 140400
"""

try:
    execute_cursor.execute(AnomaliesRows)
    Anomalies_total_rows = execute_cursor.fetchall()
    
    for  Date ,CompanyCode, Excel,non_sys,ALK,non_Distribute,Final_Result ,FactExport, Difference ,TodayDate in Anomalies_total_rows:
        Anomalies_rows.append(
            {
                "CompanyCode": CompanyCode,
                "TodayDate": TodayDate,
                'Date': Date,
                "Excel": Excel,
                "Non_System":non_sys,
                "Alkowsar":ALK,
                'non_Distribute':non_Distribute,
                "Final_Result":Final_Result,
                "FactExport":FactExport,
                "Difference":Difference
            }
        )

except Exception as e:
    print("Executing query failed:", e)


Anomalies_df = pd.DataFrame(Anomalies_rows, columns=["CompanyCode", "TodayDate", "Date","Excel","Non_System" , "Alkowsar" ,"non_Distribute" , "Final_Result" , "FactExport" , "Difference" ])


execute_cursor.close()
execute_connection.close()



def str_to_jalali(date_val):
    date_str = str(date_val) 
    year = date_str[:4]      
    month = date_str[4:6]     
    return f"{year}{month}"   


Anomalies_df['TodayJalaliDate'] = Anomalies_df['TodayDate'].apply(str_to_jalali)



Anomalies_df['TodayJalaliDate'] = Anomalies_df['TodayDate'].apply(str_to_jalali)
max_date = Anomalies_df['TodayJalaliDate'].max()
Anomalies_df['DateJalali'] = Anomalies_df['Date'].apply(str_to_jalali)
max_date_Anomali = Anomalies_df['TodayDate'].max()
today = jdatetime.date.today().strftime('%Y%m%d')



if today == max_date_Anomali:
    count = Anomalies_df[Anomalies_df['TodayDate'] == max_date_Anomali]
    
count.drop(['TodayJalaliDate' , 'DateJalali'] , axis = 1 , inplace = True)    





if not count.empty:
    payload = json.dumps({
        "token": "P880P2HLA6MO71PWTXTWR",
        "providerId": 6,
        "sendType": 0,
        "email": 'soleimani.yeganeh@golrang.com',
        "ccRecipients": ['Aghdasifam.Masoud@Golrang.com'],
        "subject": "Alert: KPI Values Exceeded In Export!",
        "body": f" \n\n {count.to_html(index=False, border=1, justify='center')}",
        "fileAttachmentAddress": ""
    })
    headers = {'Content-Type': 'application/json'}
    response = requests.post("https://Esp-api.gig.services/email/saveEmail", headers=headers, data=payload)
    print(response.status_code, response.text)
else:
    payload = json.dumps({
        "token": "P880P2HLA6MO71PWTXTWR",
        "providerId": 6,
        "sendType": 0,
        "email":'soleimani.yeganeh@golrang.com',
        "ccRecipients": ['Aghdasifam.Masoud@Golrang.com'], #It should be in []
        "subject": "Confirmation of Descending program execution in Export!",
        "body": "انجام شد",
        "fileAttachmentAddress": ""
    })
    headers = {'Content-Type': 'application/json'}
    response = requests.post("https://Esp-api.gig.services/email/saveEmail", headers=headers, data=payload)
    print(response.status_code, response.text)
