import os
import pandas as pd
import numpy as np
import pyodbc
from datetime import datetime, timedelta
import os
import json
from Functions import return_jalali_date, user_pyodbc_connection, pyodbc_connection
from Functions import Node1va2_Username, Node1va2_Password

print(Node1va2_Username)

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
os.chdir(r'D:\Projects\Export Dashboard')




#Reading txt / xlsx Data
NewExport = pd.read_excel(r'''\\172.31.50.113\Sale & Distribution\NewExport.xlsx''' , sheet_name=  'CurrencyParityRate')
nonsys = pd.read_csv(r'''\\172.31.50.113\Sale & Distribution\NonSysyemiForPowerBI.txt''', sep=',')
FactExcelAggr = pd.read_excel(r'\\172.31.50.113\Sale & Distribution\FactExcelAggr.xlsx') #FactExcelAggr + Khazane = nonsys


#Preprocess

CurrencyToCode = {'دینار' : '107',
        'افغاني' : '103',
         'دلار امريكا' : '101',
          'ريال' : '108',
           'يورو' : '102',
            'روبل' : '106',
             'dolar' : '101' ,
             'ريال' : '108',
             'دلار':'101',
             'روبل':'106',
        'Afghani' : '103',
        'IraqiDinar':'107',
         'USD' : '101',
          'Rial' : '108',
           'Euro' : '102',
            'RUB' : '106',
             'Dirham': '109',
              'TL': '104',
               'CAD':'105'}

CodeToCurrency = { '107': 'IraqDinar',
         '103':'Afghani',
         '101': 'USD',
         '108': 'Rial',
         '102':'Euro',
         '106': 'RUB'
}

#dict2 = {'ريال' : '108',
#        'Afghani' : '103',
#        'IraqiDinar':'107',
#         'USD' : '101',
#          'Rial' : '108',
#           'Euro' : '102',
#            'RUB' : '106',
#             'Dirham': '109',
#              'TL': '104',
#               'CAD':'105' }






nonsys = nonsys[nonsys['TableInfo'] == 'JameMali'] #Beacuse of difference between nonsys and xlsx file
nonsys['NetAmount'] = nonsys['NetAmount']/1000000 #Turn into Toman
FactExcelAggr['Amount_Net'] = FactExcelAggr['Amount_Net']/1000000 #Turn into Toman



#Preprocess and maindates

nonsys['MainDate'].astype('str')
nonsys['CurrencyCode'] = nonsys['CurrencyDesc'].map(CurrencyToCode)
nonsys['MainDate'] = nonsys['MainDate'].astype('str')
nonsys['YearMonth'] = nonsys['MainDate'].str[:6]
nonsys['YearMonth']=nonsys['YearMonth'].astype('str')

NewExport['YearMonth'] = NewExport['YearMonth'].astype(str)
NewExport['YearMonth'] = NewExport['YearMonth'].astype("string")
NewExport['CurrencyID'] = NewExport['CurrencyID'].astype("string")


FactExcelAggr['CurrencyCode'] = FactExcelAggr['CurrencyDesc'].map(CurrencyToCode)
FactExcelAggr['Main_Date'] = FactExcelAggr['Main_Date'].astype('str')
FactExcelAggr['YearMonth'] = FactExcelAggr['Main_Date'].str[:6]
FactExcelAggr = FactExcelAggr[FactExcelAggr['Flag_InOut'] == 2]


FactExcelAggr = FactExcelAggr[FactExcelAggr['YearMonth'] >= '140300']
nonsys = nonsys[nonsys['YearMonth'] >= '140300']





# DataQuality Connection



# DataQuality Connection
connection, cursor = user_pyodbc_connection("172.31.31.29", "DataQuality")

# Connection string to DataWarehouse DataBase / Node2
execute_connection, execute_cursor = user_pyodbc_connection(server="172.31.50.98",
                                                            database_name="SaleIntegratedModel",
                                                            username=Node1va2_Username,
                                                            password=Node1va2_Password)


today_date = datetime.now()
current_time = str(today_date.time())
persian_today_date = return_jalali_date(today_date)[0]
year = 1403


companies_App_dict = {3: {'CompanyName': 'گلرنگ سیستم', 'ApplicationCode': 1},
                      4: {'CompanyName': 'حمل و نقل گلرنگ ترابر', 'ApplicationCode': 1}, 
                      8: {'CompanyName': 'صنایع الکترونیک گلرنگ', 'ApplicationCode': 1}, 
                     13: {'CompanyName': 'آریان تجارت شرق', 'ApplicationCode': 1}, 
                     15: {'CompanyName': 'گل پخش اول', 'ApplicationCode': 1}, 
                     16: {'CompanyName': 'گروه صنعتی پاکشو', 'ApplicationCode': [1,4]}, 
                     17: {'CompanyName': 'صنایع سلولزی مارینا سان', 'ApplicationCode': 3}, 
                     18: {'CompanyName': 'پاکان پلاستکار', 'ApplicationCode': 1}, 
                     19: {'CompanyName': 'ماستر فوده', 'ApplicationCode': 1}, 
                     20: {'CompanyName': 'ایراندار', 'ApplicationCode': 1}, 
                     21: {'CompanyName': 'مارینا پخش هستی', 'ApplicationCode': 1}, 
                     24: {'CompanyName': 'افق توسعه معادن خاورمیانه', 'ApplicationCode': 1}, 
                     27: {'CompanyName': 'ویرا سیستم پویا', 'ApplicationCode': 1}, 
                     33: {'CompanyName': 'توسعه صنعتی سرافراز پارس (توسعه صنعتی سولفور کوب پارس (اسپیدکو))', 'ApplicationCode': 1}, 
                     34: {'CompanyName': 'توسعه فرآورده های نفتی افق خاورمیانه (موپیکو)', 'ApplicationCode': 1}, 
                     37: {'CompanyName': 'پاکان بسپار آرین', 'ApplicationCode': 3}, 
                     39: {'CompanyName': 'آریان کیمیا تک', 'ApplicationCode': [1,3]}, 
                     41: {'CompanyName': 'آزما نانو سیستم', 'ApplicationCode': 3}, 
                     42: {'CompanyName': 'فروشگاه های زنجیره ای افق کوروش', 'ApplicationCode': 5}, 
                     43: {'CompanyName': 'گلرنگ پخش', 'ApplicationCode': 1}, 
                     45: {'CompanyName': 'صنعت غذایی کورش', 'ApplicationCode': [5,1]}, 
                     47: {'CompanyName': 'فروشگاه های زنجیره ای فامیلی مدرن', 'ApplicationCode': [13,5]}, 
                     48: {'CompanyName': 'سامان پویش تامین (اسپات)', 'ApplicationCode': [1,11]}, 
                     51: {'CompanyName': 'کشت و صنعت برنج کوروش', 'ApplicationCode': 1}, 
                     54: {'CompanyName': 'صنعت خشکبار و حبوبات کوروش', 'ApplicationCode': 1}, 
                     55: {'CompanyName': 'هستی آرین تامین', 'ApplicationCode': [1,2]}, 
                     56: {'CompanyName': 'طلای ناب کوروش', 'ApplicationCode': 1}, 
                     59: {'CompanyName': 'پخش پدیده پایدار', 'ApplicationCode': 1}, 
                     60: {'CompanyName': 'سلامت پخش هستی', 'ApplicationCode': 9}, 
                     61: {'CompanyName': 'پدیده شیمی قرن', 'ApplicationCode': 1}, 
                     62: {'CompanyName': 'تولیدی فاران شیمی تویسرکان', 'ApplicationCode': 1}, 
                     63: {'CompanyName': 'پدیده شیمی غرب', 'ApplicationCode': [1,2,10]}, 
                     64: {'CompanyName': 'آرین سلامت سینا', 'ApplicationCode': [1,2,10]}, 
                     65: {'CompanyName': 'سپهر پلاستیک پدیده', 'ApplicationCode': 1}, 
                     66: {'CompanyName': 'پدیده شیمی جم', 'ApplicationCode': 1}, 
                     70: {'CompanyName': 'هستی آریا شیمی', 'ApplicationCode': 1}, 
                     71: {'CompanyName': 'آروین پلاست هکمتانه', 'ApplicationCode': 1}, 
                     73: {'CompanyName': 'مایا زیست فرآیند', 'ApplicationCode': 1}, 
                     74: {'CompanyName': 'ابیان دارو (هستی آرین دارو)', 'ApplicationCode': 1}, 
                     82: {'CompanyName': 'دنیا اینترنشنال گروپ', 'ApplicationCode': [1,4]}, 
                     83: {'CompanyName': 'تحقیقاتی و تولیدی واریان فارمد (واریان دارو پژوه)', 'ApplicationCode': 1}, 
                     88: {'CompanyName': 'خدمات تحقیقات آرین گستر', 'ApplicationCode': 1}, 
                     90: {'CompanyName': 'گلبرگ غذایی کوروش', 'ApplicationCode': 1}, 
                    107: {'CompanyName': 'ستاره طلایی سینا', 'ApplicationCode': [1,20]}, 
                    112: {'CompanyName': 'پدیده پایدار صنعت ساختمان (پدیده شیمی آرین)', 'ApplicationCode': 1}, 
                    113: {'CompanyName': 'فرآورده های پروتئینی کورش', 'ApplicationCode': 1}, 
                    116: {'CompanyName': 'سروش مانا فارمد', 'ApplicationCode': 1}, 
                    117: {'CompanyName': 'گروه صنایع غذایی پاکبان', 'ApplicationCode': 1}, 
                    119: {'CompanyName': 'دارانامهر', 'ApplicationCode': 1}, 
                    127: {'CompanyName': 'کالا رسان هستی (کوروش پخش)', 'ApplicationCode': 1}, 
                    135: {'CompanyName': 'پاکان به شو', 'ApplicationCode': 4}, 
                    149: {'CompanyName': 'آذر زیست دارو (فاران فارمد، شایان سلامت پارسه)', 'ApplicationCode': 1}, 
                    150: {'CompanyName': 'پدیده زیستی نانو', 'ApplicationCode': 1}, 
                    151: {'CompanyName': 'آوین شیمی پلاست', 'ApplicationCode': 1}, 
                    153: {'CompanyName': 'فرآورده های غذایی پروشات کوروش', 'ApplicationCode': 1}, 
                    156: {'CompanyName': 'گلرنگ ترابر قزوین', 'ApplicationCode': 1}, 
                    159: {'CompanyName': 'دنیا پخش افغانستان', 'ApplicationCode': [1,4]}, 
                    160: {'CompanyName': 'دام و طیور کوروش', 'ApplicationCode': 5}, 
                    167: {'CompanyName': 'افق الکوثر عراق', 'ApplicationCode': 5}, 
                    190: {'CompanyName': 'صنایع غذایی کیلوس', 'ApplicationCode': [1,10]}, 
                    191: {'CompanyName': 'هستی الکتریک تارا', 'ApplicationCode': 1}, 
                    194: {'CompanyName': 'صنایع بهداشتی پایدار', 'ApplicationCode': 1}, 
                    196: {'CompanyName': 'آرین تجارت رایحه آفرین', 'ApplicationCode': 1}, 
                    197: {'CompanyName': 'گندم طلایی کوروش', 'ApplicationCode': 5}, 
                    200: {'CompanyName': 'نوین زعفران', 'ApplicationCode': [1,11]}, 
                    201: {'CompanyName': 'آمیزه های پلیمری سپهر پاک', 'ApplicationCode': [1,10]}, 
                    202: {'CompanyName': 'فارمد سلامت سینا', 'ApplicationCode': 1}, 
                    204: {'CompanyName': 'پژوهش گستران تغذیه آسان', 'ApplicationCode': 1}, 
                    206: {'CompanyName': 'گروه صنعتی سورنا سلولز', 'ApplicationCode': 3}, 
                    212: {'CompanyName': 'شیرینی و شکلات کوروش', 'ApplicationCode': 1}, 
                    274: {'CompanyName': 'پارس فیلم نت', 'ApplicationCode': 3}, 
                    277: {'CompanyName': 'آریو سلولز خزر', 'ApplicationCode': 3}, 
                    357: {'CompanyName': 'سپید ماکیان', 'ApplicationCode': 11}, 
                    406: {'CompanyName': 'ابیان فارمد', 'ApplicationCode': 1}, 
                    413: {'CompanyName': 'سوئیس رز عراق', 'ApplicationCode': 22}, 
                    414: {'CompanyName': 'جویا بهنود', 'ApplicationCode': 1}
                    }


companies_JameMali_dict = {2: {'CompanyName': 'مجتمع تجاری فرهنگی کوروش', 'ApplicationCode': -1, 'Moein': [611133,611144,641171]},
                           6: {'CompanyName': 'سرزمین بازی کوروش', 'ApplicationCode': -1, 'Moein': [611131]},
                           7: {'CompanyName': 'پردیس سینمایی کوروش', 'ApplicationCode': -1, 'Moein': [611131,611132]},
                          20: {'CompanyName': 'ایراندار', 'ApplicationCode': -1, 'Moein': [611141]},
                          22: {'CompanyName': 'آرین سلولز صنعت', 'ApplicationCode': -1, 'Moein': [611133,611131]},
                          30: {'CompanyName': 'مهندسین مشاور افق معادن خاورمیانه', 'ApplicationCode': -1, 'Moein': [611131,641181]},
                          53: {'CompanyName': 'واسپاری ارزش آفرین گلرنگ (لیزینگ)', 'ApplicationCode': -1, 'Moein': [611162,611163,641153,641154]},
                          62: {'CompanyName': 'فاران شیمی', 'ApplicationCode': -1, 'Moein': [611141]},
                          65: {'CompanyName': 'سپهر پلاستیک پدیده', 'ApplicationCode': -1, 'Moein': [611141]},
                         118: {'CompanyName': 'تارا', 'ApplicationCode': -1, 'Moein': [641131,611131,641158,611111,631111]},
                         128: {'CompanyName': 'بیمه سامان', 'ApplicationCode': -1, 'Moein': [611131]},
                         136: {'CompanyName': 'چیکن فامیلی', 'ApplicationCode': -1, 'Moein': [611111,631111]},
                         145: {'CompanyName': 'سبد گردان کوروش', 'ApplicationCode': -1, 'Moein': [611163]},
                         162: {'CompanyName': 'کارکیا سورنا', 'ApplicationCode': -1, 'Moein': [611131]},
                         176: {'CompanyName': 'مریدین روسیه', 'ApplicationCode': -1, 'Moein': [611111,621111,631111]},
                         186: {'CompanyName': 'خدمات خودرویی سامیا', 'ApplicationCode': -1, 'Moein': [611111,611131,611133,631111,621111,641131]},
                         187: {'CompanyName': 'مینیمم مارکت', 'ApplicationCode': -1, 'Moein': [611111,621111,631111]},
                         192: {'CompanyName': 'امین پدیدار', 'ApplicationCode': -1, 'Moein': [611111,611131,621111,631111]},
                         211: {'CompanyName': 'غذای سالم هستی', 'ApplicationCode': -1, 'Moein': [611111,611131,631111]},
                         212: {'CompanyName': 'شیرینی و شکلات کوروش', 'ApplicationCode': -1, 'Moein': [611141,621111]},
                         259: {'CompanyName': 'فاوا فناوری افق', 'ApplicationCode': -1, 'Moein': [611111,611131]},
                         265: {'CompanyName': 'فروتل', 'ApplicationCode': -1, 'Moein': [611131,621131,631111]},
                         266: {'CompanyName': 'خوان گستر هستی', 'ApplicationCode': -1, 'Moein': [611111,621111]},
                         282: {'CompanyName': 'فرا تجارت زردکوه', 'ApplicationCode': -1, 'Moein': [611111]},
                         288: {'CompanyName': 'آرین تامین آفرین', 'ApplicationCode': -1, 'Moein': [611163,611131]},
                         336: {'CompanyName': 'نوین پوش هستی', 'ApplicationCode': -1, 'Moein': [611111]},
                         340: {'CompanyName': 'آی پوش', 'ApplicationCode': -1, 'Moein': [611111,621111,631111]},
                         353: {'CompanyName': 'غذای ناب کوروش', 'ApplicationCode': -1, 'Moein': [611111]},
                         }





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





whole_final = pd.merge(iframe_df_Net , Total_rows_Net_tot_final ,how = 'outer' , right_on= ['CompanyCode'] , left_on =  ['CompanyCode'])
whole_final = pd.merge(whole_final , ALK_df ,how = 'outer' , right_on= ['CompanyCode'] , left_on =  ['CompanyCode'])

print(whole_final.N_NetAmount)
whole_final.N_NetAmount = whole_final.N_NetAmount.fillna(0)
whole_final.O_NetAmount = whole_final.O_NetAmount.fillna(0)
whole_final.X_NetAmount = whole_final.X_NetAmount.fillna(0)
whole_final.ALK_NetAmount = whole_final.ALK_NetAmount.fillna(0)



whole_final['Iframe_Total'] = whole_final['Iframe_Total'].astype('int')
whole_final['N_NetAmount'] = whole_final['N_NetAmount'].astype('int')
whole_final['X_NetAmount'] = whole_final['X_NetAmount'].astype('int')
whole_final['O_NetAmount'] = whole_final['O_NetAmount'].astype('int')
whole_final['ALK_NetAmount'] = whole_final['ALK_NetAmount'].astype('int')





whole_final['TotalSale_Whole'] = whole_final['N_NetAmount'] + whole_final['O_NetAmount'] + whole_final['X_NetAmount'] + whole_final['ALK_NetAmount']
whole_final.drop(['N_NetAmount' , 'O_NetAmount' ,'X_NetAmount' ] , axis = 1 , inplace = True)



whole_final['Differenece'] = whole_final['Iframe_Total'] - whole_final['TotalSale_Whole']
df_anomalies = whole_final[whole_final['Differenece'] > 10]
df_anomalies.drop(['TotalSale_Whole' , 'ALK_NetAmount' , 'Iframe_Total'] ,axis = 1 ,  inplace = True)
df_anomalies.columns = ['نام شرکت' , 'کد مستر شرکت' , 'اختلاف'] #Creating df of anomalies




#Sending Email

import requests
import json

if not df_anomalies.empty:
    payload = json.dumps({

        "token": "P880P2HLA6MO71PWTXTWR",
        "providerId": 6,
        "sendType": 0,
        "email": 'soleimani.yeganeh@golrang.com',
        "ccRecipients": ['Aghdasifam.Masoud@Golrang.com'],
        "subject": "Alert: KPI Values Exceeded!",
        "body": f"موارد زیر مغایرت دارند: \n\n {df_anomalies.to_html(index=False, border=1, justify='center')}",
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
        "subject": "Confirmation of Descending program execution!",
        "body": "انجام شد",
        "fileAttachmentAddress": ""
    })
    headers = {'Content-Type': 'application/json'}
    response = requests.post("https://Esp-api.gig.services/email/saveEmail", headers=headers, data=payload)
    print(response.status_code, response.text)







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

mergedFactExport=pd.merge(Nondistribute_df, NewExport[['CurrencyParityRateDolar' , 'YearMonth' , 'CurrencyID']] , right_on= ['YearMonth' , 'CurrencyID'] , left_on=['YearMonth' , 'CurrencyCode'] , how = 'left')
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

merged_FactExcel = pd.merge(FactExcelAggr , NewExport , left_on = ['YearMonth' , 'CurrencyCode'] , right_on = ['YearMonth' , 'CurrencyID'] , how = 'left')
merged_nonsys = pd.merge(nonsys , NewExport , left_on = ['YearMonth' , 'CurrencyCode'] , right_on = ['YearMonth' , 'CurrencyID'] , how = 'left')
merged_ALK = pd.merge(ALK_df_R , NewExport[['CurrencyParityRateDolar' ,'YearMonth', 'CurrencyID']] , left_on = ['YearMonth' , 'CurrencyCode'] , right_on = ['YearMonth' , 'CurrencyID'] , how = 'left' )


## Check Marjuei
merged_FactExcel['RealAmount_Dollar'] = merged_FactExcel['CurrencyParityRateDolar'] * merged_FactExcel['RealAmount']
merged_nonsys['RealAmount_Dollar_sys'] =  merged_nonsys['CurrencyParityRateDolar'] * merged_nonsys['RealAmount']
merged_ALK['RealAmount_Dollar_ALK'] = merged_ALK['CurrencyParityRateDolar'] * merged_nonsys['RealAmount']



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
print(merged_All.columns)




#################################### Fill Na ##########################


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

##############################Sending Emails for real amount ##############################

if not merged_All_Final_Anomalies.empty:
    payload = json.dumps({

        "token": "P880P2HLA6MO71PWTXTWR",
        "providerId": 6,
        "sendType": 0,
        "email": 'soleimani.yeganeh@golrang.com',
        "ccRecipients": ['Aghdasifam.Masoud@Golrang.com'],
        "subject": "Alert: KPI Values Exceeded!",
        "body": f"موارد زیر مغایرت دارند: \n\n {merged_All_Final_Anomalies.to_html(index=False, border=1, justify='center')}",
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
        "subject": "Confirmation of Export_RealAmount execution!",
        "body": "انجام شد",
        "fileAttachmentAddress": ""
    })
    headers = {'Content-Type': 'application/json'}
    response = requests.post("https://Esp-api.gig.services/email/saveEmail", headers=headers, data=payload)
    print(response.status_code, response.text)



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
status = insert_to_sql(merged_All_Final, '172.31.31.29', 'DataQuality')
print("Insert status:", status)




