import requests
import json
from datetime import datetime, timedelta
import pandas as pd
import gspread

url = "http://13.126.125.132:4000"

payload = "{\"query\":\"query suggi_GetSaleWiseProfit { suggi_getSaleWiseProfit { _id customerDetails { name onboarding_date  village { name }  pincode { code } cust_type GSTIN iscouponapplicable address  { city street  pincode } phone customer_uid} storeDetails { name address {   pincode } vertical territory { name zone { name } } } userDetails { name address {   pincode } } invoiceno invoicedate product { soldqty servicecharge extradiscount sellingprice gst purchaseProductDetails {   name   category {     name   }   manufacturer {     name   }   sub_category {     name   } } lotDetails {  discount rate id sellingprice invoice { supplier_ref invoiceno createddate invoicedate } cnStockproduct subqty transportCharges otherCharges landingrate } supplierDetails { name } storeTargetPerProduct } payment { card cash upi } grosstotal storeCost } } \",\"variables\":{}}"
headers = {
  'Content-Type': 'application/json'
}

response = requests.post(url, headers=headers, data=payload)
data = response.json()["data"]["suggi_getSaleWiseProfit"]
df = pd.DataFrame(data)

start_date_str = "2024-07-22"
end_date_str = "2024-07-29"
start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
month = start_date.month
year = start_date.year

start_date_week = start_date
start_date_month = datetime(2024, month, 1)
start_date_year = datetime(year, 4, 1)
start_date_week_str = start_date_week.strftime("%Y-%m-%dT00:00:00Z")
start_date_month_str = start_date_month.strftime("%Y-%m-%dT00:00:00Z")
start_date_year_str = start_date_year.strftime("%Y-%m-%dT00:00:00Z")

end_date = (datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)).strftime("%Y-%m-%dT23:59:59Z")
start_dates = [start_date_week_str, start_date_month_str, start_date_year_str]

revenue = []    
for date_str in start_dates:    
      df['invoicedate'] = pd.to_datetime(df['invoicedate'])
      
      start_date = pd.to_datetime(date_str)
      end_date = pd.to_datetime(end_date)
      
      
      filtered_df = df[(df['invoicedate'] >= start_date) & (df['invoicedate'] <= end_date)]
      df1 = filtered_df.explode(['product'])
      df1['sellingprice_unit'] = df1['product'].apply(lambda x: x['sellingprice'])
      df1['servicecharge_unit'] = df1['product'].apply(lambda x: x['servicecharge'])
      df1['soldqty'] = df1['product'].apply(lambda x: x['soldqty'])

      df1['gst%_unit'] = df1['product'].apply(lambda x: x['gst'])
      df1['sellingprice_unit_xgst'] = df1['sellingprice_unit'] / (1 + (df1['gst%_unit'] / 100))
      df1['servicecharge_unit_xgst'] = df1['servicecharge_unit'] / 1.18
      df1['sale_value_xgst'] = df1['soldqty'] * (df1['sellingprice_unit_xgst'] + df1['servicecharge_unit_xgst'])
      sum = (df1['sale_value_xgst'].sum())
      revenue.append(sum)
      
revenue_float = [round(float(x),2) for x in revenue]
print(revenue_float)