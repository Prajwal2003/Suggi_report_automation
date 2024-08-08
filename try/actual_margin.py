import requests
import json
from datetime import datetime, timedelta
import pandas as pd

url = "http://13.126.125.132:4000"

payload = "{\"query\":\"query suggi_GetSaleWiseProfit { suggi_getSaleWiseProfit { _id customerDetails { name onboarding_date  village { name }  pincode { code } cust_type GSTIN iscouponapplicable address  { city street  pincode } phone customer_uid} storeDetails { name address {   pincode } vertical territory { name zone { name } } } userDetails { name address {   pincode } } invoiceno invoicedate product { soldqty servicecharge extradiscount sellingprice gst purchaseProductDetails {   name   category {     name   }   manufacturer {     name   }   sub_category {     name   } } lotDetails {  discount rate id sellingprice invoice { supplier_ref invoiceno createddate invoicedate } cnStockproduct subqty transportCharges otherCharges landingrate } supplierDetails { name } storeTargetPerProduct } payment { card cash upi } grosstotal storeCost } } \",\"variables\":{}}"
headers = {
  'Content-Type': 'application/json'
}

response = requests.post(url, headers=headers, data=payload)
data = response.json()["data"]["suggi_getSaleWiseProfit"]
df = pd.DataFrame(data)

# Convert invoicedate to datetime format
df['invoicedate'] = pd.to_datetime(df['invoicedate'])

# Date range input
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

actual_margins = []

for date_str in start_dates:
    start_date = pd.to_datetime(date_str)
    end_date = pd.to_datetime(end_date)

    # Filter data by date range
    filtered_df = df[(df['invoicedate'] >= start_date) & (df['invoicedate'] <= end_date)]
    
    # Explode the product list
    df_exploded = filtered_df.explode('product')
    
    # Normalize product details
    df_product = pd.json_normalize(df_exploded['product'])
    print(df_product['lotDetails.otherCharges'])
    
    df_product['sellingprice_unit_xgst'] = df_product['sellingprice'] / (1 + (df_product['gst'] / 100))
    df_product['servicecharge_unit_xgst'] = df_product['servicecharge'] / (1 + (df_product['gst'] / 100))
    df_product['sale_value_xgst'] = df_product['soldqty'] * (df_product['sellingprice_unit_xgst'] + df_product['servicecharge_unit_xgst'])
    print(df_product.columns.tolist())
    # Normalize lot details for purchase value calculations
    #df_lot = pd.json_normalize(df_product['lotDetails'])
    
    df_product['purchasingprice_unit'] = df_product['lotDetails.rate']
    df_product['lotDetails.discount'] = df_product['lotDetails.discount']
    df_product['purchase_value_xgst'] = df_product['soldqty'] * (df_product['purchasingprice_unit'] - df_product['lotDetails.discount'])
    
    # Calculate actual margin
    df_product['actual_margin'] = df_product['sale_value_xgst'] - (
        df_product['purchase_value_xgst'] -
        (df_product['soldqty'] * (df_product['lotDetails.cnStockproduct'] / df_product['lotDetails.subqty'])) +
        (df_product['soldqty'] * (df_product['lotDetails.transportCharges'] / df_product['lotDetails.subqty'])) +
        (df_product['soldqty'] * (df_product['lotDetails.otherCharges'] / df_product['lotDetails.subqty']))
    )
    
    actual_margins.append(df_product['actual_margin'].sum())

actual_margins_float = [round(float(x), 2) for x in actual_margins]
print(actual_margins_float)
