import requests
import json
from datetime import datetime, timedelta
import pandas as pd


url = "http://13.126.125.132:4000"

payload = "{\"query\":\"query Suggi_StoreTarget {\\n  suggi_StoreTarget {\\n    Store\\n    TM\\n    Category\\n    Month\\n    Year\\n    Target\\n    Date\\n    Daily_Target\\n  }\\n}\",\"variables\":{}}"
headers = {
  'Content-Type': 'application/json'
}

response = requests.post(url, headers=headers, data=payload)
data = response.json()["data"]["suggi_StoreTarget"]
df = pd.DataFrame(data)
print(df.columns.to_list())

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
    df['Date'] = pd.to_datetime(df['Date'])
    
    start_date = pd.to_datetime(date_str)
    end_date = pd.to_datetime(end_date)
    
    
    filtered_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
    grouped_df = df.groupby('TM')
    for x,y in grouped_df:
      print(x)
      print(y)
      
sum = (grouped_df['Daily_Target'].sum())
      
revenu = []
revenue.append(sum)
for x in revenue:
  x = float(x)/100000
  revenu.append(x)
print(revenu)
