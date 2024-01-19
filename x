import pandas as pd
import sqlalchemy
import numpy as np
import datetime as dt
import requests_kerberos as rkrb
import requests

def get_last_workday():
today = dt.date.today()
one_day = dt.timedelta(days=1)
while today.weekday() >= 5: # Saturday (5) or Sunday (6)
today -= one_day
return today - one_day # Subtract one day to get the last workday

# Calculate the last workday - 1
last_workday_minus_one = get_last_workday()

# Convert the last workday - 1 to a string
last_workday_minus_one_str = last_workday_minus_one.strftime("%Y-%m-%d")

server = ****
port = ****
database = ****
driver = ****
# driver = ****
connection_string = f"mssql+pyodbc://{server}:{port}/{database}?trusted_connection=yes&driver={driver}"
engine = sqlalchemy.create_engine(connection_string)

# Get Data From Med_bi
sql_query = "SELECT Linkage_Execution, entity, Deal#, Deal_date, date_expo, Credit_due_date, Credit_pty_trf, Credit_Term, Linkage#, Trader, [sale_due_amt $], " \
"[purchase_due_amt $],[sale_due_amt $] - [purchase_due_amt $] AS [NET], [Pay_risk $ (pty -8d)], [Pay_risk_secured $ (pty-8d)], [Unit price$], [purchase_gbl_amt $], [sale_gbl_amt $], Quantity " \
"FROM MED_BI.dbo.credit_phy_deal WHERE date_expo = (SELECT MAX(date_expo) FROM MED_BI.dbo.credit_phy_deal)"
df = pd.read_sql(sql_query, engine)

# Removing spaces in the column Deal#
df['Deal#'] = df['Deal#'].str.replace(' ', '')

sql_query_2 = "SELECT contract_split, physical_risk, mtm_formula, qty_b, cpt, cpt_grp, date_prop_status," \
"cpt_country, cpt_group_country, [contractual execution] FROM MED_BI.dbo.Tradphys_header_split"
df_2 = pd.read_sql(sql_query_2, engine)

# Merging both tables
df_3 = pd.merge(df, df_2, left_on= 'Deal#', right_on='contract_split', how="inner")

# Filter column physical_risk with starts with 'RIN'
df_4 = df_3[df_3['physical_risk'].fillna('').str.startswith('RIN')]

# Add a column for month names
credit_pty_trf_index = df_4.columns.get_loc("Credit_pty_trf")
df_4.insert(credit_pty_trf_index + 1, "Month_pty_trf", pd.to_datetime(df_4['Credit_pty_trf']).dt.month)

# Add a column for year names
credit_pty_trf_index = df_4.columns.get_loc("Credit_pty_trf")
df_4.insert(credit_pty_trf_index + 1, "Year_pty_trf", pd.to_datetime(df_4['Credit_pty_trf']).dt.year)

# Add a column to df4 called RIN, keeping the first 6 characters of physical_risk
df_4.loc[:, 'RIN'] = df_4['physical_risk'].apply(lambda x: x[:6])


# Define a function to calculate the deadline based on the Deal_date month
def calculate_deadline(deal_date):
if deal_date.month in [1, 2, 3]:
return deal_date.replace(month=6, day=1)
elif deal_date.month in [4, 5, 6]:
return deal_date.replace(month=9, day=1)
elif deal_date.month in [7, 8, 9]:
return deal_date.replace(month=12, day=1)
else:
# For October, November, December
return deal_date.replace(year=deal_date.year + 1, month=3, day=31)


# Apply the function to create the 'Deadline_Rin' column
df_4['Deadline_Rin'] = df_4['Deal_date'].apply(calculate_deadline)

# Get RIN D3 data from Medeco
auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D3_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=true"
response = requests.get(url_to_call, auth=auth, verify=False)
d3 = pd.DataFrame(response.json()['values'])
d3['RIN'] = 'RIN D3'

# Get RIN D4 data from Medeco
auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D4_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=true"
response = requests.get(url_to_call, auth=auth, verify=False)
d4 = pd.DataFrame(response.json()['values'])
d4['RIN'] = 'RIN D4'

# Get RIN D5 data from Medeco
auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D5_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=true"
response = requests.get(url_to_call, auth=auth, verify=False)
d5 = pd.DataFrame(response.json()['values'])
d5['RIN'] = 'RIN D5'

# Get RIN D6 data from Medeco
auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D6_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=true"
response = requests.get(url_to_call, auth=auth, verify=False)
d6 = pd.DataFrame(response.json()['values'])
d6['RIN'] = 'RIN D6'

# Merge all dataframes on RIN with pd.merge
dfx = pd.concat([d3, d4, d5, d6], ignore_index=True)

# Convert 'value' column to numeric
dfx['value'] = pd.to_numeric(dfx['value'])

# add a column "Market_price_RIN_B" = 'value' * 42 (gallon to BBL)
dfx['Market_price_RIN_B'] = dfx['value'] * 42 / 100

# rename value to market_price_RIN_gal
dfx.rename(columns={'value': 'Market_price_RIN_gal'}, inplace=True)

# divide Market_price_RIN_gal by 100 (cent to USD)
dfx['Market_price_RIN_gal'] = dfx['Market_price_RIN_gal'] / 100

# rename date to workday - 1
dfx.rename(columns={'date': 'workday_minus_one'}, inplace=True)

# Merge df_4 and dfx
df_4 = pd.merge(df_4, dfx, on='RIN', how="inner")

# Add a column "Market_price" = Market_price_RIN_B * qty_b
df_4['Market_price'] = df_4['Market_price_RIN_B'] * df_4['qty_b']

# Add a column "MTM_purchase_gbl" = Market_price - purchase_gbl_amt $ when purchase_gbl_amt $ is not 0
df_4['MTM_purchase_gbl'] = df_4.apply(
lambda row: np.nan if row['purchase_gbl_amt $'] == 0 else row['Market_price'] - row['purchase_gbl_amt $'],
axis=1)

# Add a column "MTM_sale_gbl" = sale_gbl_amt $ - Market_price when sale_gbl_amt $ is not 0
df_4['MTM_sale_gbl'] = df_4.apply(
lambda row: np.nan if row['sale_gbl_amt $'] == 0 else row['sale_gbl_amt $'] - row['Market_price'], axis=1)

# Fill NaN values with 0 to prevent issues in further calculations
df_4['MTM_purchase_gbl'].fillna(0, inplace=True)
df_4['MTM_sale_gbl'].fillna(0, inplace=True)

# Add a "MTM_NET" column = MTM_sale_gbl - MTM_purchase_gbl
df_4['MTM_NET'] = df_4['MTM_sale_gbl'] + df_4['MTM_purchase_gbl']

# Add "performance_risk" column = MTM_NET when MTM_NET > 0, 0 otherwise
df_4['performance_risk'] = df_4.apply(lambda row: row['MTM_NET'] if row['MTM_NET'] > 0 else 0, axis=1)

print(df_4)
