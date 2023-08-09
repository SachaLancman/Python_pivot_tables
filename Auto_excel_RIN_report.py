import pandas as pd
import sqlalchemy
import numpy as np
import datetime as dt
import requests_kerberos as rkrb
import requests
import win32com.client as win32

def get_last_workday():
    today = dt.date.today()
    one_day = dt.timedelta(days=1)
    while today.weekday() >= 5:  # Saturday (5) or Sunday (6)
        today -= one_day
    return today - one_day  # Subtract one day to get the last workday

# Calculate the last workday - 1
last_workday_minus_one = get_last_workday()

# Convert the last workday - 1 to a string
last_workday_minus_one_str = last_workday_minus_one.strftime("%Y-%m-%d")

def create_RIN_report():
    server = "xxx"
    port = xxx
    database = "xxx"
    driver = "xxx"
    # driver = "SQL Server"
    connection_string = f"mssql+pyodbc://{server}:{port}/{database}?trusted_connection=yes&driver={driver}"
    engine = sqlalchemy.create_engine(connection_string)

    # Get Data From Med_bi
    sql_query = "SELECT Linkage_Execution, entity, Deal#, Deal_date, date_expo, Credit_due_date, Credit_pty_trf, Credit_Term, Linkage#, Trader, [sale_due_amt $], " \
                "[purchase_due_amt $],[sale_due_amt $] - [purchase_due_amt $] AS [NET], [Pay_risk $ (pty -8d)], [Pay_risk_secured $ (pty-8d)], [Unit price$], [purchase_gbl_amt $], [sale_gbl_amt $], Quantity " \
                "FROM MED_BI.dbo.credit_phy_deal WHERE date_expo = (SELECT MAX(date_expo) FROM MED_BI.dbo.credit_phy_deal)"
    df = pd.read_sql(sql_query, engine)

    sql_query_2 = "SELECT contract_split, physical_risk, mtm_formula, qty_mtm_open_b, qty_mtm_open_mt, qty_mtm_open_m3, cpt, cpt_grp," \
                  "cpt_country, cpt_group_country FROM MED_BI.dbo.Tradphys_header_split"
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
    df_4['RIN'] = df_4['physical_risk'].apply(lambda x: x[:6])

    # Get RIN D3 data from Medeco
    auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
    url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D3_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=false"
    response = requests.get(url_to_call, auth=auth, verify=False)
    d3 = pd.DataFrame(response.json()['values'])
    d3['RIN'] = 'RIN D3'

    # Get RIN D4 data from Medeco
    auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
    url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D4_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=false"
    response = requests.get(url_to_call, auth=auth, verify=False)
    d4 = pd.DataFrame(response.json()['values'])
    d4['RIN'] = 'RIN D4'

    # Get RIN D5 data from Medeco
    auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
    url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D5_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=false"
    response = requests.get(url_to_call, auth=auth, verify=False)
    d5 = pd.DataFrame(response.json()['values'])
    d5['RIN'] = 'RIN D5'

    # Get RIN D6 data from Medeco
    auth = rkrb.HTTPKerberosAuth(mutual_authentication=rkrb.DISABLED, force_preemptive=True)
    url_to_call = f"https://api-medeco.dts.corp.local/api/series/mu.RIN_D6_2023A.MLI/values?from={last_workday_minus_one_str}&to={last_workday_minus_one_str}&frequency=2&exclude_na=false"
    response = requests.get(url_to_call, auth=auth, verify=False)
    d6 = pd.DataFrame(response.json()['values'])
    d6['RIN'] = 'RIN D6'

    # Merge all dataframes on RIN with pd.merge
    dfx = pd.concat([d3, d4, d5, d6], ignore_index=True)

    # add a column "Market_price_RIN_B" =  'value' * 42 (gallon to BBL)
    dfx['Market_price_RIN_B'] = dfx['value'] * 42 / 100

    # rename value to market_price_RIN_gal
    dfx.rename(columns={'value': 'Market_price_RIN_gal'}, inplace=True)

    # divide Market_price_RIN_gal by 100 (cent to USD)
    dfx['Market_price_RIN_gal'] = dfx['Market_price_RIN_gal'] / 100

    # rename date to workday - 1
    dfx.rename(columns={'date': 'workday_minus_one'}, inplace=True)

    # Merge df_4 and dfx
    df_4 = pd.merge(df_4, dfx, on='RIN', how="inner")

    # Add a column "Market_price" = Market_price_RIN_B * Quantity
    df_4['Market_price'] = df_4['Market_price_RIN_B'] * df_4['Quantity']

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
    df_4['MTM_NET'] = df_4['MTM_sale_gbl'] - df_4['MTM_purchase_gbl']

    # Save table to Excel
    df_4.to_excel('auto_rin.xlsx', index=False)

    # Create the Excel application object
    xlApp = win32.Dispatch('Excel.Application')
    xlApp.Visible = True

    # Open the workbook
    wb = xlApp.Workbooks.Open(r'C:\Users\SLANCM\PycharmProjects\python_Project_1\auto_rin.xlsx')
    ws_data = wb.Worksheets(1)

    # Convert the table to a real Excel table
    table_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(df_4.shape[0] + 1, df_4.shape[1]))
    table = ws_data.ListObjects.Add(1, table_range, 1, 1)

    # Change the style of the table
    table.TableStyle = "TableStyleMedium10"

    # Format the currency fields in the first sheet
    currency_fields = ['purchase_due_amt $', 'sale_due_amt $', 'NET', 'Pay_risk $ (pty -8d)', 'Pay_risk_secured $ (pty-8d)',
                       'Unit price$', 'purchase_gbl_amt $', 'sale_gbl_amt $', 'Market_price_RIN_gal', 'Market_price_RIN_B',
                       'Market_price', 'MTM_purchase_gbl', 'MTM_sale_gbl', 'MTM_NET']
    for field in currency_fields:
        column_index = df_4.columns.get_loc(field) + 1  # Get the column index (1-based)
        ws_data.Columns(column_index).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    # Auto-fit the columns in the first sheet
    ws_data.UsedRange.Columns.AutoFit()

    # Clear pivot tables on the Report tab
    def clear_pts(ws):
        for pt in ws.PivotTables():
            pt.TableRange2.Clear()

    # Create the first pivot table
    ws_report1 = wb.Worksheets.Add()
    clear_pts(ws_report1)
    pt_cache1 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt1 = pt_cache1.CreatePivotTable(ws_report1.Range("B3"), "PivotTable1")

    # Insert pivot table field settings for the first pivot table
    def insert_pt_field_set1(pt):
        field_filters = {}
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['physical_risk'] = pt.PivotFields("physical_risk")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['cpt'] = pt.PivotFields("cpt")
        field_rows['Deal#'] = pt.PivotFields("Deal#")

        field_values = {}
        field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
        field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
        field_values['NET'] = pt.PivotFields("NET")
        field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

        # Insert filter fields
        field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Credit_pty_trf'].Position = 1

        field_filters['physical_risk'].Orientation = 3  # xlPageField
        field_filters['physical_risk'].Position = 2

        field_filters['Credit_Term'].Orientation = 3  # xlPageField
        field_filters['Credit_Term'].Position = 3

        field_filters['Trader'].Orientation = 3  # xlPageField
        field_filters['Trader'].Position = 4

        field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Month_pty_trf'].Position = 5

        field_filters['Year_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Year_pty_trf'].Position = 6

        # Insert row field
        field_rows['cpt'].Orientation = 1
        field_rows['cpt'].Position = 1

        field_rows['Deal#'].Orientation = 1
        field_rows['Deal#'].Position = 2

        # Insert values field
        field_values['sale_due_amt $'].Orientation = 4
        field_values['sale_due_amt $'].Function = -4157  # xlSum
        field_values['sale_due_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['purchase_due_amt $'].Orientation = 4
        field_values['purchase_due_amt $'].Function = -4157
        field_values['purchase_due_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['NET'].Orientation = 4
        field_values['NET'].Function = -4157  # xlSum
        field_values['NET'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['Pay_risk $ (pty -8d)'].Orientation = 4
        field_values['Pay_risk $ (pty -8d)'].Function = -4157  # xlSum
        field_values['Pay_risk $ (pty -8d)'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    insert_pt_field_set1(pt1)

    # Change pivot table style for the first pivot table
    pt1.TableStyle2 = "PivotStyleDark3"

    # Adjust column widths for the first pivot table
    ws_report1.Columns("D:E").ColumnWidth = 30

    # Create a slicer for "Month_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt1, "Month_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report1)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 700
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt1, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report1)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 95
    slicer.Left = 700
    slicer.Top = 315

    # Create a slicer for "physical_risk"
    slicer_cache = wb.SlicerCaches.Add(pt1, "physical_risk")
    slicer = slicer_cache.Slicers.Add(ws_report1)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 700
    slicer.Top = 415

    # Create the second pivot table
    ws_report2 = wb.Worksheets.Add()
    clear_pts(ws_report2)
    pt_cache2 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt2 = pt_cache2.CreatePivotTable(ws_report2.Range("B3"), "PivotTable2")

    # Insert pivot table field settings for the second pivot table
    def insert_pt_field_set2(pt):
        field_filters = {}
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['cpt'] = pt.PivotFields("cpt")
        field_rows['Deal#'] = pt.PivotFields("Deal#")
        field_rows['physical_risk'] = pt.PivotFields("physical_risk")

        field_values = {}
        field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
        field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
        field_values['NET'] = pt.PivotFields("NET")
        field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

        # Insert filter fields
        field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Credit_pty_trf'].Position = 1

        field_filters['Credit_Term'].Orientation = 3  # xlPageField
        field_filters['Credit_Term'].Position = 2

        field_filters['Trader'].Orientation = 3  # xlPageField
        field_filters['Trader'].Position = 3

        field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Month_pty_trf'].Position = 4

        field_filters['Year_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Year_pty_trf'].Position = 5

        # Insert row field
        field_rows['cpt'].Orientation = 1
        field_rows['cpt'].Position = 1

        field_rows['Deal#'].Orientation = 1
        field_rows['Deal#'].Position = 2

        field_rows['physical_risk'].Orientation = 1
        field_rows['physical_risk'].Position = 3

        # Insert values field
        field_values['sale_due_amt $'].Orientation = 4
        field_values['sale_due_amt $'].Function = -4157  # xlSum
        field_values['sale_due_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['purchase_due_amt $'].Orientation = 4
        field_values['purchase_due_amt $'].Function = -4157  # xlSum
        field_values['purchase_due_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['NET'].Orientation = 4
        field_values['NET'].Function = -4157  # xlSum
        field_values['NET'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['Pay_risk $ (pty -8d)'].Orientation = 4
        field_values['Pay_risk $ (pty -8d)'].Function = -4157  # xlSum
        field_values['Pay_risk $ (pty -8d)'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    insert_pt_field_set2(pt2)

    # Change pivot table style for the second pivot table
    pt2.TableStyle2 = "PivotStyleDark3"

    # Adjust column widths for the second pivot table
    ws_report2.Columns("D:E").ColumnWidth = 30

    # Create a slicer for "Month_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt2, "Month_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report2)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 700
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt2, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report2)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 95
    slicer.Left = 700
    slicer.Top = 315

    # Create a slicer for "physical_risk"
    slicer_cache = wb.SlicerCaches.Add(pt2, "physical_risk")
    slicer = slicer_cache.Slicers.Add(ws_report2)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 700
    slicer.Top = 415

    # Create the third pivot table
    ws_report3 = wb.Worksheets.Add()
    clear_pts(ws_report3)
    pt_cache3 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt3 = pt_cache3.CreatePivotTable(ws_report3.Range("B3"), "PivotTable2")

    # Insert pivot table field settings for the third pivot table
    def insert_pt_field_set3(pt):
        field_filters = {}
        field_filters['physical_risk'] = pt.PivotFields("physical_risk")
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['cpt'] = pt.PivotFields("cpt")

        field_values = {}
        field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
        field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
        field_values['NET'] = pt.PivotFields("NET")
        field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

        # Insert filter fields
        field_filters['physical_risk'].Orientation = 3  # xlPageField
        field_filters['physical_risk'].Position = 1

        field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Credit_pty_trf'].Position = 2

        field_filters['Credit_Term'].Orientation = 3  # xlPageField
        field_filters['Credit_Term'].Position = 3

        field_filters['Trader'].Orientation = 3  # xlPageField
        field_filters['Trader'].Position = 4

        field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Month_pty_trf'].Position = 5

        field_filters['Year_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Year_pty_trf'].Position = 6

        # Insert row field
        field_rows['cpt'].Orientation = 1
        field_rows['cpt'].Position = 1

        # Insert values field
        field_values['sale_due_amt $'].Orientation = 4
        field_values['sale_due_amt $'].Function = -4157  # xlSum
        field_values['sale_due_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['purchase_due_amt $'].Orientation = 4
        field_values['purchase_due_amt $'].Function = -4157  # xlSum
        field_values['purchase_due_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['NET'].Orientation = 4
        field_values['NET'].Function = -4157  # xlSum
        field_values['NET'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['Pay_risk $ (pty -8d)'].Orientation = 4
        field_values['Pay_risk $ (pty -8d)'].Function = -4157  # xlSum
        field_values['Pay_risk $ (pty -8d)'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    insert_pt_field_set3(pt3)

    # Change pivot table style for the second pivot table
    pt3.TableStyle2 = "PivotStyleDark3"

    # Adjust column widths for the second pivot table
    ws_report3.Columns("D:E").ColumnWidth = 30

    # Create a slicer for "Month_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt3, "Month_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report3)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 700
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt3, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report3)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 95
    slicer.Left = 700
    slicer.Top = 315

    # Create a slicer for "physical_risk"
    slicer_cache = wb.SlicerCaches.Add(pt3, "physical_risk")
    slicer = slicer_cache.Slicers.Add(ws_report3)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 700
    slicer.Top = 415

    # Create the fourth pivot table
    ws_report4 = wb.Worksheets.Add()
    clear_pts(ws_report4)
    pt_cache4 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt4 = pt_cache4.CreatePivotTable(ws_report4.Range("B3"), "PivotTable2")

    # Insert pivot table field settings for the third pivot table
    def insert_pt_field_set4(pt):
        field_filters = {}
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['cpt'] = pt.PivotFields("cpt")
        field_rows['Deal#'] = pt.PivotFields("Deal#")

        field_values = {}
        field_values['sale_gbl_amt $'] = pt.PivotFields("sale_gbl_amt $")
        field_values['purchase_gbl_amt $'] = pt.PivotFields("purchase_gbl_amt $")
        field_values['Market_price'] = pt.PivotFields("Market_price")
        field_values['MTM_purchase_gbl'] = pt.PivotFields("MTM_purchase_gbl")
        field_values['MTM_sale_gbl'] = pt.PivotFields("MTM_sale_gbl")
        field_values['MTM_NET'] = pt.PivotFields("MTM_NET")

        # Insert filter fields
        field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Credit_pty_trf'].Position = 1

        field_filters['Credit_Term'].Orientation = 3  # xlPageField
        field_filters['Credit_Term'].Position = 2

        field_filters['Trader'].Orientation = 3  # xlPageField
        field_filters['Trader'].Position = 3

        field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Month_pty_trf'].Position = 4

        field_filters['Year_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Year_pty_trf'].Position = 5

        # Insert row field
        field_rows['cpt'].Orientation = 1
        field_rows['cpt'].Position = 1

        field_rows['Deal#'].Orientation = 1
        field_rows['Deal#'].Position = 2

        # Insert values field
        field_values['sale_gbl_amt $'].Orientation = 4
        field_values['sale_gbl_amt $'].Function = -4157  # xlSum
        field_values['sale_gbl_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['purchase_gbl_amt $'].Orientation = 4
        field_values['purchase_gbl_amt $'].Function = -4157  # xlSum
        field_values['purchase_gbl_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['Market_price'].Orientation = 4
        field_values['Market_price'].Function = -4157  # xlSum
        field_values['Market_price'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['MTM_sale_gbl'].Orientation = 4
        field_values['MTM_sale_gbl'].Function = -4157  # xlSum
        field_values['MTM_sale_gbl'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['MTM_purchase_gbl'].Orientation = 4
        field_values['MTM_purchase_gbl'].Function = -4157  # xlSum
        field_values['MTM_purchase_gbl'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['MTM_NET'].Orientation = 4
        field_values['MTM_NET'].Function = -4157  # xlSum
        field_values['MTM_NET'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    insert_pt_field_set4(pt4)

    # Change pivot table style for the second pivot table
    pt4.TableStyle2 = "PivotStyleDark8"

    # Adjust column widths for the second pivot table
    ws_report4.Columns("D:E").ColumnWidth = 30

    # Create a slicer for "Month_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt4, "Month_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report4)
    slicer.Style = "SlicerStyleOther1"
    slicer.Height = 220
    slicer.Left = 850
    slicer.Top = 90

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt4, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report4)
    slicer.Style = "SlicerStyleOther1"
    slicer.Height = 95
    slicer.Left = 850
    slicer.Top = 315

    # Create a slicer for "Physical_risk"
    slicer_cache = wb.SlicerCaches.Add(pt4, "physical_risk")
    slicer = slicer_cache.Slicers.Add(ws_report4)
    slicer.Style = "SlicerStyleOther1"
    slicer.Height = 240
    slicer.Left = 850
    slicer.Top = 415

    # Rename the sheets
    ws_data.Name = 'table'
    ws_report1.name = 'counterpaties + deal#'
    ws_report2.name = 'counterpaties + deal# + RINs'
    ws_report3.name = 'counterpaties'
    ws_report4.name = 'MTM'

    # Get the range of the "MTM_purchase_gbl" and "MTM_sale_gbl" columns
    mtm_purchase_gbl_range = ws_report4.Range(
        "F9:F" + str(ws_report4.Cells(ws_report4.Rows.Count, "F").End(-4162).Row))
    mtm_sale_gbl_range = ws_report4.Range("G9:G" + str(ws_report4.Cells(ws_report4.Rows.Count, "G").End(-4162).Row))
    mtm_net_range = ws_report4.Range("H9:H" + str(ws_report4.Cells(ws_report4.Rows.Count, "H").End(-4162).Row))

    # Apply conditional formatting to "MTM_purchase_gbl" column
    mtm_purchase_gbl_range.FormatConditions.Add(Type=1, Operator=5, Formula1="5")
    mtm_purchase_gbl_range.FormatConditions(mtm_purchase_gbl_range.FormatConditions.Count).SetFirstPriority()
    mtm_purchase_gbl_range.FormatConditions(1).Interior.Color = 144238260  # Green

    mtm_purchase_gbl_range.FormatConditions.Add(Type=1, Operator=6, Formula1="-5")
    mtm_purchase_gbl_range.FormatConditions(mtm_purchase_gbl_range.FormatConditions.Count).SetFirstPriority()
    mtm_purchase_gbl_range.FormatConditions(1).Interior.Color = 13421823  # Red

    # Apply conditional formatting to "MTM_sale_gbl" column
    mtm_sale_gbl_range.FormatConditions.Add(Type=1, Operator=5, Formula1="5")
    mtm_sale_gbl_range.FormatConditions(mtm_sale_gbl_range.FormatConditions.Count).SetFirstPriority()
    mtm_sale_gbl_range.FormatConditions(1).Interior.Color = 144238260  # Green

    mtm_sale_gbl_range.FormatConditions.Add(Type=1, Operator=6, Formula1="-5")
    mtm_sale_gbl_range.FormatConditions(mtm_sale_gbl_range.FormatConditions.Count).SetFirstPriority()
    mtm_sale_gbl_range.FormatConditions(1).Interior.Color = 13421823  # Red

    # Apply conditional formatting to "MTM_NET" column
    mtm_net_range.FormatConditions.Add(Type=1, Operator=5, Formula1="5")
    mtm_net_range.FormatConditions(mtm_net_range.FormatConditions.Count).SetFirstPriority()
    mtm_net_range.FormatConditions(1).Interior.Color = 144238260  # Green

    mtm_net_range.FormatConditions.Add(Type=1, Operator=6, Formula1="-5")
    mtm_net_range.FormatConditions(mtm_net_range.FormatConditions.Count).SetFirstPriority()
    mtm_net_range.FormatConditions(1).Interior.Color = 13421823  # Red

    # Save the modified workbook
    wb.Save()

    print("Tables imported and pivot tables created in 'auto_rin.xlsx'")

if __name__ == '__main__':
    create_RIN_report()
