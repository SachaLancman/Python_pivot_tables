import pandas as pd
import sqlalchemy
import win32com.client as win32

server = "medissys.bi.dts.corp.local"
port = 1435
database = "med_bi"
driver = "ODBC Driver 13 for SQL Server"
#driver = "SQL Server"
connection_string = f"mssql+pyodbc://{server}:{port}/{database}?trusted_connection=yes&driver={driver}"
engine = sqlalchemy.create_engine(connection_string)

# Get Data From Med_bi
sql_query = "SELECT id_counter, Linkage_Execution, entity, Deal#, Deal_date, date_expo, Credit_due_date, Credit_pty_trf, Credit_Term, Linkage#, Trader, [sale_due_amt $], " \
            "[purchase_due_amt $],[sale_due_amt $] - [purchase_due_amt $] AS [NET], [Pay_risk $ (pty -8d)], [Unit price$], [purchase_gbl_amt $], [sale_gbl_amt $], Quantity " \
            "FROM MED_BI.dbo.credit_phy_deal WHERE Credit_pty_trf BETWEEN '2023-07-01' AND '2023-12-31' AND date_expo = (SELECT MAX(date_expo) FROM MED_BI.dbo.credit_phy_deal)"
df = pd.read_sql(sql_query, engine)

sql_query_2 = "SELECT id_counter, Name_full, Group_name_full FROM MED_BI.dbo.COUNTERPARTY"
df_2 = pd.read_sql(sql_query_2, engine)

url = 'https://aegis-energy.com/insights/lcfs-rin-pricing-report-through-july-14-2023'

tables = pd.read_html(url)

dfp = tables[3]

dfp = dfp.dropna(axis=0, how='all')

dfp = dfp.T

dfp = dfp.iloc[:, :-2]

dfp = dfp[1:]

dfp = dfp.rename(columns={2: "RIN2", 3 : "Price_Rin"})

dfp['Price_Rin'] = dfp['Price_Rin'].str.replace('$', '')

dfp['Price_Rin'] = dfp['Price_Rin'].astype(float)

df_3 = pd.read_excel(r'H:\\Documents and Settings\\Personal\\HVs_RINs.xlsm')

dfp['RIN2'] = dfp['RIN2'].str.strip()

df_3['RIN2'] = df_3['RIN2'].str.strip()

df_3 = pd.merge(df_3, dfp, on='RIN2', how="inner")

df_3['Price_Rin (b)'] = df_3['Price_Rin'] * 42

df_4 = pd.merge(df, df_3, on='Linkage#', how="inner")

df_5 = pd.merge(df_4, df_2, on='id_counter', how="inner")

df_5['MTM'] = df_5['Quantity'] * df_5['Price_Rin (b)']

# Add a column for month names
credit_pty_trf_index = df_5.columns.get_loc("Credit_pty_trf")
df_5.insert(credit_pty_trf_index + 1, "Month_pty_trf", pd.to_datetime(df_5['Credit_pty_trf']).dt.month)

# Save table to Excel
df_5.to_excel('auto_rin.xlsx', index=False)

# Create the Excel application object
xlApp = win32.Dispatch('Excel.Application')
xlApp.Visible = True

# Open the workbook
wb = xlApp.Workbooks.Open(r'C:\Users\SLANCM\PycharmProjects\python_Project_1\auto_rin.xlsx')
ws_data = wb.Worksheets(1)

# Convert the table to a real Excel table
table_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(df_5.shape[0] + 1, df_5.shape[1]))
table = ws_data.ListObjects.Add(1, table_range, 1, 1)

# Change the style of the table
table.TableStyle = "TableStyleMedium10"

# Format the currency fields in the first sheet
currency_fields = ['purchase_due_amt $', 'sale_due_amt $', 'NET', 'Pay_risk $ (pty -8d)','Price_Rin', 'Price_Rin (b)', 'Unit price$', 'purchase_gbl_amt $', 'sale_gbl_amt $', 'MTM']
for field in currency_fields:
    column_index = df_5.columns.get_loc(field) + 1  # Get the column index (1-based)
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
    field_filters['Year_and_RIN2'] = pt.PivotFields("Year_and_RIN2")
    field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
    field_filters['Linkage_Execution'] = pt.PivotFields("Linkage_Execution")
    field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
    field_filters['Trader'] = pt.PivotFields("Trader")
    field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")

    field_rows = {}
    field_rows['Name_full'] = pt.PivotFields("Name_full")
    field_rows['Deal#'] = pt.PivotFields("Deal#")

    field_values = {}
    field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
    field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
    field_values['NET'] = pt.PivotFields("NET")
    field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

    # Insert filter fields
    field_filters['Linkage_Execution'].Orientation = 3  # xlPageField
    field_filters['Linkage_Execution'].Position = 1

    field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Credit_pty_trf'].Position = 2

    field_filters['Year_and_RIN2'].Orientation = 3  # xlPageField
    field_filters['Year_and_RIN2'].Position = 3

    field_filters['Credit_Term'].Orientation = 3  # xlPageField
    field_filters['Credit_Term'].Position = 4

    field_filters['Trader'].Orientation = 3  # xlPageField
    field_filters['Trader'].Position = 5

    field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Month_pty_trf'].Position = 6

    # Insert row field
    field_rows['Name_full'].Orientation = 1
    field_rows['Name_full'].Position = 1

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
slicer.Height = 140
slicer.Left = 800
slicer.Top = 100

# Create a slicer for "Year_and_RIN2"
slicer_cache = wb.SlicerCaches.Add(pt1, "Year_and_RIN2")
slicer = slicer_cache.Slicers.Add(ws_report1)
slicer.Style = "SlicerStyleLight2"
slicer.Height = 290
slicer.Left = 800
slicer.Top = 245

# Create the second pivot table
ws_report2 = wb.Worksheets.Add()
clear_pts(ws_report2)
pt_cache2 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
pt2 = pt_cache2.CreatePivotTable(ws_report2.Range("B3"), "PivotTable2")

# Insert pivot table field settings for the second pivot table
def insert_pt_field_set2(pt):
    field_filters = {}
    field_filters['Linkage_Execution'] = pt.PivotFields("Linkage_Execution")
    field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
    field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
    field_filters['Trader'] = pt.PivotFields("Trader")
    field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")

    field_rows = {}
    field_rows['Name_full'] = pt.PivotFields("Name_full")
    field_rows['Deal#'] = pt.PivotFields("Deal#")
    field_rows['Year_and_RIN2'] = pt.PivotFields("Year_and_RIN2")

    field_values = {}
    field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
    field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
    field_values['NET'] = pt.PivotFields("NET")
    field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

    # Insert filter fields
    field_filters['Linkage_Execution'].Orientation = 3  # xlPageField
    field_filters['Linkage_Execution'].Position = 1

    field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Credit_pty_trf'].Position = 2

    field_filters['Credit_Term'].Orientation = 3  # xlPageField
    field_filters['Credit_Term'].Position = 3

    field_filters['Trader'].Orientation = 3  # xlPageField
    field_filters['Trader'].Position = 4

    field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Month_pty_trf'].Position = 5

    # Insert row field
    field_rows['Name_full'].Orientation = 1
    field_rows['Name_full'].Position = 1

    field_rows['Deal#'].Orientation = 1
    field_rows['Deal#'].Position = 2

    field_rows['Year_and_RIN2'].Orientation = 1
    field_rows['Year_and_RIN2'].Position = 3

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
slicer.Height = 140
slicer.Left = 800
slicer.Top = 100

# Create a slicer for "Year_and_RIN2"
slicer_cache = wb.SlicerCaches.Add(pt2, "Year_and_RIN2")
slicer = slicer_cache.Slicers.Add(ws_report2)
slicer.Style = "SlicerStyleLight2"
slicer.Height = 290
slicer.Left = 800
slicer.Top = 245

# Create the third pivot table
ws_report3 = wb.Worksheets.Add()
clear_pts(ws_report3)
pt_cache3 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
pt3 = pt_cache3.CreatePivotTable(ws_report3.Range("B3"), "PivotTable2")

# Insert pivot table field settings for the third pivot table
def insert_pt_field_set3(pt):
    field_filters = {}
    field_filters['Linkage_Execution'] = pt.PivotFields("Linkage_Execution")
    field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
    field_filters['entity'] = pt.PivotFields("entity")
    field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
    field_filters['Trader'] = pt.PivotFields("Trader")
    field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")

    field_rows = {}
    field_rows['Name_full'] = pt.PivotFields("Name_full")

    field_values = {}
    field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
    field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
    field_values['NET'] = pt.PivotFields("NET")
    field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

    # Insert filter fields
    field_filters['Linkage_Execution'].Orientation = 3  # xlPageField
    field_filters['Linkage_Execution'].Position = 1

    field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Credit_pty_trf'].Position = 2

    field_filters['entity'].Orientation = 3  # xlPageField
    field_filters['entity'].Position = 3

    field_filters['Credit_Term'].Orientation = 3  # xlPageField
    field_filters['Credit_Term'].Position = 4

    field_filters['Trader'].Orientation = 3  # xlPageField
    field_filters['Trader'].Position = 5

    field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Month_pty_trf'].Position = 6

    # Insert row field
    field_rows['Name_full'].Orientation = 1
    field_rows['Name_full'].Position = 1

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
slicer.Height = 140
slicer.Left = 800
slicer.Top = 100

# Create the fourth pivot table
ws_report4 = wb.Worksheets.Add()
clear_pts(ws_report4)
pt_cache4 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
pt4 = pt_cache4.CreatePivotTable(ws_report4.Range("B3"), "PivotTable2")

# Insert pivot table field settings for the third pivot table
def insert_pt_field_set4(pt):
    field_filters = {}
    field_filters['Linkage_Execution'] = pt.PivotFields("Linkage_Execution")
    field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
    field_filters['entity'] = pt.PivotFields("entity")
    field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
    field_filters['Trader'] = pt.PivotFields("Trader")
    field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")

    field_rows = {}
    field_rows['Name_full'] = pt.PivotFields("Name_full")
    field_rows['RIN2'] = pt.PivotFields("RIN2")

    field_values = {}
    field_values['sale_gbl_amt $'] = pt.PivotFields("sale_gbl_amt $")
    field_values['purchase_gbl_amt $'] = pt.PivotFields("purchase_gbl_amt $")
    field_values['MTM'] = pt.PivotFields("MTM")

    # Insert filter fields
    field_filters['Linkage_Execution'].Orientation = 3  # xlPageField
    field_filters['Linkage_Execution'].Position = 1

    field_filters['Credit_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Credit_pty_trf'].Position = 2

    field_filters['entity'].Orientation = 3  # xlPageField
    field_filters['entity'].Position = 3

    field_filters['Credit_Term'].Orientation = 3  # xlPageField
    field_filters['Credit_Term'].Position = 4

    field_filters['Trader'].Orientation = 3  # xlPageField
    field_filters['Trader'].Position = 5

    field_filters['Month_pty_trf'].Orientation = 3  # xlPageField
    field_filters['Month_pty_trf'].Position = 6

    # Insert row field
    field_rows['RIN2'].Orientation = 1
    field_rows['RIN2'].Position = 1

    field_rows['Name_full'].Orientation = 1
    field_rows['Name_full'].Position = 2

    # Insert values field
    field_values['sale_gbl_amt $'].Orientation = 4
    field_values['sale_gbl_amt $'].Function = -4157  # xlSum
    field_values['sale_gbl_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    field_values['purchase_gbl_amt $'].Orientation = 4
    field_values['purchase_gbl_amt $'].Function = -4157  # xlSum
    field_values['purchase_gbl_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    field_values['MTM'].Orientation = 4
    field_values['MTM'].Function = -4157  # xlSum
    field_values['MTM'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

insert_pt_field_set4(pt4)

# Change pivot table style for the second pivot table
pt4.TableStyle2 = "PivotStyleDark3"

# Adjust column widths for the second pivot table
ws_report4.Columns("D:E").ColumnWidth = 30

# Rename the sheets
ws_data.Name = 'table'
ws_report1.name = 'counterpaties + deal#'
ws_report2.name = 'counterpaties + deal# + RINs'
ws_report3.name = 'counterpaties'
ws_report4.name = 'MTM'

# Save the modified workbook
wb.Save()

print("Tables imported and pivot tables created in 'auto_rin.xlsx'")
