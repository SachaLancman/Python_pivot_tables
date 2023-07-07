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
sql_query = "SELECT id_counter, Linkage_Execution, entity, Deal#, Deal_date, Credit_due_date, Credit_pty_trf, Credit_Term, Linkage#, Trader, [sale_due_amt $], " \
            "[purchase_due_amt $],[sale_due_amt $] - [purchase_due_amt $] AS [NET], [Pay_risk $ (pty -8d)]" \
            "FROM MED_BI.dbo.credit_phy_deal WHERE Credit_pty_trf BETWEEN '2023-07-01' AND '2023-12-31'"
df_credit_phy_deal = pd.read_sql(sql_query, engine)

sql_query_2 = "SELECT id_counter, Name_full FROM MED_BI.dbo.COUNTERPARTY"
df_Counterparty = pd.read_sql(sql_query_2, engine)

df_HV_RIN = pd.read_excel(r'H:\\Documents and Settings\\Personal\\HVs_RINs.xlsm')

df_credit_RIN = pd.merge(df_credit_phy_deal, df_HV_RIN, on='Linkage#', how="inner")

df_credit_RIN_CTPT = pd.merge(df_credit_RIN, df_Counterparty, on='id_counter', how="inner")

# Save table to Excel
df_credit_RIN_CTPT.to_excel('auto_rin_july.xlsx', index=False)

# Create the Excel application object
xlApp = win32.Dispatch('Excel.Application')
xlApp.Visible = True

# Open the workbook
wb = xlApp.Workbooks.Open(r'C:\Users\SLANCM\PycharmProjects\python_Project_1\auto_rin_july.xlsx')
ws_data = wb.Worksheets(1)

# Convert the table to a real Excel table
table_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(df_credit_RIN_CTPT.shape[0] + 1, df_credit_RIN_CTPT.shape[1]))
table = ws_data.ListObjects.Add(1, table_range, 1, 1)

# Change the style of the table
table.TableStyle = "TableStyleMedium10"

# Format the currency fields in the first sheet
currency_fields = ['purchase_due_amt $', 'sale_due_amt $', 'NET', 'Pay_risk $ (pty -8d)']
for field in currency_fields:
    column_index = df_credit_RIN_CTPT.columns.get_loc(field) + 1  # Get the column index (1-based)
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
    field_filters['Year_and_RIN'] = pt.PivotFields("Year_and_RIN")
    field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
    field_filters['Linkage_Execution'] = pt.PivotFields("Linkage_Execution")
    field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
    field_filters['Trader'] = pt.PivotFields("Trader")

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

    field_filters['Year_and_RIN'].Orientation = 3  # xlPageField
    field_filters['Year_and_RIN'].Position = 3

    field_filters['Credit_Term'].Orientation = 3  # xlPageField
    field_filters['Credit_Term'].Position = 4

    field_filters['Trader'].Orientation = 3  # xlPageField
    field_filters['Trader'].Position = 5

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

# Create a slicer for "Credit_pty_trf"
slicer_cache = wb.SlicerCaches.Add(pt1, "Credit_pty_trf")
slicer = slicer_cache.Slicers.Add(ws_report1)
slicer.Style = "SlicerStyleLight2"
slicer.Height = 240
slicer.Left = 800

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

    field_rows = {}
    field_rows['Name_full'] = pt.PivotFields("Name_full")
    field_rows['Deal#'] = pt.PivotFields("Deal#")
    field_rows['Year_and_RIN'] = pt.PivotFields("Year_and_RIN")

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

    # Insert row field
    field_rows['Name_full'].Orientation = 1
    field_rows['Name_full'].Position = 1

    field_rows['Deal#'].Orientation = 1
    field_rows['Deal#'].Position = 2

    field_rows['Year_and_RIN'].Orientation = 1
    field_rows['Year_and_RIN'].Position = 3

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

# Create a slicer for "Credit_pty_trf"
slicer_cache = wb.SlicerCaches.Add(pt2, "Credit_pty_trf")
slicer = slicer_cache.Slicers.Add(ws_report2)
slicer.Style = "SlicerStyleLight2"
slicer.Height = 240
slicer.Left = 800

# Create the third pivot table
ws_report3 = wb.Worksheets.Add()
clear_pts(ws_report3)
pt_cache3 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
pt3 = pt_cache3.CreatePivotTable(ws_report3.Range("B3"), "PivotTable2")

# Insert pivot table field settings for the second pivot table
def insert_pt_field_set3(pt):
    field_filters = {}
    field_filters['Linkage_Execution'] = pt.PivotFields("Linkage_Execution")
    field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
    field_filters['entity'] = pt.PivotFields("entity")
    field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
    field_filters['Trader'] = pt.PivotFields("Trader")

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

# Create a slicer for "Credit_pty_trf"
slicer_cache = wb.SlicerCaches.Add(pt3, "Credit_pty_trf")
slicer = slicer_cache.Slicers.Add(ws_report3)
slicer.Style = "SlicerStyleLight2"
slicer.Height = 240
slicer.Left = 800

# Rename the sheets
ws_data.Name = 'table'
ws_report1.name = 'counterpaties + deal#'
ws_report2.name = 'counterpaties + deal# + RINs'
ws_report3.name = 'counterpaties'

# Save the modified workbook
wb.Save()

print("Tables imported and pivot tables created in 'auto_rin_july.xlsx'")
