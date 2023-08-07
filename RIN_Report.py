import pandas as pd
import sqlalchemy
import win32com.client as win32

def create_RIN_report():
    server = "xxx"
    port = xxx
    database = "xxx"
    driver = "xxx"
    # driver = "SQL Server"
    connection_string = f"mssql+pyodbc://{server}:{port}/{database}?trusted_connection=yes&driver={driver}"
    engine = sqlalchemy.create_engine(connection_string)

    # Get Data From Med_bi
    sql_query = "SELECT id_counter, Linkage_Execution, entity, Deal#, Deal_date, date_expo, Credit_due_date, Credit_pty_trf, Credit_Term, Linkage#, Trader, [sale_due_amt $], " \
                "[purchase_due_amt $],[sale_due_amt $] - [purchase_due_amt $] AS [NET], [Pay_risk $ (pty -8d)], [Unit price$], [purchase_gbl_amt $], [sale_gbl_amt $], Quantity " \
                "FROM MED_BI.dbo.credit_phy_deal WHERE date_expo = (SELECT MAX(date_expo) FROM MED_BI.dbo.credit_phy_deal)"
    df = pd.read_sql(sql_query, engine)

    sql_query_2 = "SELECT id_counter, Name_full, Group_name_full FROM MED_BI.dbo.COUNTERPARTY"
    df_2 = pd.read_sql(sql_query_2, engine)

    # Importing the RINs price data from online and formatting the table
    url = 'https://aegis-energy.com/insights/lcfs-rin-pricing-report-through-july-28-2023'

    tables = pd.read_html(url)

    dfp = tables[3]

    dfp = dfp.dropna(axis=0, how='all')

    dfp = dfp.T

    dfp = dfp.iloc[:, :-2]

    dfp = dfp[1:]

    dfp = dfp.rename(columns={2: "RIN2", 3: "Price_Rin"})

    dfp['Price_Rin'] = dfp['Price_Rin'].str.replace('$', '')

    dfp['Price_Rin'] = dfp['Price_Rin'].astype(float)

    # Importing the data from the HVs_RINs file and formatting the table
    df_3 = pd.read_excel(r'H:\\Documents and Settings\\Personal\\HVs_RINs.xlsm')

    dfp['RIN2'] = dfp['RIN2'].str.strip()

    df_3['RIN2'] = df_3['RIN2'].str.strip()

    # Merging the tables
    df_3 = pd.merge(df_3, dfp, on='RIN2', how="inner")

    # Create a column for the price of RINs in $/b (1 b = 42 gallons))
    df_3['Price_Rin (b)'] = df_3['Price_Rin'] * 42

    # Merging the tables
    df_4 = pd.merge(df, df_3, on='Linkage#', how="inner")

    df_5 = pd.merge(df_4, df_2, on='id_counter', how="inner")

    # Create a column for the Price_market, MTM_purchase_gbl, MTM_sale_gbl
    df_5['Price_market'] = df_5['Quantity'] * df_5['Price_Rin (b)']

    df_5['MTM_purchase_gbl'] = df_5.apply(
        lambda row: "" if row['purchase_gbl_amt $'] == 0 else row['Price_market'] - row['purchase_gbl_amt $'], axis=1)

    df_5['MTM_sale_gbl'] = df_5.apply(
        lambda row: "" if row['sale_gbl_amt $'] == 0 else row['sale_gbl_amt $'] - row['Price_market'], axis=1)

    # Add a column for month names
    credit_pty_trf_index = df_5.columns.get_loc("Credit_pty_trf")
    df_5.insert(credit_pty_trf_index + 1, "Month_pty_trf", pd.to_datetime(df_5['Credit_pty_trf']).dt.month)

    # Add a column for year names
    credit_pty_trf_index = df_5.columns.get_loc("Credit_pty_trf")
    df_5.insert(credit_pty_trf_index + 1, "Year_pty_trf", pd.to_datetime(df_5['Credit_pty_trf']).dt.year)

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
    currency_fields = ['purchase_due_amt $', 'sale_due_amt $', 'NET', 'Pay_risk $ (pty -8d)', 'Price_Rin',
                       'Price_Rin (b)', 'Unit price$', 'purchase_gbl_amt $', 'sale_gbl_amt $', 'Price_market',
                       'MTM_purchase_gbl', 'MTM_sale_gbl']
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
        field_filters['RIN2'] = pt.PivotFields("RIN2")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['Name_full'] = pt.PivotFields("Name_full")
        field_rows['Deal#'] = pt.PivotFields("Deal#")

        field_values = {}
        field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
        field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
        field_values['NET'] = pt.PivotFields("NET")
        field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

        # Insert filter fields
        field_filters['RIN2'].Orientation = 3  # xlPageField
        field_filters['RIN2'].Position = 1

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

        field_filters['Year_pty_trf'].Orientation = 3  # xlPageField
        field_filters['Year_pty_trf'].Position = 7

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
    slicer.Height = 200
    slicer.Left = 800
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt1, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report1)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 95
    slicer.Left = 800
    slicer.Top = 315

    # Create a slicer for "Year_and_RIN2"
    slicer_cache = wb.SlicerCaches.Add(pt1, "Year_and_RIN2")
    slicer = slicer_cache.Slicers.Add(ws_report1)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 800
    slicer.Top = 415

    # Create the second pivot table
    ws_report2 = wb.Worksheets.Add()
    clear_pts(ws_report2)
    pt_cache2 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt2 = pt_cache2.CreatePivotTable(ws_report2.Range("B3"), "PivotTable2")

    # Insert pivot table field settings for the second pivot table
    def insert_pt_field_set2(pt):
        field_filters = {}
        field_filters['RIN2'] = pt.PivotFields("RIN2")
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

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
        field_filters['RIN2'].Orientation = 3  # xlPageField
        field_filters['RIN2'].Position = 1

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
    slicer.Height = 200
    slicer.Left = 800
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt2, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report2)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 95
    slicer.Left = 800
    slicer.Top = 315

    # Create a slicer for "Year_and_RIN2"
    slicer_cache = wb.SlicerCaches.Add(pt2, "Year_and_RIN2")
    slicer = slicer_cache.Slicers.Add(ws_report2)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 800
    slicer.Top = 415

    # Create the third pivot table
    ws_report3 = wb.Worksheets.Add()
    clear_pts(ws_report3)
    pt_cache3 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt3 = pt_cache3.CreatePivotTable(ws_report3.Range("B3"), "PivotTable2")

    # Insert pivot table field settings for the third pivot table
    def insert_pt_field_set3(pt):
        field_filters = {}
        field_filters['RIN2'] = pt.PivotFields("RIN2")
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['Name_full'] = pt.PivotFields("Name_full")

        field_values = {}
        field_values['sale_due_amt $'] = pt.PivotFields("sale_due_amt $")
        field_values['purchase_due_amt $'] = pt.PivotFields("purchase_due_amt $")
        field_values['NET'] = pt.PivotFields("NET")
        field_values['Pay_risk $ (pty -8d)'] = pt.PivotFields("Pay_risk $ (pty -8d)")

        # Insert filter fields
        field_filters['RIN2'].Orientation = 3  # xlPageField
        field_filters['RIN2'].Position = 1

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
    slicer.Height = 200
    slicer.Left = 800
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt3, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report3)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 95
    slicer.Left = 800
    slicer.Top = 315

    # Create a slicer for "Year_and_RIN2"
    slicer_cache = wb.SlicerCaches.Add(pt3, "Year_and_RIN2")
    slicer = slicer_cache.Slicers.Add(ws_report3)
    slicer.Style = "SlicerStyleLight2"
    slicer.Height = 200
    slicer.Left = 800
    slicer.Top = 415

    # Create the fourth pivot table
    ws_report4 = wb.Worksheets.Add()
    clear_pts(ws_report4)
    pt_cache4 = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt4 = pt_cache4.CreatePivotTable(ws_report4.Range("B3"), "PivotTable2")

    # Insert pivot table field settings for the third pivot table
    def insert_pt_field_set4(pt):
        field_filters = {}
        field_filters['RIN2'] = pt.PivotFields("RIN2")
        field_filters['Credit_pty_trf'] = pt.PivotFields("Credit_pty_trf")
        field_filters['Credit_Term'] = pt.PivotFields("Credit_Term")
        field_filters['Trader'] = pt.PivotFields("Trader")
        field_filters['Month_pty_trf'] = pt.PivotFields("Month_pty_trf")
        field_filters['Year_pty_trf'] = pt.PivotFields("Year_pty_trf")

        field_rows = {}
        field_rows['Name_full'] = pt.PivotFields("Name_full")
        field_rows['Deal#'] = pt.PivotFields("Deal#")

        field_values = {}
        field_values['sale_gbl_amt $'] = pt.PivotFields("sale_gbl_amt $")
        field_values['purchase_gbl_amt $'] = pt.PivotFields("purchase_gbl_amt $")
        field_values['Price_market'] = pt.PivotFields("Price_market")
        field_values['MTM_purchase_gbl'] = pt.PivotFields("MTM_purchase_gbl")
        field_values['MTM_sale_gbl'] = pt.PivotFields("MTM_sale_gbl")

        # Insert filter fields
        field_filters['RIN2'].Orientation = 3  # xlPageField
        field_filters['RIN2'].Position = 1

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
        field_rows['Name_full'].Orientation = 1
        field_rows['Name_full'].Position = 1

        field_rows['Deal#'].Orientation = 1
        field_rows['Deal#'].Position = 2

        # Insert values field
        field_values['sale_gbl_amt $'].Orientation = 4
        field_values['sale_gbl_amt $'].Function = -4157  # xlSum
        field_values['sale_gbl_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['Price_market'].Orientation = 4
        field_values['Price_market'].Function = -4157  # xlSum
        field_values['Price_market'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['purchase_gbl_amt $'].Orientation = 4
        field_values['purchase_gbl_amt $'].Function = -4157  # xlSum
        field_values['purchase_gbl_amt $'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['MTM_purchase_gbl'].Orientation = 4
        field_values['MTM_purchase_gbl'].Function = -4157  # xlSum
        field_values['MTM_purchase_gbl'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

        field_values['MTM_sale_gbl'].Orientation = 4
        field_values['MTM_sale_gbl'].Function = -4157  # xlSum
        field_values['MTM_sale_gbl'].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)"

    insert_pt_field_set4(pt4)

    # Change pivot table style for the second pivot table
    pt4.TableStyle2 = "PivotStyleDark8"

    # Adjust column widths for the second pivot table
    ws_report4.Columns("D:E").ColumnWidth = 30

    # Create a slicer for "Month_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt4, "Month_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report4)
    slicer.Style = "SlicerStyleOther1"
    slicer.Height = 200
    slicer.Left = 900
    slicer.Top = 110

    # Create a slicer for "Year_pty_trf"
    slicer_cache = wb.SlicerCaches.Add(pt4, "Year_pty_trf")
    slicer = slicer_cache.Slicers.Add(ws_report4)
    slicer.Style = "SlicerStyleOther1"
    slicer.Height = 95
    slicer.Left = 900
    slicer.Top = 315

    # Create a slicer for "Year_and_RIN2"
    slicer_cache = wb.SlicerCaches.Add(pt4, "RIN2")
    slicer = slicer_cache.Slicers.Add(ws_report4)
    slicer.Style = "SlicerStyleOther1"
    slicer.Height = 120
    slicer.Left = 900
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

    # Save the modified workbook
    wb.Save()

    print("Tables imported and pivot tables created in 'auto_rin.xlsx'")

if __name__ == '__main__':
    create_RIN_report()
