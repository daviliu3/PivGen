import win32com.client as win32
import openpyxl as op
import pywintypes

def getFileName():
    file_name = input(r"Input path: ")
    file_name = file_name.strip('"')
    return file_name

def getParamFromFile(file_name):
    print("Getting parameters from file...\n")
    sheet_name = "PivGenParam"
    workbook = op.load_workbook(filename=file_name, data_only=True)
    data = workbook[sheet_name]
    rows_iter = data.iter_rows()
    param_list = [[cell.value for cell in list(row)] for row in rows_iter]
    del param_list[0]
    for row in param_list:
        sheetName = row[0] + "." + row[3] + " PivGen"
        if row[1] == None:
            row[1] = sheetName
        row[2] = row[2].split(", ")
    return param_list

def invalidSheetName(sheet_name):
    print(f"\nThe sheet name: {sheet_name}")
    print("is invalid, please provide another name below")
    newSheetName = input("New sheet name: ")
    print()
    return newSheetName

def addPivGenSheets(wb, param):
    for sheet in param:
        new_sheet = wb.Worksheets.Add(Before=None, After = wb.Sheets(wb.Sheets.count))
        
        while True:
            try:
                new_sheet.Name = sheet[1]
                print(f"sheet '{sheet[1]}' created")
                break
            except:
                sheet[1] = invalidSheetName(sheet[1])
        

def getWorkBook(file_name):
    xlApp = win32.Dispatch("Excel.Application")
    xlApp.Visible = True
    wb = xlApp.Workbooks.Open(file_name)
    return wb

def insert_pt_field(pt, param):
    row_items = param[2]
    col_item = param[3]
    val_item = param[4]

    # Insert fields to pt
    i = 0
    for item in row_items:
        pt.PivotFields(item).Orientation = 1
        pt.PivotFields(item).Position = i + 1
        i += 1

    pt.PivotFields(col_item).Orientation = 4

    # Insert data fields
    pt.PivotFields(val_item).Orientation = 2

def clear_pts(ws):
    for pt in ws.PivotTables():
        pt.TableRange2.Clear()

def create_pt_designer(wb, param_item):
    pt_name = param_item[1]
    ws_data = wb.Worksheets(param_item[0])
    ws_report = wb.Worksheets(pt_name)
    clear_pts(ws_report)
    pt_cache = wb.PivotCaches().Create(1, ws_data.Range(param_item[-1]).CurrentRegion)
    pt = pt_cache.CreatePivotTable(ws_report.Range("A1"), pt_name)
    return pt
    
def config_pt_designer(pt):
    # Grande total
    pt.ColumnGrand = True
    pt.RowGrand = True

    # Pivot Table Style
    pt.TableStyle2 = "PivotStyleMedium9"

def main():
    # * gets the file name
    file_name = getFileName()
    
    # * get Parameter from file
    param_list = getParamFromFile(file_name)
    
    # * Opens up the Excel file
    print("Calling Excel API")
    wb = getWorkBook(file_name)

    # * Creates the PivGen sheet
    addPivGenSheets(wb, param_list)
    print("All new sheets have been created")
    
    print("\nBegin creating pivot tables")
    for pt_param in param_list:
        # * Create a pivot table base within the sheet
        print("\nInput: {} || Output: {}".format(pt_param[0], pt_param[1]))
        pt = create_pt_designer(wb, pt_param)
        
        # * Configures the designer
        # Picks a pivot table style, etc.
        config_pt_designer(pt)
        
        # * creates the fields for the pivot table
        # and then creates the pivot table.
        insert_pt_field(pt, pt_param)
        print("Pivot Table Created")
        print('-' * 75)


if __name__ == "__main__":
    print("\nBegin PivGen\n")
    main()
    print("\nEnd PivGen")
