import spreadsheets as s

test_workbook_path = "data/Test1.xlsx"

new_workbook = s.setupNewWorkbook(test_workbook_path)

for sheet in s.getSheetListByFilename(test_workbook_path):
    values = s.getSpreadsheetValues(test_workbook_path, sheet)

    new_values = {"Column 2": ["New", "Values"]}

    new_workbook = s.createSheetWithValues(new_workbook, values, new_values, sheet)


s.saveWorkbook(new_workbook, "data/", "Test2.xlsx")