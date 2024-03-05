import spreadsheets as s

def testMultipleSheets():
    test_workbook_path = "data/Test1.xlsx"

    new_workbook = s.setupNewWorkbook(test_workbook_path)

    for sheet in s.getSheetListByFilename(test_workbook_path):
        values = s.getSpreadsheetValues(test_workbook_path, sheet)

        new_values = {"Column 2": ["New", "Values"]}

        new_workbook = s.createSheetWithValues(new_workbook, values, new_values, sheet)


    s.saveWorkbook(new_workbook, "data/", "Test2.xlsx")


def testRenameColumns():
    test_workbook_path = "data/Test1.xlsx"

    
    newHeaders1 = ["Column 1", "New Column 2"]
    newHeaders2 = ["Column 1", "New Column 2", "New Column 3"]
    newHeaders3 = ["Column A"]

    values = s.getSpreadsheetValues(test_workbook_path)

    new_workbook = s.setupNewWorkbook(test_workbook_path)
    new_sheet1 = s.replaceColumnHeaders(values, newHeaders1)
    new_workbook1 = s.createSheetWithValues(new_workbook, new_sheet1)
    s.saveWorkbook(new_workbook1, "data/", "Output1.xlsx")

    new_workbook = s.setupNewWorkbook(test_workbook_path)
    new_sheet2 = s.replaceColumnHeaders(values, newHeaders2)
    new_workbook2 = s.createSheetWithValues(new_workbook, new_sheet2)
    s.saveWorkbook(new_workbook2, "data/", "Output2.xlsx")

    new_workbook = s.setupNewWorkbook(test_workbook_path)
    new_sheet3 = s.replaceColumnHeaders(values, newHeaders3)
    new_workbook3 = s.createSheetWithValues(new_workbook, new_sheet3)
    s.saveWorkbook(new_workbook3, "data/", "Output3.xlsx")


testRenameColumns()