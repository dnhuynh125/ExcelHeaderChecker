import wx
import xlrd
import xlsxwriter


def GetNumTests(filePath):
    """
        returns number of tests based on length of row splice from excel file

        :params:
            filePath: string value of entire file path

        :returns:
            num: integer value of number of test columns
    """
    
    prodWorkbook = xlrd.open_workbook(prodFilePath)
    prodWorksheet = prodWorkbook.sheet_by_index(0)                                
    num = len(prodWorksheet.row_values(8, start_colx = 2, end_colx = None))
    
    return num


def GetColLetter(number):
    """
        converts column integer value into character value and returns as a string

        :params:
            number: integer value of column coordinate in excel file starting at 0 as column A
            
        :returns:
            string: string value of column lettering
    """
    
    string = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        string = chr(65 + remainder) + string                   # adds character by character in reverse order to get column letter
        
    return string


def WriteHeaders(finalWorksheet, filePath, row, col, rowRef):
    """
        writes 6 row splices of header information to finalOutput worksheet then
        returns updated row & col values for future writing location

        :params:
            finalWorksheet: object variable of an excel worksheet
            filePath: string value of entire file path
            row: integer value of current row location for writing
            col: integer value of current col location for writing
            rowRef: integer value of row referenced from reading location which is filePath
            
        :returns:
            row: integer value of updated row coordinate
            col: integer value of updated col coordinate
    """
    
    workbook = xlrd.open_workbook(filePath)                                
    worksheet = workbook.sheet_by_index(0)                                       

    for i in range(6):
        for header in worksheet.row_values(rowRef, start_colx = 2, end_colx = None):
            finalWorksheet.write(row, col, header)                                          
            col += 1
        row += 1
        col = 0
        rowRef += 1

    return row, col


def ReceiveFile(prompt):
    """
        returns a single excel file path by using file dialog box

        :params:
            prompt: string value to display in title bar of file dialog
            
        :returns:
            filePath: string value of entire file path
    """
    
    openFileDialog = wx.FileDialog(None, prompt, "", "", "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx")
    openFileDialog.ShowModal()                                                 
    filePath = openFileDialog.GetPath()                                        
    openFileDialog.Destroy()
    
    return filePath                                                            


if __name__ == "__main__":
    app = wx.App()

    sampFilePath = ReceiveFile("Open Sample File")
    prodFilePath = ReceiveFile("Open Production File")
    
    if sampFilePath and prodFilePath:
        prodFilePathSplit = prodFilePath.replace('.xls', '').replace('.xlsx', '')       # split file path to remove extension
        fileName = prodFilePathSplit                                                    # save first element as fileName
        workbook = xlsxwriter.Workbook(fileName + '_CheckHeaders.xlsx')                 # create header check workbook with saved file name
        finalWorksheet = workbook.add_worksheet()
        failFormat = workbook.add_format({'bg_color': '#FFC7CE',
                                          'font_color': '#9C0006'})                     # light red fill with dark red text
        
        # initial writing location of header check workbook
        row = 0         
        col = 0
        
        # initial reading location of samp workbook
        rowSamp = 8
        # write headers from sample file to final worksheet
        row, col = WriteHeaders(finalWorksheet, sampFilePath, row, col, rowSamp)        
        row += 3
        
        # initial reading location of prod workbook
        rowProd = 8
        # write headers from production file to final worksheet
        row, col = WriteHeaders(finalWorksheet, prodFilePath, row, col, rowProd)
        row += 3
        
        # initial reading location of samp and prod workbooks
        rowSamp = 1
        rowProd = 10
        numTests = GetNumTests(prodFilePath)
        for i in range(6):
            for header in range(numTests):
                colLetter = GetColLetter(col + 1)                                               # get column letter of current column
                formula = '=' + colLetter + str(rowSamp) + '=' + colLetter + str(rowProd)       # write comparison formula
                finalWorksheet.write(row, col, formula)                                         # write formula to cell
                col += 1
            row += 1
            rowSamp += 1
            rowProd += 1
            col = 0

        # column letter of final column
        colLetter = GetColLetter(numTests)
        # apply conditional format for entire range of header comparisons
        finalWorksheet.conditional_format('A19:' + colLetter + '24', {'type': 'text',
                                                                 'criteria': 'containing',
                                                                 'value': 'FALSE',
                                                                 'format': failFormat})
        workbook.close()        
    else:
        if not sampFilePath and prodFilePath:
            wx.MessageBox("You did not select a sample file. Quitting program...", "ERROR", wx.OK|wx.ICON_ERROR)
        elif sampFilePath and not prodFilePath:
            wx.MessageBox("You did not select a production file. Quitting program...", "ERROR", wx.OK|wx.ICON_ERROR)
        else:
            wx.MessageBox("You did not select any sample or lot files. Quitting program...", "ERROR", wx.OK|wx.ICON_ERROR)
            
    app.MainLoop()




