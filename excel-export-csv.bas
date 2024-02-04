Sub SaveFilteredDataToCSV()
    Dim wsSource As Worksheet
    Dim wsTemp As Worksheet
    Dim wbTemp As Workbook
    Dim strFilePath As String

    ' Set reference to the sheet with filtered data
    Set wsSource = ActiveSheet

    ' Add a new workbook; temporary workbook for CSV
    Set wbTemp = Workbooks.Add
    Set wsTemp = wbTemp.Sheets(1)

    ' Copy visible (filtered) cells to the new workbook
    wsSource.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    wsTemp.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' Prompt user to select where to save the CSV
    strFilePath = Application.GetSaveAsFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Save CSV As")

    ' Check if user canceled the dialog
    If strFilePath = "False" Then Exit Sub

    ' Save the copied data as a CSV file
    wbTemp.SaveAs Filename:=strFilePath, FileFormat:=xlCSV

    ' Close the temporary workbook without saving
    wbTemp.Close SaveChanges:=False
End Sub
