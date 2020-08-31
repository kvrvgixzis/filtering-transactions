Attribute VB_Name = "Module1"
Option Explicit

Sub filterTransactions():
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ActiveSheet.Name = "backup"
    ActiveSheet.Copy Before:=Worksheets(Sheets.Count)
    ActiveSheet.Name = "result"
    
    Application.UseSystemSeparators = False
    Application.DecimalSeparator = "."
    
    Dim firstRow%, lastRow&, firstCol%, lastCol%, _
        idCol%, typeCol%, sum1Col%, sum2Col%, sum3Col%, j&, i&, _
        curRow As Range, curSum1Cel As Range, curSum2Cel As Range, _
        curSum3Cel As Range, newValue#

    firstRow = 1
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    firstCol = 1
    lastCol = 24
    
    idCol = 18
    typeCol = 14
    
    sum1Col = 11
    sum2Col = 12
    sum3Col = 13
    
    For j = firstRow To lastRow
        Cells(j, sum1Col).NumberFormat = "0.00"
        Cells(j, sum2Col).NumberFormat = "0.00"
        Cells(j, sum3Col).NumberFormat = "0.00"
        
        If Cells(j, typeCol) = "Retail" Then
            Set curRow = Range(Cells(j, firstCol), Cells(j, lastCol))
            Set curSum1Cel = Cells(j, sum1Col)
            Set curSum2Cel = Cells(j, sum2Col)
            Set curSum3Cel = Cells(j, sum3Col)
            
            For i = j + 1 To lastRow
                If Not IsEmpty(Cells(j, firstCol)) _
                And Cells(i, idCol).Value = Cells(j, idCol).Value _
                And Cells(i, typeCol).Value = Cells(j, typeCol).Value Then
                    curRow.Interior.Color = 255
                    
                    newValue = CDbl(Replace(curSum1Cel.Value, ".", ",")) + CDbl(Replace(Cells(i, sum1Col).Value, ".", ","))
                    curSum1Cel.Value = newValue
                    curSum1Cel.NumberFormat = "0.00"
                    
                    newValue = CDbl(Replace(curSum2Cel.Value, ".", ",")) + CDbl(Replace(Cells(i, sum2Col).Value, ".", ","))
                    curSum2Cel.Value = newValue
                    curSum2Cel.NumberFormat = "0.00"
                    
                    newValue = CDbl(Replace(curSum3Cel.Value, ".", ",")) + CDbl(Replace(Cells(i, sum3Col).Value, ".", ","))
                    curSum3Cel.Value = newValue
                    curSum3Cel.NumberFormat = "0.00"
                    
                    Rows(i).EntireRow.Delete
                    i = i - 1
                    lastRow = lastRow - 1
                End If
            Next i
        End If
        Application.StatusBar = "Progress: " & j & "/" & lastRow
        If j Mod 100 = 0 Then DoEvents
    Next j
    
    For j = 1 To Cells(Rows.Count, 1).End(xlUp).Row: Cells(j, 1).Value = j: Next j
    
    Dim wbkExport As Workbook, shtToExport As Worksheet
    Set shtToExport = ThisWorkbook.Worksheets("result")
    Set wbkExport = Application.Workbooks.Add
    
    shtToExport.Copy Before:=wbkExport.Worksheets(wbkExport.Worksheets.Count)
    Application.DisplayAlerts = False
    
    wbkExport.SaveAs Filename:=ThisWorkbook.Path & "\result.csv", FileFormat:=xlCSV, Local:=True
    
    Application.DisplayAlerts = True
    wbkExport.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Application.UseSystemSeparators = True
End Sub
