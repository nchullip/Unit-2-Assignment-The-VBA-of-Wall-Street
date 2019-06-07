Sub CalculateTotalStockVolume()
'----------------------------------------------------
' This subroutine is used to calculate the Total Stock volume for each Ticker
'
'----------------------------------------------------
    ' Variable Declaration
    Dim stock As String
    Dim volume As Long
    Dim lRow As Long
    Dim temp_stock As String
    Dim Tot_volume As Double
    Dim cntr As Integer
    Dim open_stock As Double
    Dim close_stock As Double
    Dim yearly_change As Double
    Dim perc_change As Double
    Dim Column_Headers() As String
    Column_Headers = Split("Ticker,Yearly Change,Percent Change,Total Stock Volume", ",")
        
    ' Temporary Variable Initialization
    temp_stock = ""
    Tot_volume = 0
    cntr = 2
    a = 9
    
    ' Output Column Formatting
    For k = 0 To 3
        ActiveWorkbook.ActiveSheet.Cells(1, a) = Column_Headers(k)
        ActiveWorkbook.ActiveSheet.Cells(1, a).Font.Bold = True
        a = a + 1
    Next k
    ActiveWorkbook.ActiveSheet.Columns("L").ColumnWidth = 20
    
    ' Find the last row in the sheet
    lRow = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    For i = 2 To Int(lRow)
        ' Get Stock and Volume Details
        stock = ActiveWorkbook.ActiveSheet.Cells(i, 1).Value
        volume = ActiveWorkbook.ActiveSheet.Cells(i, 7).Value
        
        If temp_stock = "" Or temp_stock = stock Then
            If temp_stock = "" Then
                open_stock = ActiveWorkbook.ActiveSheet.Cells(i, 3).Value
            Else
                close_stock = ActiveWorkbook.ActiveSheet.Cells(i, 6).Value
            End If
            Tot_volume = Tot_volume + volume
            temp_stock = stock
            
        ElseIf (temp_stock <> stock) Then
            ' Printing Output
            ActiveWorkbook.ActiveSheet.Cells(cntr, 9) = temp_stock
            ActiveWorkbook.ActiveSheet.Cells(cntr, 12) = Tot_volume
            yearly_change = close_stock - open_stock
            ActiveWorkbook.ActiveSheet.Cells(cntr, 10) = yearly_change
            
            ' Computing Percentage Increase/Decrease only if 'open_stock' field is greater than 0
            If open_stock <> 0 Then
                perc_change = (yearly_change / open_stock)
                ActiveWorkbook.ActiveSheet.Cells(cntr, 11) = perc_change
            
            End If
                
            ' Output Cell Formatting
            ActiveWorkbook.ActiveSheet.Cells(cntr, 11).NumberFormat = "0.00%"
            If yearly_change > 0 Then
                ActiveWorkbook.ActiveSheet.Cells(cntr, 10).Interior.ColorIndex = 4
            Else
                ActiveWorkbook.ActiveSheet.Cells(cntr, 10).Interior.ColorIndex = 3
            End If
            
            'Re-initialize variables
            temp_stock = stock
            open_stock = ActiveWorkbook.ActiveSheet.Cells(i, 3).Value
            Tot_volume = volume
            cntr = cntr + 1
        End If
    Next i
    ' Printing Output for the last entry
    ActiveWorkbook.ActiveSheet.Cells(cntr, 9) = temp_stock
    ActiveWorkbook.ActiveSheet.Cells(cntr, 12) = Tot_volume
    yearly_change = close_stock - open_stock
    ActiveWorkbook.ActiveSheet.Cells(cntr, 10) = yearly_change
    
    ' Computing Percentage Increase/Decrease only if 'open_stock' field is greater than 0
    If open_stock <> 0 Then
        perc_change = (yearly_change / open_stock)
        ActiveWorkbook.ActiveSheet.Cells(cntr, 11) = perc_change
    End If
    
    ' Output Cell Formatting
    ActiveWorkbook.ActiveSheet.Cells(cntr, 11).NumberFormat = "0.00%"
    If yearly_change > 0 Then
        ActiveWorkbook.ActiveSheet.Cells(cntr, 13).Interior.ColorIndex = 4
    Else
        ActiveWorkbook.ActiveSheet.Cells(cntr, 13).Interior.ColorIndex = 3
    End If
End Sub


Sub ComputeMaxAndMin()
' -------------------------------------------
' This subroutine is used to calculate the greatest increase,
' greatest decrease and greatest Total volume
' -------------------------------------------
    
    ' Variable Declaration
    Dim lLastRow As Long
    Dim row As Long
    Dim Column_Headers() As String
    Column_Headers = Split("Greatest % Increase,Greatest % Decrease,Greatest Total Volume", ",")
    a = 2
    
    ' Fetching the last Row
    lLastRow = Range("K" & Rows.Count).End(xlUp).row
    
    ' Cell Formatting
    For i = 0 To 2
        ActiveWorkbook.ActiveSheet.Cells(a, 15) = Column_Headers(i)
        ActiveWorkbook.ActiveSheet.Cells(a, 15).Font.Bold = True
        a = a + 1
    Next i
    ActiveWorkbook.ActiveSheet.Range("P1") = "Ticker"
    ActiveWorkbook.ActiveSheet.Cells(1, 16).Font.Bold = True
    ActiveWorkbook.ActiveSheet.Range("Q1") = "Value"
    ActiveWorkbook.ActiveSheet.Cells(1, 17).Font.Bold = True
    ActiveWorkbook.ActiveSheet.Columns("O").ColumnWidth = 20
    
    ' Computing the Greatest Increase and Corresponding Ticker
    ActiveWorkbook.ActiveSheet.Range("Q2").Formula = "=Max(K1:K" & lLastRow & ")"
    row = WorksheetFunction.Match(ActiveWorkbook.ActiveSheet.Range("Q2").Value, ActiveWorkbook.ActiveSheet.Range("K2:K" & lLastRow), 0)
    ActiveWorkbook.ActiveSheet.Range("P2") = ActiveWorkbook.ActiveSheet.Range("I" & (row + 1)).Value
    ActiveWorkbook.ActiveSheet.Range("Q2").NumberFormat = "0.00%"
    
    ' Computing the Greatest Decrease and Corresponding Ticker
    ActiveWorkbook.ActiveSheet.Range("Q3").Formula = "=Min(K1:K" & lLastRow & ")"
    row = WorksheetFunction.Match(ActiveWorkbook.ActiveSheet.Range("Q3").Value, ActiveWorkbook.ActiveSheet.Range("K2:K" & lLastRow), 0)
    ActiveWorkbook.ActiveSheet.Range("P3") = ActiveWorkbook.ActiveSheet.Range("I" & (row + 1)).Value
    ActiveWorkbook.ActiveSheet.Range("Q3").NumberFormat = "0.00%"
    
    ' Computing the Greatest Total Stock Volume and Corresponding Ticker
    ActiveWorkbook.ActiveSheet.Range("Q4").Formula = "=Max(L1:L" & lLastRow & ")"
    row = WorksheetFunction.Match(ActiveWorkbook.ActiveSheet.Range("Q4").Value, ActiveWorkbook.ActiveSheet.Range("L2:L" & lLastRow), 0)
    ActiveWorkbook.ActiveSheet.Range("P4") = ActiveWorkbook.ActiveSheet.Range("I" & (row + 1)).Value
End Sub


Sub RunAcrossWorkbook()
' -------------------------------------------
' This subroutine is used to call a different subroutine and
' and execute it across all sheets in the Workbook
' -------------------------------------------
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        
        ' Calling subroutine to compute the values
        Call CalculateTotalStockVolume
        
        ' Calling subroutine to fetch the greatest increase/decrease and greatest total volume
        Call ComputeMaxAndMin
    Next
    Application.ScreenUpdating = True
    MsgBox ("Execution Complete!!")
End Sub




