Attribute VB_Name = "Module1"
Sub worksheet_looper()
    Dim ws As Worksheet
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ' Grabbed the WorksheetName
        Debug.Print ws.Name
        stock_analysis
    Next ws
End Sub

Sub stock_analysis()
    
    ' First compute how many rows are there
    NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    last_row = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'last_col = Range("A1").End(xlToRight).Column
    
    ' Renew the sheets by clearing contents in the selected columns
    Columns("I:Q").ClearContents
    Columns("I:Q").ClearFormats
    
    ' Set column sizes
    Columns("I:L").ColumnWidth = 20
    Columns("O:Q").ColumnWidth = 20
    
    ' Declaring essential variables
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim sum_range As String
    Dim tickerCount As Integer
    
    ' Hardcoding output columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Generating the ticker summary results
    tickerCount = 0
    
    For i = 2 To NumRows
        curr_ticker = Cells(i, 1)
        tickerCount = tickerCount + 1
        ' Identifying the end location of a ticker
        r = i
        While Cells(r, 1).Value = curr_ticker
            r = r + 1
        Wend
        yearly_change = Cells(r - 1, 6).Value - Cells(i, 3).Value
        percent_change = yearly_change / Cells(i, 3).Value * 100
        percent_change = Round(percent_change, 2)
        
        Range("I" & CStr(tickerCount + 1)).Value = curr_ticker
        
        ' Assigning yearly_change values to the column and doing the conditional formatting
        Range("J" & CStr(tickerCount + 1)).Value = yearly_change
        Range("J" & CStr(tickerCount + 1)).NumberFormat = "0.00"
        If yearly_change < 0 Then
            Range("J" & CStr(tickerCount + 1)).Interior.ColorIndex = 3
        ElseIf yearly_change >= 0 Then
            Range("J" & CStr(tickerCount + 1)).Interior.ColorIndex = 4
        End If
        ' Assigning percent_change values to the column and doing the conditional formatting
        Range("K" & CStr(tickerCount + 1)).Value = CStr(percent_change) & "%"
        If percent_change < 0 Then
            Range("K" & CStr(tickerCount + 1)).Interior.ColorIndex = 3
        ElseIf percent_change >= 0 Then
            Range("K" & CStr(tickerCount + 1)).Interior.ColorIndex = 4
        End If
        ' Gathering and computing the sum of volumes of each ticker group
        sum_range = "G" & CStr(i) & ":" & "G" & CStr(r - 1)
        Range("L" & CStr(tickerCount + 1)).Value = Application.WorksheetFunction.Sum(Range(sum_range))
        i = r - 1
    Next
    
    ' Computing the extreme results and their indices
    mx_vol = Application.WorksheetFunction.Max(Range("L2", Range("L2").End(xlDown)))
    mx_vol_RowIndex = Application.WorksheetFunction.Match(mx_vol, Range("L1" & ":" & "L" & CStr(last_row)), 0)
    
    mx_inc = Application.WorksheetFunction.Max(Range("K1", Range("K2").End(xlDown)))
    mx_inc_RowIndex = Application.WorksheetFunction.Match(mx_inc, Range("K1" & ":" & "K" & CStr(last_row)), 0)
    
    mx_dec = Application.WorksheetFunction.Min(Range("K1", Range("K1").End(xlDown)))
    mx_dec_RowIndex = Application.WorksheetFunction.Match(mx_dec, Range("K1" & ":" & "K" & CStr(last_row)), 0)
        
    ' Hardcoding output rows columns for the extreme results
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Assigning the extreme results values
    Range("Q2").Value = CStr(mx_inc * 100) & "%"
    Range("Q3").Value = CStr(mx_dec * 100) & "%"
    Range("Q4").Value = mx_vol
    Range("Q4").NumberFormat = "0.00E+00"
    
    
    Range("P2").Value = Range("I" & CStr(mx_inc_RowIndex)).Value
    Range("P3").Value = Range("I" & CStr(mx_dec_RowIndex)).Value
    Range("P4").Value = Range("I" & CStr(mx_vol_RowIndex)).Value
    
    'MsgBox ("Done")
End Sub
