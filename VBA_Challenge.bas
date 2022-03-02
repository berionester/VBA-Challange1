Attribute VB_Name = "Module1"
Sub Stock_Analysis3()
Dim ws As Worksheet

For Each ws In Worksheets
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Columns("I:P").EntireColumn.ClearContents
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To last_row
        ticker_name = Range("A" & i)
        opening_price = Range("C" & i)
        ticker_count = Application.WorksheetFunction.CountIf(Columns("A"), ticker_name)
        'i=2    ticker_count=261
        last_ticker_row_number = i + ticker_count - 1
        closing_price = Range("F" & last_ticker_row_number)
        Yearly_Change = closing_price - opening_price
        If opening_price = 0 Then
            Percent_change = 0
        Else
            Percent_change = closing_price / opening_price - 1
        End If
        total_stock_volume = Application.WorksheetFunction.Sum(Range("G" & i & ":G" & last_ticker_row_number))
        summary_table_last_row = Cells(Rows.Count, 9).End(xlUp).Row + 1
        ws.Range("I" & summary_table_last_row) = ticker_name
        ws.Range("J" & summary_table_last_row) = Yearly_Change
        ws.Range("K" & summary_table_last_row) = Percent_change
        ws.Range("L" & summary_table_last_row) = total_stock_volume
        i = last_ticker_row_number
     
    Next
    
    ws.Range("N2") = "Greatest%Increase"
    ws.Range("N3") = "Greatest%Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-4],0))"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-4],0))"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-3],0))"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-5])"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-5])"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-4])"
    Range("P5").Select
    ws.Columns("A:P").EntireColumn.AutoFit

    
    
    
    
Next



End Sub


