Attribute VB_Name = "Module1"
Sub stock_analysis()
Dim i As Long
Dim j As Long
Dim LastRow As Long
Dim ticker As String
Dim volume_total As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
Dim opening_price As Double
Dim closing_price As Double
Dim greatest_vol_total As Double
Dim percent_increase As Double
Dim percent_decrease As Double
Dim greatest_vol_ticker As String
Dim percent_increase_ticker As String
Dim percent_decrease_ticker As String
Dim x As Long

For Each ws In Worksheets
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume:"

volume_total = 0
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2

For i = 2 To CLng(LastRow)
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        volume_total = volume_total + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("L" & Summary_Table_Row).Value = volume_total
        volume_total = 0
        closing_price = ws.Cells(i, 6).Value

            yearly_change = closing_price - opening_price
            percent_change = (closing_price - opening_price) / opening_price

        
        ws.Range("J" & Summary_Table_Row).Value = yearly_change
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        ws.Range("K" & Summary_Table_Row).Style = "Percent"
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Summary_Table_Row = Summary_Table_Row + 1
        
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        opening_price = ws.Cells(i, 3).Value
          
    Else
        volume_total = volume_total + ws.Cells(i, 7).Value
        
    End If
    
    Next i
For j = 2 To LastRow
    If ws.Range("J" & j).Value > 0 Then
        ws.Range("J" & j).Interior.ColorIndex = 4
    Else
        ws.Range("J" & j).Interior.ColorIndex = 3
    End If
    
    Next j
    
    
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

For x = 2 To LastRow
    If ws.Range("L" & x).Value > ws.Range("Q4").Value Then
        greatest_vol_total = ws.Range("L" & x).Value
        greatest_vol_ticker = Range("I" & x).Value
        ws.Range("P4").Value = greatest_vol_ticker
        ws.Range("Q4").Value = greatest_vol_total
        ws.Range("Q4").NumberFormat = "0"
    End If
    If ws.Range("K" & x).Value > ws.Range("Q2").Value Then
        percent_increase = ws.Range("K" & x).Value
        percent_increase_ticker = ws.Range("I" & x).Value
        ws.Range("P2").Value = percent_increase_ticker
        ws.Range("Q2").Value = percent_increase
    End If
    If ws.Range("K" & x).Value < ws.Range("Q3").Value Then
        percent_decrease = ws.Range("K" & x).Value
        percent_decrease_ticker = ws.Range("I" & x).Value
        ws.Range("P3").Value = percent_decrease_ticker
        ws.Range("Q3").Value = percent_decrease
        ws.Range("Q2:Q3").Style = "Percent"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    End If
    Next x
    
ws.Columns.AutoFit

Next ws

End Sub

