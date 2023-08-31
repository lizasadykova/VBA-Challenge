Attribute VB_Name = "Module1"
Sub Ticker()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets
  
'new cell names
[I1] = "Ticker"
[J1] = "Yearly Change"
[K1] = "Percent Change"
[L1] = "Total Stock Volume"
[O2] = "Greatest % Increase"
[O3] = "Greatest % Decrease"
[O4] = "Greatest Total Volume"
[P1] = "Ticker"
[Q1] = "Value"

'set up the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
    
'holding tickers
Dim lastDataRow As Long: lastDataRow = Cells(Rows.Count, "A").End(xlUp).Row
Dim summaryTableRowIndex As Integer:
summaryTableRowIndex = 2
Dim currentTicker As String:
currentTicker = Cells(2, 1).Value
Dim open_price As Double:
open_price = Cells(2, 3).Value

'dim close price in column C
Dim close_price As Double
Dim volume As Double:
volume = 0
Dim RowIndex As Long:
RowIndex = 2

'count
For RowIndex = 2 To (lastDataRow + 1)
    
        
'updates to new ticker symbol
If Cells(RowIndex, 1).Value <> currentTicker Then
Cells(summaryTableRowIndex, 9).Value = currentTicker
currentTicker = Cells(RowIndex, 1).Value

'calculating open + close prices
Dim percent_change As Double
Dim year_change As Double
close_price = Cells(RowIndex - 1, 6).Value
year_change = close_price - open_price
Cells(summaryTableRowIndex, 10).Value = year_change
percent_change = (close_price - open_price) / open_price
Cells(summaryTableRowIndex, 11).Value = percent_change
            
open_price = Cells(RowIndex, 3).Value
Cells(summaryTableRowIndex, 12).Value = volume
volume = Cells(RowIndex, 7).Value
summaryTableRowIndex = summaryTableRowIndex + 1
Else
volume = volume + Cells(RowIndex, 7).Value
End If

Next
    
Dim Greatest_Total_Volume As Double: Greatest_Total_Volume = 0
Dim Greatest_Increase As Double: Greatest_Increase = 0
Dim Greatest_Decrease As Double: Greatest_Decrease = 0
    
lastDataRow = Cells(Rows.Count, "L").End(xlUp).Row
For i = 2 To lastDataRow
Dim current_volume As Double
current_volume = Cells(i, 12).Value
If current_volume > Greatest_Total_Volume Then
Greatest_Total_Volume = current_volume
Range("P4").Value = Cells(i, 9).Value
End If

'new tickers
If Cells(i, 11).Value > Greatest_Increase Then
Greatest_Increase = Cells(i, 11).Value
Range("Q2").Value = Cells(i, 11).Value
Range("P2").Value = Cells(i, 9).Value
End If
        
If Cells(i, 11).Value < Greatest_Decrease Then
Greatest_Decrease = Cells(i, 11).Value
Range("Q3").Value = Cells(i, 11).Value
Range("P3").Value = Cells(i, 9).Value
End If
Next
Range("Q4").Value = Greatest_Total_Volume

For i = 2 To lastDataRow

'googled this "Select Case" concept for conditional formatting, really like this'
Select Case Cells(i, 10).Value
Case Is >= 0.01
'green
Cells(i, 10).Interior.ColorIndex = 4
Case Is < 0.01
'red
Cells(i, 10).Interior.ColorIndex = 3

End Select
Next i

    
'autofit columns A - Q
Columns("A:Q").AutoFit
'percent format for column K
Columns("K:K").NumberFormat = "0.00%"
    
Next ws
    
End Sub

