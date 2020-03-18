TSub homework():
For Each ws In Worksheets
Dim Ticker As String
Dim open_price As Double
open_price = ws.Cells(2, 3).Value
Dim close_price As Double
Dim Yearly_Chang As Double
Dim Percent_Chang As Double
Dim Total_Stock_Volume As Double
Dim ticker_summary_row As Integer
ticker_summary_row = 2
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Chang"
ws.Cells(1, 11).Value = "Percent Chang"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatst_Percent_Increase"
ws.Cells(3, 15).Value = "Greatst_Percent_Decrease"
ws.Cells(4, 15).Value = "Greatst_Total_Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Range("i" & ticker_summary_row).Value = ws.Cells(i, 1).Value
ws.Range("l" & ticker_summary_row).Value = Total_Stock_Volume
close_price = ws.Cells(i, 6).Value
yearly_change = close_price - open_price
ws.Range("j" & ticker_summary_row).Value = yearly_change
If open_price = 0 Then
percent_change = 0
Else
percent_change = yearly_change / open_price
End If
ws.Range("k" & ticker_summary_row).Value = percent_change
ws.Range("k" & ticker_summary_row).NumberFormat = "0.0 %"
Total_Stock_Volume = 0
ticker_summary_row = ticker_summary_row + 1
open_price = ws.Cells(i + 1, 3)
End If
Next i
lastrow_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
  For i = 2 To lastrow_summary
  If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
ws.Cells(i, 10).Interior.ColorIndex = 3
End If
Next i
For i = 2 To lastrow_summary
If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:k" & lastrow_summary)) Then
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
ws.Cells(2, 17).NumberFormat = "0.0%"
ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("k2:k" & lastrow_summary)) Then
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
ws.Cells(3, 17).NumberFormat = "0.0%"
ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:k" & lastrow_summary)) Then
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
End If
Next i
Next ws
End Subype your solution here