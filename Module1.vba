Attribute VB_Name = "Module1"
Sub AnalyseStocks()

Dim row As Long
Dim RowCount As Long
Dim NextRow As Long
Dim totalStockVolume As Double
Dim openVal As Double
Dim closeVal As Double
Dim yearlyChange As Double
Dim ws As Worksheet
Dim percentChange As Double
 
For Each ws In Sheets


totalStockVolume = 0
NextRow = 2
ws.Range("K1") = "Percent Change"
ws.Range("J1") = "Yearly Change"
ws.Range("I1") = "Ticker"
ws.Range("L1") = "Stock Volume"
ws.Cells(2, 14) = "Increase"
ws.Cells(3, 14) = "Decrease"
ws.Cells(4, 14) = "Volume"

RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
For row = 2 To RowCount

If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
totalStockVolume = 0
openVal = ws.Cells(row, 3).Value
yearlyChange = 0
End If




totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value
If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
closeVal = ws.Cells(row, 6).Value
ws.Cells(NextRow, 9).Value = ws.Cells(row, 1).Value
ws.Cells(NextRow, 12).Value = totalStockVolume
yearlyChange = openVal - closeVal
percentChange = (closeVal - openVal) / openVal
ws.Cells(NextRow, 10).Value = yearlyChange
ws.Cells(NextRow, 11).Value = percentChange
ws.Cells(NextRow, 11).NumberFormat = "0.00%"
If yearlyChange > 0 Then
ws.Cells(NextRow, 10).Interior.Color = RGB(0, 255, 0)
ElseIf yearlyChange < 0 Then
ws.Cells(NextRow, 10).Interior.Color = RGB(255, 0, 0)

End If
NextRow = NextRow + 1
End If

Next row

Dim greatestIncrease
Dim greatestDecrease
Dim highestVolume
Dim increaseTicker
Dim decreaseTicker
Dim volumeTicker


greatestIncrease = ws.Cells(2, 11).Value
increaseTicker = ws.Cells(2, 9).Value
greatestDecrease = ws.Cells(2, 11).Value
decreaseTicker = ws.Cells(2, 9).Value
highestVolume = ws.Cells(2, 12).Value
volumeTicker = ws.Cells(2, 9).Value


For row = 2 To RowCount
If ws.Cells(row, 11).Value > greatestIncrease Then
greatestIncrease = ws.Cells(row, 11).Value
increaseTicker = ws.Cells(row, 9).Value
End If

If ws.Cells(row, 11).Value < greatestDecrease Then
greatestDecrease = ws.Cells(row, 11).Value
decreaseTicker = ws.Cells(row, 9).Value
End If
If ws.Cells(row, 12).Value > highestVolume Then
highestVolume = ws.Cells(row, 12).Value
volumeTicker = ws.Cells(row, 9).Value
End If

Next row

ws.Range("P2") = Format(greatestIncrease, "Percent")
ws.Range("O2") = increaseTicker
ws.Range("P3") = Format(greatestDecrease, "Percent")
ws.Range("O3") = decreaseTicker
ws.Range("O4") = volumeTicker
ws.Range("P4") = highestVolume

Next ws


End Sub


