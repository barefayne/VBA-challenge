Attribute VB_Name = "TickerStat_F"
Sub TickerStat()

For Each ws In Worksheets
 
 Dim i As Long
 Dim Ticker As Long
 Dim TCount As Long
 Dim LastRow As Long
 Dim VTotal As Long
 Dim YChange As Double
 Dim PChange As Double
 
 
'Set column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Set ticker counter to start from row 2
TCount = 2
'Set ticker to start from row 2
Ticker = 2

'Find last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all tickers
For i = 2 To LastRow

' Check if we are still within the same ticker type, if it is not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
' Print ticker per group in col
ws.Cells(TCount, 9).Value = ws.Cells(i, 1).Value

'Calculate Yearly Change and print in col J
ws.Cells(TCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(Ticker, 3).Value

'Format/Color Yearly Change
If ws.Cells(TCount, 10).Value < 0 Then
ws.Cells(TCount, 10).Interior.ColorIndex = 3
Else
ws.Cells(TCount, 10).Interior.ColorIndex = 4
End If
 
'Calculate Percent Change and print in col K
PChange = ((ws.Cells(i, 6).Value - ws.Cells(Ticker, 3).Value) / ws.Cells(Ticker, 3).Value)

'Format percent change
ws.Cells(TCount, 11).Value = Format(PChange, "Percent")

'Sum volumne and print in col L
ws.Cells(TCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(Ticker, 7), ws.Cells(i, 7)))
 
'Increase ticker counter by 1
TCount = TCount + 1

'Set row start
Ticker = i + 1

End If

Next i

'Adjust column width
ws.Columns("A:L").AutoFit

Next ws


MsgBox ("Fixes Complete")

End Sub

        


