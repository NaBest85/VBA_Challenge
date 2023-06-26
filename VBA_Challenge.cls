VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()

Dim ws As Worksheet
Dim Row As Double
Dim curRow As Double
Dim firstRow As Double
Dim lastRow As Double
Dim yearOpen As Double
Dim yearClose As Double
Dim volume As Double
    
    
For Each ws In Sheets
Worksheets(ws.Name).Activate
Row = 2
curRow = 2
firstRow = 2
lastRow = Range("A2").End(xlDown).Row
volume = 0
        
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
        
For i = firstRow To lastRow
tickers = Cells(i, 1).Value
tickers2 = Cells(i - 1, 1).Value
If tickers <> tickers2 Then
Cells(curRow, 9).Value = tickers
curRow = curRow + 1
End If
Next i

For i = firstRow To lastRow + 1
tickers = Cells(i, 1).Value
tickers2 = Cells(i - 1, 1).Value
If tickers = tickers2 And i > 2 Then
volume = volume + Cells(i, 7).Value
ElseIf i > 2 Then
Cells(Row, 12).Value = volume
Row = Row + 1
volume = 0
Else
volume = volume + Cells(i, 7).Value
End If
Next i
        
Row = 2
For i = firstRow To lastRow
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
yearClose = Cells(i, 6).Value
ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
yearOpen = Cells(i, 3).Value
End If
If yearOpen > 0 And yearClose > 0 Then
increase = yearClose - yearOpen
percent_increase = increase / yearOpen
Cells(Row, 10).Value = increase
Cells(Row, 11).Value = FormatPercent(percent_increase)
yearClose = 0
yearOpen = 0
Row = Row + 1
End If
Next i
        
max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))

Range("Q2").Value = FormatPercent(max_per)
Range("Q3").Value = FormatPercent(min_per)
Range("Q4").Value = max_vol
        
For i = firstRow To lastRow
If max_per = Cells(i, 11).Value Then
Range("P2").Value = Cells(i, 9).Value
ElseIf min_per = Cells(i, 11).Value Then
Range("P3").Value = Cells(i, 9).Value
ElseIf max_vol = Cells(i, 12).Value Then
Range("P4").Value = Cells(i, 9).Value
End If
Next i
For i = firstRow To lastRow
If IsEmpty(Cells(i, 10).Value) Then Exit For
If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
Else
Cells(i, 10).Interior.ColorIndex = 3
End If
Next i
Next ws
End Sub
