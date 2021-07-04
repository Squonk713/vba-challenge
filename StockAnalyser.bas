Attribute VB_Name = "StockAnalyser"
Sub StockAnalyser()

'*****Prep Each Worksheet****

Dim ws As Worksheet
Dim Year_Start As Long
Dim Year_End As Long
Year_Start = 20170101
Year_End = 20171230

For Each ws In Sheets

ws.Activate

Dim StockVolume As Double
Dim i As Long
Dim lastRow As Long
Dim Summary_Table_Row As Long
Dim Yearly_Change As Long

Year_Start = Year_Start - 10000
Year_End = Year_End - 10000
StockVolume = 0
Summary_Table_Row = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Range("A1:Q1").EntireColumn.AutoFit

For i = 2 To lastRow
If Cells(i, 1).Value <> Cells(i - 1, 1) Then
Ticker = Cells(i, 1).Value
Range("I" & Summary_Table_Row).Value = Ticker
Summary_Table_Row = Summary_Table_Row + 1
Else
End If
Next i

'*****Sum total volume of each ticker*****

Dim Volume As Double
Volume = 0
Summary_Table_Row = 2
For i = 2 To lastRow
If Cells(i, 1).Value = Cells(i + 1, 1) Then
Volume = Volume + Cells(i, 7).Value
Else
Volume = Volume + Cells(i, 7).Value
Range("L" & Summary_Table_Row).Value = Volume
Summary_Table_Row = Summary_Table_Row + 1
Volume = 0
End If
Next i


Range("L:L").NumberFormat = "0"
Summary_Table_Row = 2

'*****Calculate yearly price change & percentage change*****

Dim Year_Start_Open As Double
Dim Year_End_Close As Double
Dim Change As Double
Dim Percentage_Change As Variant

Year_Start_Open = 0
Year_End_Close = 0

For i = 2 To lastRow

If Cells(i, 1).Value <> Cells(i - 1, 1) Then

Year_Start_Open = Year_Start_Open + Cells(i, 3).Value

ElseIf Cells(i, 1).Value <> Cells(i + 1, 1) Then

Year_End_Close = Year_End_Close + Cells(i, 6).Value
Range("J" & Summary_Table_Row).Value = Year_End_Close - Year_Start_Open
Change = Year_End_Close - Year_Start_Open

Percentage_Change = VBA.IIf(Change = 0, 1, Change) / VBA.IIf(Year_Start_Open = 0, 1, Year_Start_Open)

'I'm sure there's a cleaner way to vix the div0 error, but I've searched four hours and this is the only thing that would work for the pesky PLNT range of zero values

If Percentage_Change = 1 Then
Percentage_Change = 0
Else
End If

Range("K" & Summary_Table_Row).Value = Percentage_Change
Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
Summary_Table_Row = Summary_Table_Row + 1
Year_Start_Open = 0
Year_End_Close = 0

ElseIf Cells(i, 1).Value = Cells(i + 1, 1) Then
End If
Next i

Summary_Table_Row = 2

'*****Calculate greatest % increase*****

Set Rng = Range(Cells(1, 11), Cells(lastRow, 11))
GreatestIncrease = Application.WorksheetFunction.Max(Rng)
Cells(2, 17).Value = GreatestIncrease
Range("Q2").NumberFormat = "0.00%"

For i = 2 To lastRow
If Cells(i, 11).Value <> Cells(2, 17) Then
Else
Ticker = Cells(i, 9).Value
Cells(2, 16).Value = Ticker
End If
Next i

'*****Calculate greatest % decrease*****

Set Rng = Range(Cells(1, 11), Cells(lastRow, 11))
GreatestDecrease = Application.WorksheetFunction.Min(Rng)
Cells(3, 17).Value = GreatestDecrease
Range("Q3").NumberFormat = "0.00%"

For i = 2 To lastRow
If Cells(i, 11).Value <> Cells(3, 17) Then
Else
Ticker = Cells(i, 9).Value
Cells(3, 16).Value = Ticker
End If
Next i

'*****Calculate greatest total volume*****

Set Rng = Range(Cells(1, 12), Cells(lastRow, 12))
GreatestVolume = Application.WorksheetFunction.Max(Rng)
Cells(4, 17).Value = GreatestVolume

For i = 2 To lastRow
If Cells(i, 12).Value <> Cells(4, 17) Then
Else
Ticker = Cells(i, 9).Value
Cells(4, 16).Value = Ticker
End If
Next i

'*****Conditional Formatting for percentage increase*****

Summary_Table_Row = 2
For i = 2 To lastRow
If Range("K" & i).Value > 0 Then
Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
Summary_Table_Row = Summary_Table_Row + 1
ElseIf Range("K" & i).Value < 0 Then
Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
Summary_Table_Row = Summary_Table_Row + 1
ElseIf Range("K" & i).Value = "0" Then
Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
Summary_Table_Row = Summary_Table_Row + 1
ElseIf Range("K" & i).Value = "" Then
Range("K" & Summary_Table_Row).Interior.ColorIndex = 0
Summary_Table_Row = Summary_Table_Row + 1
End If
Next i

Range("A1:Q1").EntireColumn.AutoFit

Next ws

End Sub
