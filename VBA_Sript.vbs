Sub VBA_Homework()
For Each ws In Worksheets

Dim Last_Row As Long
Dim Ticker_Symbol As String
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Stock_Volume As Double
Dim Summary_Table_Row As Long
Dim Open_Start As Long
Dim Open_Price As Double
Dim Close_Price As Double
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease As Double
Dim Last_Row_Percentage_Change As Long
Dim Greatest_Total_Volume As Double

Summary_Table_Row = 2
Total_Stock_Volume = 0
Open_Start = 2
Greatest_Total_Volume = 0
Greatest_Percent_Increase = 0
Greatest_Percent_Decrease = 0

Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To Last_Row

Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker_Symbol = ws.Cells(i, 1).Value

ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
Total_Stock_Volume = 0

Open_Price = ws.Range("C" & Open_Start)
Close_Price = ws.Range("F" & i)
Yearly_Change = Close_Price - Open_Price
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
Else
ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If

If Open_Price = 0 Then
Percentage_Change = 0
Else
Open_Price = ws.Range("C" & Open_Start)
Percentage_Change = Yearly_Change / Open_Price
End If
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
ws.Range("K" & Summary_Table_Row).Value = Percentage_Change

Summary_Table_Row = Summary_Table_Row + 1
Open_Start = i + 1
End If
Next i

Last_Row_Percentage_Change = ws.Cells(Rows.Count, 11).End(xlUp).Row
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

For i = 2 To Last_Row_Percentage_Change
If ws.Range("K" & i).Value > Greatest_Percent_Increase Then
Greatest_Percent_Increase = ws.Range("K" & i).Value
ws.Range("Q2").Value = Greatest_Percent_Increase
ws.Range("P2").Value = ws.Range("I" & i).Value
End If
If ws.Range("K" & i).Value < Greatest_Percent_Decrease Then
Greatest_Percent_Decrease = ws.Range("K" & i).Value
ws.Range("Q3").Value = Greatest_Percent_Decrease
ws.Range("P3").Value = ws.Range("I" & i).Value
End If
If ws.Range("L" & i).Value > Greatest_Total_Volume Then
Greatest_Total_Volume = ws.Range("L" & i).Value
ws.Range("Q4").Value = Greatest_Total_Volume
ws.Range("P4").Value = ws.Range("I" & i).Value
End If
Next i

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Columns("I:Q").AutoFit
Next ws

End Sub
