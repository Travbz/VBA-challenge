Attribute VB_Name = "Module1"
Sub StonkScreener()

'Variables list
Dim Ticker As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim i As Double
Dim Last_Row As Double
Dim Summary_Table_Row As Double
Dim Volume As Double
Dim Total_volume As Double
Dim ws As Worksheet

For Each ws In Worksheets
'populate summary headers w titles and format with appropriate symbols
ws.Range("I1").Value = "Ticker"
ws.Range("I1").Font.Bold = True
ws.Range("J1").Value = "Yearly Change"
ws.Range("J1").Font.Bold = True
ws.Columns("J").NumberFormat = "$###,##0.00"
ws.Range("K1").Value = "Percent Change"
ws.Columns("K").NumberFormat = "0.00%"
ws.Range("K1").Font.Bold = True
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("L1").Font.Bold = True
ws.Columns("L").NumberFormat = "$###,###,###,##0.00"

'set variable for last row in ws
Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set initial values to variables
Ticker = ws.Range("A2").Value
Open_Price = ws.Cells(2, 3).Value
Yearly_Change = 0
Percent_Change = 0
Close_Price = 0
Total_volume = 0

'set first output row for summary table, skip the header row and populate starting at row 2
Summary_Table_Row = 2

'loop through rows from 2nd row to last row
For i = 2 To Last_Row

'if statement to check for same ticker in subsequent row
If Ticker = ws.Cells(i + 1, 1).Value Then

'volume value
Volume = ws.Cells(i, 7).Value

'Add each days volume together until a different ticker apperears in next row
Total_volume = Total_volume + Volume

'else statement for if next row ticker is different than prior row
Else

' assign value to volume and close_price from this row
Volume = ws.Cells(i, 7).Value
Close_Price = ws.Cells(i, 6).Value


'calculate yearly change (open price minus close)
Yearly_Change = Close_Price - Open_Price

'check if open_price is zero for division( kept getting a divide by zero error so just set zero value for percent change)
If Open_Price <> 0 Then

'calculate the percent change
Percent_Change = (Yearly_Change / Open_Price)

Else
'value if open price starts with a zero for division error
Percent_Change = 0

End If
'calc total yearly trading volume
Total_volume = Total_volume + Volume

'populate summary table rows
ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
ws.Cells(Summary_Table_Row, 12).Value = Total_volume
ws.Cells(Summary_Table_Row, 9).Value = Ticker
ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change


If Yearly_Change >= 0 Then

'format cells green for gains and red for painz
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

Else

ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3

End If

'next ticker value, new open price and reset volume at 0 for next ticker in rows
Ticker = ws.Cells(i + 1, 1).Value
Open_Price = ws.Cells(i + 1, 3).Value
Close_Price = 0
Total_volume = 0

'new calculations input into next row in summary table
Summary_Table_Row = Summary_Table_Row + 1
End If
Next i
Next ws

End Sub


