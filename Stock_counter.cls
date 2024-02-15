VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_counter()
For Each ws In Worksheets

Dim i As Long
Dim x As Long
Dim total, summary_i As Long
Dim percent_change As Double
Dim great_volume, great_inc, great_dec As Double




'declare column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
 
 
'initially declare variable values
total = 0
summary_i = 2
'declare variable to find total length of column
x = ws.Cells(Rows.Count, 1).End(xlUp).Row

 
 
'First for loop - Loop through all rows
For i = 2 To x
'increment through rows and add cell value to total
total = ws.Cells(i, 7).Value + total
'Conditional statement
If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
'the same
Else

If ws.Cells(i, 3).Value <> 0 Then
   percent_change = ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value)
  'Displays percent values in column, and formats answer as a percentage
   ws.Cells(summary_i, 11).Value = Format(percent_change, "percent")
Else
   ws.Cells(summary_i, 11).Value = Format(0, "percent")
End If

'Displays single values of stock symbol
ws.Cells(summary_i, 9).Value = ws.Cells(i, 1).Value
'Displays yearly change by comparing closing price to opening price
ws.Cells(summary_i, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
'displays the total volume of all stocks
ws.Cells(summary_i, 12).Value = total


'increments through and adds to total volume of each stock
summary_i = summary_i + 1
total = 0
End If

'New conditional statement -> Highlights positive or negative change
If ws.Cells(i, 10).Value < 0 Then
'changes cell to red
ws.Cells(i, 10).Interior.ColorIndex = 3
Else
'changes cell to green
ws.Cells(i, 10).Interior.ColorIndex = 4
End If

Next i


great_volume = ws.Cells(2, 12).Value
great_inc = ws.Cells(2, 11).Value
great_dec = ws.Cells(2, 11).Value


'Second for loop to indicate Greatest incread/decreast, and total volume
For i = 2 To x
'Conditional statement compares all cells in percent change column and finds greatest value
If ws.Cells(i, 11).Value > great_inc Then
    great_inc = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
Else
    great_inc = great_inc
End If

'Conditional statement compares all cells in percent change column and finds smallest value
If ws.Cells(i, 11).Value < great_dec Then
    great_dec = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
Else
    great_dec = great_dec
End If
'Conditional statement compares all total volume cells and finds the maximum
If ws.Cells(i, 12).Value > great_volume Then
    great_volume = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
Else
    great_volume = great_volume
End If
'Format function formats cells to have percentage sign, and scientific value number
ws.Cells(2, 17).Value = Format(great_inc, "Percent")
ws.Cells(3, 17).Value = Format(great_dec, "Percent")
ws.Cells(4, 17).Value = Format(great_volume, "Scientific")

Next i

Next ws
End Sub





