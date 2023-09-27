Sub module_challenge()
MsgBox ("Sequence Started")

' ---------------------------------------------------
' Loop through worksheets
' ---------------------------------------------------
For Each ws In Worksheets


' ---------------------------------------------------
' Variable declarations for loop and conditional sequence
' ---------------------------------------------------
' Declare variable for ticker name
Dim ticker As String

' Declare for yearly change value
Dim yearly_change As Double

' Declare for percent change value
Dim yearly_percent As Double

' Declare for total stock volume
Dim total_volume As LongLong

' Add headers for ticker name, yearly change, percent change, and total volume
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' Declare variable to keep track of the row location for each stock ticker and set to starting row
Dim table_row As Integer
table_row = 2

' Declare variable to hold the opening value of the stock at the beginning of the year and set starting to starting cell
Dim ticker_open As Double
ticker_open = ws.Cells(2, 3).Value

' Declare variable to hold the closing value of the stock at the end of the year
Dim ticker_close As Double

' Declare variables to hold values for greatest increase, decrease, volume, and ticker names
Dim greatest_increase As Double
greatest_increase = 0

Dim increase_ticker As String

Dim greatest_decrease As Double
greatest_decrease = 0

Dim decrease_ticker As String

Dim greatest_volume As LongLong
greatest_volume = 0

Dim volume_ticker As String

' ---------------------------------------------------
' Start a for loop to run through each stock ticker up to the lastrow
' ---------------------------------------------------
Dim LastRow As LongLong
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow


' ---------------------------------------------------
' If a change is detected the following sequence is run:
' ---------------------------------------------------
' If statement to check for change in stock ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Print the stock ticker name
ticker = ws.Cells(i, 1).Value
ws.Cells(table_row, 9).Value = ticker

' Calculate and print the yearly change value
ticker_close = ws.Cells(i, 6).Value
yearly_change = ticker_close - ticker_open
ws.Cells(table_row, 10).Value = yearly_change

' Fill yearly change value with red for negative change and green for positive
If yearly_change >= 0 Then
ws.Cells(table_row, 10).Interior.ColorIndex = 4
Else
ws.Cells(table_row, 10).Interior.ColorIndex = 3
End If

' Calculate and print the percent change
yearly_percent = yearly_change / ticker_open
ws.Cells(table_row, 11).NumberFormat = "0.00%"
ws.Cells(table_row, 11).Value = yearly_percent

' Fill percentage change value with red for negative change and green for positive
If yearly_percent >= 0 Then
ws.Cells(table_row, 11).Interior.ColorIndex = 4
Else
ws.Cells(table_row, 11).Interior.ColorIndex = 3
End If

' Print the total volume
ws.Cells(table_row, 12).NumberFormat = "#,###"
ws.Cells(table_row, 12).Value = total_volume

' Check for greatest increase and store value
If yearly_percent >= greatest_increase Then
greatest_increase = yearly_percent
increase_ticker = ticker
End If

' Check for greatest decrease and store value
If yearly_percent <= greatest_decrease Then
greatest_decrease = yearly_percent
decrease_ticker = ticker
End If

' Check for greatest volume and store value
If total_volume >= greatest_volume Then
greatest_volume = total_volume
volume_ticker = ticker
End If

' Add one to the row location value for the next stock ticker
table_row = table_row + 1

' Reset the stock ticker total volume
total_volume = 0

' Set the opening value for the next stock ticker
ticker_open = ws.Cells(i + 1, 3).Value


' ---------------------------------------------------
' If no change is detected the following sequence is run:
' ---------------------------------------------------
Else

' Add to the sum total volume of the stock ticker
total_volume = total_volume + ws.Cells(i, 7).Value

End If
Next i

' ---------------------------------------------------
' Print values for greatest increase, decrease, and volume
' ---------------------------------------------------

' Print headers
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Print row for greatest increase
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("P2").Value = increase_ticker
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q2").Value = greatest_increase

' Print row for greatest decrease
ws.Range("O3").Value = "Greatest % Increase"
ws.Range("P3").Value = decrease_ticker
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q3").Value = greatest_decrease

' Print row for greatest volume
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P4").Value = volume_ticker
ws.Range("Q4").NumberFormat = "#,###"
ws.Range("Q4").Value = greatest_volume

' Adjust column width to fit values
ws.Columns("A:Q").AutoFit

Next ws

MsgBox ("Sequence Complete")

End Sub