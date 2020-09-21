# VBA-challenge

My code was in the Excel document

I have pasted it for reference below:

Sub stocks() 'Loop through each worksheet

Dim WS As Worksheet For Each WS In ActiveWorkbook.Worksheets WS.Activate

'Declare new variables

Dim ticker As String Dim yearly_change As Double Dim percent_change As Double Dim open_price As Double Dim close_price As Double Dim volume As Double volume = 0 Dim row As Integer row = 2

'Insert Column Haders Cells(1, 9).Value = "Ticker" Cells(1, 10).Value = "Yearly Change" Cells(1, 11).Value = "Percent Change" Cells(1, 12).Value = "Total Stock Volume"

'search last row that has text and count it

LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row

'set open_price to beg year otherwise loop it

open_price = Cells(row, 3).Value

'Start loop through ticker symbols For i = 2 To LastRow

'Check if we are still within same ticker and assign ticker symbol

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'Set the ticker symbol

ticker = Cells(i, 1).Value Cells(row, 9).Value = ticker

'closing price close_price = Cells(i, 6).Value

'yearly change yearly_change = close_price - open_price Cells(row, 10) = yearly_change 'percent change

If (open_price = 0 And close_price = 0) Then percent_change = 0

ElseIf (open_price = 0 And close_price <> 0) Then percent_change = 1

Else percent_change = yearly_change / open_price

Cells(row, 11).Value = percent_change Cells(row, 11).NumberFormat = "0.00%"

End If

'Calculate volume adding the volume for all the same tickers

volume = volume + Cells(i, 7).Value Cells(row, 12).Value = volume

'Set the new price for the ticker

open_price = Cells(i + 1, 3).Value

' Add new row to summary table row = row + 1

'Set the volume back to zero if there is a new ticker symbol

volume = 0

'otherwise, if it's till the same ticker keep summing volume values Else volume = volume + Cells(i, 7).Value

End If

Next i 'cell color LastRow_yearly_change = Cells(Rows.Count, 10).End(xlUp).row

For j = 2 To LastRow_yearly_change

If Cells(j, 10).Value > 0 Then Cells(j, 10).Interior.ColorIndex = 4

Else

Cells(j, 10).Interior.ColorIndex = 3

End If

Next j

'Headers for the new variable

Cells(2, 14).Value = "greatest % increase" Cells(3, 14).Value = "greatest % decrease" Cells(4, 14).Value = "greatest % volume" Cells(1, 15).Value = "ticker" Cells(1, 16).Value = "value"

For k = 2 To LastRow_yearly_change

'> % increase

If Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & LastRow_yearly_change)) Then 'ticker symbol Cells(2, 15).Value = Cells(k, 9).Value Cells(2, 16).Value = Cells(k, 11).Value Cells(2, 16).NumberFormat = "0.00%"

ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & LastRow_yearly_change)) Then 'ticker symbol Cells(3, 15).Value = Cells(k, 9).Value Cells(3, 16).Value = Cells(k, 11).Value Cells(3, 16).NumberFormat = "0.00%"

ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & LastRow_yearly_change)) Then 'ticker symbol Cells(4, 15).Value = Cells(k, 9).Value Cells(4, 16).Value = Cells(k, 12).Value

End If

Next k

Next WS

End Sub
