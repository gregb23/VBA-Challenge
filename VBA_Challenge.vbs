Sub TickerCheck ()
    'Declare all data types
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Stock_Volume As Double
    Dim openPrice As Double
    Dim closePrice As Double   
    Dim SummaryTableRow As Double
    Dim wsCount As Integer
    
    'make worksheet active
    'source for how to loop through all worksheets - https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
    wsCount = ActiveWorkbook.Worksheets.Count

'find last row of spreadsheet
'source for finding last row - https://www.excel-pratique.com/en/vba_tricks/last-row
'lastRowState = ws.Cells(Rows.count, "A").End(xlUp).Row

'set headers using cell ranges 
'change from cell to worksheet range to add to all sheets

For ws = 1 To wsCount
    Worksheets(ws).Range("I1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly_Change"
    Worksheets(ws).Range("K1") = "Percentage_Change"
    Worksheets(ws).Range("L1") = "Total_Stock_Volume"

SummaryTableRow = 2
'set up loop
For I = 2 To Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
    'figure out ticker symbol and total volume
    Ticker = Worksheets(ws).Cells(I, 1)
    Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7)

    'figure out open price
    IF openPrice= 0 Then
        openPrice = Worksheets(ws).Cells(I, 3)
End IF

    Next i
    Next ws

End Sub






