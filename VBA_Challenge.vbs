Sub TickerCheck ()
    'Declare all data types
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Stock_Volume As Double
Dim openPrice As Double
Dim closePrice As Double
Dim SummaryTableRow
Dim wsCount As Integer
    
    'make worksheet active
    'source for how to loop through all worksheets - https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook 
wsCount = ActiveWorkbook.Worksheets.Count

'find last row of spreadsheet
'source for finding last row - https://www.excel-pratique.com/en/vba_tricks/last-row

'set headers using cell ranges 
'change from cell to worksheet range to add to all sheets

For ws = 1 To wsCount
    Worksheets(ws).Range("I1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly_Change"
    Worksheets(ws).Range("K1") = "Percentage_Change"
    Worksheets(ws).Range("L1") = "Total_Stock_Volume"

    'set up variables
    Ticker = ""
    openPrice = 0
    closePrice = 0
    Percentage_Change = 0
    Total_Stock_Volume = 0

    SummaryTableRow = 2
    'set up loop
    For I = 2 To Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
        'figure out ticker symbol and total volume
        Ticker = Worksheets(ws).Cells(I, 1)
        Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7)

        'figure out open price
        If openPrice = 0 Then
        openPrice = Worksheets(ws).Cells(I, 3)
        End If

        'repeat open price for each sheet
        If Ticker <> Worksheets(ws).Cells((I+1),1) Then
        'figure out closing price
        closePrice = worksheets(ws).Cells(I, 6)
        Yearly_Change = closePrice - openPrice

        'show ticker on worksheet
        Worksheets(ws).range("I" & SummaryTableRow).value = Ticker
        'show yearly change
        Worksheets(ws).Range("J" & SummaryTableRow).Value = Yearly_Change
        If Yearly_Change > 0 Then
        Worksheets(ws).Range("J" & SummaryTableRow).interior.ColorIndex = 4
        else
        Worksheets(ws).Range("J" &SummaryTableRow).interior.ColorIndex = 3
        End If

        'show percent change
        If openPrice = 0 Then 
        Percentage_Change = 0
        Else
        Percentage_Change = Yearly_Change / openPrice
        
        End If
        'format percentage_change
        Worksheets(ws).Range("K" & SummaryTableRow).Value = Percentage_Change
        Worksheets(ws).Range("K" & SummaryTableRow).NumberFormat = "0.00%"

        'show stock volume
        Worksheets(ws).Range("L" & SummaryTableRow).Value = Total_Stock_Volume

        'start over
        SummaryTableRow = SummaryTableRow + 1
        Total_Stock_Volume = 0

        End If

    Next I
Next ws

End Sub






