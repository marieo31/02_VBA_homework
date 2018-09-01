Sub LoopOnWk()
'''
' Loop on all the Worksheets
''''

Dim ws As Worksheet

'Set ws = ThisWorkbook.Worksheets(1)

For Each ws In ThisWorkbook.Worksheets
    Call TotalStockVolModerate(ws)
    Call FindingExtremeValue(ws)
Next ws


End Sub


Sub TotalStockVolModerate(ws)
'''
' Date: 20180825
' Description:
' >> Easy: Create a script that will loop through each year of stock data and grab the total
' amount of volume each stock had over the year.
' >> Moderate:
' Create a script that will loop through all the stocks and take the following info.
' -> Yearly change from what the stock opened the year at to what the closing price was.
' -> The percent change from the what it opened the year at to what it closed.
' -> The total Volume of the stock
' -> ticker symbol
' You should also have conditional formatting that will highlight positive change in green and negative change in red.
'''

' Declarations
'--------------
Dim total_volume As Double 'Total volume for each brand
Dim cl_price As Double  'Closing price
Dim op_price As Double  'Opening price
Dim yr_change As Double ' Year change
Dim pc_change As Double ' % change

'Looking for the last row of the table
last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

' Writing labels
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Looping around
' Initialisations
current_ticker = ws.Cells(2, 1).Value
op_price = ws.Cells(2, 3).Value
total_volume = ws.Cells(2, 7).Value
ticker_nb = 1

For ii = 2 To last_row
    
    Next_ticker = ws.Cells(ii + 1, 1).Value

    'If the following tickers are identical
    If StrComp(Next_ticker, current_ticker) = 0 Then 'NB: StrComp = 0 when strings are similar and 1 when they are not
        ' Adding up the total volume
        total_volume = total_volume + ws.Cells(ii + 1, 7).Value
        
    
    Else  'If they are different

        ' Geting the closing price
        cl_price = ws.Cells(ii, 6).Value
        
        ' Calculating the year change
        yr_change = cl_price - op_price
        
        ' Calculating the % of change
        If op_price <> 0 Then
            pc_change = (cl_price - op_price) / op_price 'pc of change
        Else
            pc_change = 0 'is there a NaN value in VBA ?
        End If
        
        'Writing the values in the summary table
        ws.Range("I" & ticker_nb + 1).Value = current_ticker
        ws.Range("J" & ticker_nb + 1).Value = yr_change
        ws.Range("K" & ticker_nb + 1).Value = pc_change
        ws.Range("L" & ticker_nb + 1).Value = total_volume
        
        'Setting the % style for the pc_change
        ws.Range("K" & ticker_nb + 1).Style = "Percent"
        
        ' Conditional formating of the yr_change cells
        If yr_change > 0 Then
            ws.Cells(ticker_nb + 1, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf yr_change < 0 Then
            ws.Cells(ticker_nb + 1, 10).Interior.Color = RGB(255, 0, 0)
        End If
        
        'Incrementing the number of tickers
        ticker_nb = ticker_nb + 1
        
        'Reset the values for the next ticker
        total_volume = ws.Cells(ii + 1, 7).Value
        op_price = ws.Cells(ii + 1, 3).Value
        current_ticker = Next_ticker
        
    End If
    
Next ii


End Sub

Sub FindingExtremeValue(ws)
'''
' Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
'''

' Looking for the last row of the summary table
last_row = ws.Cells(Rows.Count, "K").End(xlUp).Row

' Extracting the pc_change and total Stock Volume tables
'-------------------------------------------------------
pc_change_table = ws.Range("K2:K" & last_row)
tot_StVol_table = ws.Range("L2:L" & last_row)

'Greatest pc increase (Value and Index)
    Maxpc = WorksheetFunction.Max(pc_change_table)
    MaxpcInd = WorksheetFunction.Match(Maxpc, pc_change_table, 0) + 1
'Greatest pc decrease (Value and Index)
    Minpc = WorksheetFunction.Min(pc_change_table)
    MinpcInd = WorksheetFunction.Match(Minpc, pc_change_table, 0) + 1
'Greatest volume (Value and Index)
    MaxVol = WorksheetFunction.Max(tot_StVol_table)
    MaxVolInd = WorksheetFunction.Match(MaxVol, tot_StVol_table, 0) + 1

'Writing down in the table
'-------------------------
'Labels
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
' Tickers
    ws.Range("O2").Value = ws.Range("I" & MaxpcInd).Value
    ws.Range("O3").Value = ws.Range("I" & MinpcInd).Value
    ws.Range("O4").Value = ws.Range("I" & MaxVolInd).Value
' Max, Min values
    ws.Range("P2").Value = Maxpc
    ws.Range("P3").Value = Minpc
    ws.Range("P4").Value = MaxVol
    
ws.Range("P2:P3").Style = "Percent"

End Sub
