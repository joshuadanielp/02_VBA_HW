Sub stock_ticker()

    'Need to assign a nested For loop to work through the worksheets
  For Each ws In Worksheets
  
    ' Set an initial variable for holding the ticker
    Dim ticker As String

    ' Set an initial variable for holding the total stock volume per ticker
    Dim stock_vol As Double
    stock_vol = 0

    ' Keep track of the location for each ticker in the ticker column
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    LastRowNumber = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Label column I as "Ticker" and column J as "Total Stock Volume"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"

    ' Loop through all tickers
    For i = 2 To LastRowNumber
    
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker
            ticker = ws.Cells(i, 1).Value

            ' Add to the stock vol
            stock_vol = stock_vol + ws.Cells(i, 3).Value

            ' Print the ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = ticker
            
            ' Print the stock volume to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = stock_vol

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the stock vol
            stock_vol = 0

        ' If the cell immediately following a row is the same brand...
        Else

            ' Add to the Brand Total
            stock_vol = stock_vol + ws.Cells(i, 7).Value

        End If

    Next i

  Next ws

End Sub

