# Visual-Basic-Homework
Homework 2
Sub ModerateStockTicker()

    For Each ws In Worksheets
           
        'Declare variables
        Dim Ticker As String
        Dim LastRow As Long
        Dim TotalVolume As Double
        Dim RowCount As Long
        Dim StockOpenLocation As Long
        Dim StockChange As Double
        Dim StockOpen As Double
        Dim StockClose As Double
        Dim PercentageChange As Double
        
        'Name columns for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        
        ' Set an initial variable for holding the total volume
        TotalVolume = 0
        
        ' Intial location of stock ticker in the summary table
        RowCount = 2
        
        ' Intial location of Stock Opening Amount
        StockOpenLocation = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all stock tickers
        For i = 2 To LastRow
        
            ' Add to the TotalVolume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Check if we are still within the same stock ticker, if not then
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
                    
            ' Print the stock ticker in the summary table
            ws.Range("I" & RowCount).Value = Ticker
            
            ' Print the stock ticker volume to the Summary Table
            ws.Range("L" & RowCount).Value = TotalVolume
            
            'Yearly change from what the stock opened the year at to what the closing price was.
            StockOpen = ws.Range("C" & StockOpenLocation)
            StockClose = ws.Range("F" & i)
            StockChange = StockClose - StockOpen
            
            ' Print the change in the stock price to the Summary Table
            ws.Range("J" & RowCount).Value = StockChange
            
            'Percentage Change
            If StockOpen = 0 Then
               PercentageChange = 0
               
            Else
                StockOpen = ws.Range("C" & StockOpenLocation)
                PercentageChange = StockChange / StockOpen
                
            End If
            
            ' Print the precentage change to the Summary Table
            ws.Range("K" & RowCount).Value = PercentageChange
            ws.Range("K" & RowCount).NumberFormat = "0.00%"
            
            'Conditional formatting that will highlight positive change in green and negative change in red
            If ws.Range("J" & RowCount).Value >= 0 Then
            ws.Range("J" & RowCount).Interior.ColorIndex = 4
            Else
            ws.Range("J" & RowCount).Interior.ColorIndex = 3
            End If
            
            'Add one to the summary table row
            RowCount = RowCount + 1
            
            'Next Opening Stock Location
            StockOpenLocation = i + 1
            
            'Reset Total Volume
            TotalVolume = 0
            
            End If
            
        Next i
        
    Next ws

End Sub
