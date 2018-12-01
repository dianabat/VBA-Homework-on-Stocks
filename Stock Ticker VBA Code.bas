Attribute VB_Name = "Module1"
Sub Stocks()

For Each ws In Worksheets

    'Set last row variable to the last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Create a total volume variable and set to first value
    Dim TotalVolume As Double
    TotalVolume = ws.Cells(2, 7).Value

    'Create a stock ticker symbol variable and set to first value
    Dim StockTicker As String
    StockTicker = ws.Cells(2, 1).Value
    
    'Stocks opening price stored in variable set to first value
    Dim OpenPrice As Double
    OpenPrice = ws.Cells(2, 3).Value
    
    'Stocks closing price, yearly change, percent change
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    'Variable to iterate through the rows to print the summary
    Dim PrintRow As Integer
    PrintRow = 2
    
    ws.Cells(1, 9).Value = "Stock Ticker"
    ws.Cells(1, 10).Value = "Total Volume"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"

    For i = 2 To LastRow
    
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            'If the next row is the same stock as the current then add the volume to the total
            TotalVolume = TotalVolume + ws.Cells(i + 1, 7).Value
            
        Else
            'Otherwise, print the current total volume and stock ticker in the summary columns
            ws.Cells(PrintRow, 9).Value = StockTicker
            ws.Cells(PrintRow, 10).Value = TotalVolume
            
            'Capture the close price for the stock and calculate yearly and percent change
            ClosePrice = ws.Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice
            PercentChange = YearlyChange / OpenPrice
            
            'Print yearly and percent change
            ws.Cells(PrintRow, 11).Value = YearlyChange
            ws.Cells(PrintRow, 12).Value = PercentChange
            
            'Apply color
            If YearlyChange < 0 Then
                ws.Cells(PrintRow, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(PrintRow, 11).Interior.ColorIndex = 4
            End If
            
            'Iterate the row for the next print
            PrintRow = PrintRow + 1
        
            TotalVolume = ws.Cells(i + 1, 7).Value
            StockTicker = ws.Cells(i + 1, 1).Value
    
        End If
        
    Next i

Next ws

End Sub

