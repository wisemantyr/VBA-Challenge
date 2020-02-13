Sub VBA_Wall_Street()


For Each ws in Worksheets

    'label the columns data will be recorded in
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'define variables held in results columns: ticker, yearly change, percent change, total stock
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As String
    Dim TotalStockVolume As Double

    ''define variables in loop
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim CurrentTicker As String
    Dim NextTicker As String
    Dim LastResultsRow As Integer
    Dim UniqueTickerCount As Long
    Dim LastRow As Long

    'set row and column values
    LastResultsRow = 2
    InitialRow = 2
    MasterTickerColumn = 1
    MasterOpenColumn = 3
    MasterCloseColumn = 6
    MasterTotalVolumeColumn = 7

    'set counts
    UniqueTickerCount = 0
    TotalStockVolume = 0

    'last row of original data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'loop through column A
    For RowIndex = InitialRow To LastRow

        'define values to check
        CurrentTicker = ws.Cells(RowIndex, 1).Value
        NextTicker = ws.Cells(RowIndex + 1, 1).Value

        'Add totalstockvolume to itself on each loop (sum)
        TotalStockVolume = TotalStockVolume + ws.Cells(RowIndex, MasterTotalVolumeColumn).Value

        If CurrentTicker = NextTicker Then
            'count number of times the ticker is the same
            UniqueTickerCount = UniqueTickerCount + 1

        'when a new ticker is encountered
        ElseIf CurrentTicker <> NextTicker Then
            
            'store value in ticker variable
            Ticker = ws.Cells(RowIndex, 1).Value
            'print ticker in summary table
            ws.Range("I" & LastResultsRow).Value = Ticker
            
            'grab value of open at first instance of ticker
            StockOpen = ws.Cells(RowIndex - UniqueTickerCount, MasterOpenColumn)
            'grab value of close at last instance of ticker
            StockClose = ws.Cells(RowIndex, MasterCloseColumn)
            YearlyChange = StockClose - StockOpen
            'add yearlychange to summarytable
            ws.Range("J" & LastResultsRow).Value = YearlyChange

            'conditional color formatting
            If ws.Range("J" & LastResultsRow).Value < 0 Then
                ws.Range("J" & LastResultsRow).Interior.ColorIndex = 3 
            Elseif ws.Range("J" & LastResultsRow).Value > 0 Then
                ws.Range("J" & LastResultsRow).Interior.ColorIndex = 4
            End If
            
            'StockOpen 0 exception + calculate, format, place percent change in summary table
            If StockOpen = 0 Then 
                PercentChange = Format(0, "Percent")
            Else
                PercentChange = Format((YearlyChange/StockOpen), "Percent")
            End If
            ws.Range("K" & LastResultsRow).Value = PercentChange
            

            'add sum of totalstockvolume to summary table
            ws.Range("L" & LastResultsRow).Value = TotalStockVolume
            
            'move to the next row in summary table
            LastResultsRow = LastResultsRow + 1

            'reset uniquetickercount and totalstockvolume sum
            UniqueTickerCount = 0
            TotalStockVolume = 0
        End If
    Next RowIndex
ws.Columns("A:L").Autofit

Next ws

End Sub
