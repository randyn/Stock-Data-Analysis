Sub stockMarketAnalysis()
    Dim ticker As String
    Dim tickerIndex As Long
    Dim currentWorksheet As Worksheet
    
    For Each currentWorksheet In Worksheets
        Call resizeColumns(currentWorksheet)
        Call recordTickerStats(currentWorksheet)
        Call recordAggregateStats(currentWorksheet)
    Next currentWorksheet
End Sub

Sub resizeColumns(sheet As Worksheet)
    sheet.Columns("J").ColumnWidth = 12
    sheet.Columns("K").ColumnWidth = 12.5
    sheet.Columns("L").ColumnWidth = 16
    sheet.Columns("O").ColumnWidth = 18
End Sub

Sub recordTickerStats(sheet As Worksheet)
    Dim ticker As String
    Dim tickerIndex As Long
    Dim yearlyChange As Double
    Dim startingPrice As Double
    Dim percentChange As Double
    Dim lastTickerRow As Long

    Call setUpStatsHeaders(sheet)

    ' Set starting values
    ticker = sheet.Range("A2").Value
    tickerIndex = 2
    ' Loop through tickers and record volume
    Do While ticker <> ""
        sheet.Cells(tickerIndex, tickerColumn()).Value = ticker
        ' Record volume
        lastTickerRow = getLastRowOfTicker(ticker, sheet.Range("A:A"))
        sheet.Cells(tickerIndex, totalVolumeColumn()).Value = getTotalVolume(ticker, sheet.Range("A:A"), sheet, lastTickerRow)
        
        ' Record yearly change
        yearlyChange = getYearlyChange(ticker, sheet.Range("A:A"), sheet, lastTickerRow)
        sheet.Cells(tickerIndex, yearlyChangeColumn()).Value = yearlyChange
        ' Yearly change formatting
        If (yearlyChange >= 0) Then
            sheet.Cells(tickerIndex, yearlyChangeColumn()).Interior.ColorIndex = 4
        Else
            sheet.Cells(tickerIndex, yearlyChangeColumn()).Interior.ColorIndex = 3
        End If

        ' Record Percent Change
        startingPrice = getStartingPrice(ticker, sheet.Range("A:A"), sheet)
        If (startingPrice = 0) Then
            percentChange = 0
        Else
            percentChange = yearlyChange / startingPrice
        End If
        Call recordPercentChange(tickerIndex, percentChange, sheet)
        sheet.Range("K:K").NumberFormat = "0.00%"

        ' Get next ticker and increment recorded stats index
        ticker = getNextTicker(ticker, sheet.Range("A:A"), lastTickerRow)
        tickerIndex = tickerIndex + 1
    Loop
End Sub

Sub recordAggregateStats(sheet As Worksheet)
    Call recordGreatestStatsLabels(sheet)
    Call recordGreatestIncrease(sheet)
    Call recordGreatestDecrease(sheet)
    Call recordGreatestVolume(sheet)
End Sub

Sub recordGreatestIncrease(sheet As Worksheet)
    Dim tickerIndex As Long

    tickerIndex = tickerIndexForMaxStat(sheet.Range("K:K"))
    sheet.Cells(2, 16).Value = tickerStatName(tickerIndex, sheet)
    sheet.Cells(2, 17).Value = tickerStatPercentage(tickerIndex, sheet)
    sheet.Cells(2, 17).NumberFormat = "0.00%"
End Sub

Sub recordGreatestDecrease(sheet As Worksheet)
    Dim tickerIndex As Long
    
    tickerIndex = tickerIndexForMinStat(sheet.Range("K:K"))
    sheet.Cells(3, 16).Value = tickerStatName(tickerIndex, sheet)
    sheet.Cells(3, 17).Value = tickerStatPercentage(tickerIndex, sheet)
    sheet.Cells(3, 17).NumberFormat = "0.00%"
End Sub

Sub recordGreatestVolume(sheet As Worksheet)
    Dim tickerIndex As Long

    tickerIndex = tickerIndexForMaxStat(sheet.Range("L:L"))
    sheet.Cells(4, 16).Value = tickerStatName(tickerIndex, sheet)
    sheet.Cells(4, 17).Value = tickerStatVolume(tickerIndex, sheet)
End Sub

Sub recordGreatestStatsLabels(sheet As Worksheet)
    sheet.Range("O2").Value = "Greatest % increase"
    sheet.Range("O3").Value = "Greatest % Decrease"
    sheet.Range("O4").Value = "Greatest total volume"
    sheet.Range("P1").Value = "Ticker"
    sheet.Range("Q1").Value = "Value"
End Sub

Function tickerIndexForMinStat(statColumn As Range)
    Dim highestDecrease As Double
    Dim tickerIndex As Long

    highestDecrease = WorksheetFunction.Min(statColumn)

    tickerIndexForMinStat = WorksheetFunction.Match(highestDecrease, statColumn, 0)
    
End Function

Function tickerIndexForMaxStat(statColumn As Range)
    Dim highestIncrease As Double

    highestIncrease = WorksheetFunction.Max(statColumn)

    tickerIndexForMaxStat = WorksheetFunction.Match(highestIncrease, statColumn, 0)
End Function

Sub recordAggregateStat(tickerIndex As Long, aggStat As String, sheet As Worksheet)
    Dim aggStatIndex As Long
    Dim ticker As String
    Dim aggStatValue As Double

    ticker = tickerStatName(tickerIndex, sheet)
    aggStatIndex = aggregateStateIndex(aggStat)

    sheet.Cells(aggStatIndex, 16).Value = "ticker"
    sheet.Cells(aggStatIndex, 17).Value = "value"
End Sub

Function tickerStatValue(tickerIndex As Long, aggStat As String, sheet As Worksheet)
    
End Function

Function tickerStatName(tickerIndex As Long, sheet As Worksheet)
    tickerStatName = sheet.Cells(tickerIndex, tickerColumn()).Value
End Function

Function tickerStatPercentage(tickerIndex As Long, sheet As Worksheet)
    tickerStatPercentage = sheet.Cells(tickerIndex, percentChangeColumn()).Value
End Function

Function tickerStatVolume(tickerIndex As Long, sheet As Worksheet)
    tickerStatVolume = sheet.Cells(tickerIndex, totalVolumeColumn()).Value
End Function

Function aggregateStatIndex(aggStat As String)
    Dim aggStatIndex As Long
    Select Case aggStat
        Case Is = "Greatest % increase"
            aggStatIndex = 2
        Case Is = "Greatest % Decrease"
            aggStatIndex = 3
        Case Is = "Greatest total volume"
            aggStatIndex = 4
    End Select
    
    aggregateStatIndex = aggStatIndex
End Function

Sub setUpStatsHeaders(sheet As Worksheet)
    sheet.Cells(1, tickerColumn()).Value = "Ticker"
    sheet.Cells(1, yearlyChangeColumn()).Value = "Yearly Change"
    sheet.Cells(1, percentChangeColumn()).Value = "Percent Change"
    sheet.Cells(1, totalVolumeColumn()).Value = "Total Stock Volume"
End Sub

Sub recordPercentChange(tickerIndex As Long, percentChange As Double, sheet As Worksheet)
    sheet.Cells(tickerIndex, percentChangeColumn()).Value = percentChange
End Sub

Function tickerColumn()
    tickerColumn = 9
End Function

Function yearlyChangeColumn()
    yearlyChangeColumn = 10
End Function

Function percentChangeColumn()
    percentChangeColumn = 11
End Function

Function totalVolumeColumn()
    totalVolumeColumn = 12
End Function

Function getNextTicker(currentTicker As String, column As Range, lastTickerRow As Long)
    getNextTicker = column.Cells(lastTickerRow + 1, 1).Value
End Function

Function getFirstRowOfTicker(ticker As String, column As Range)
    getFirstRowOfTicker = WorksheetFunction.Match(ticker, column, 0)
End Function

Function getLastRowOfTicker(ticker As String, column As Range)
    Dim firstRow As Long
    Dim lastTotalRow As Long

    firstRow = getFirstRowOfTicker(ticker, column)
    lastTotalRow = getLastRow(column)

    For i = firstRow To lastTotalRow + 1
        If (column.Cells(i + 1, 1).Value <> ticker) Then
            getLastRowOfTicker = i
            Exit For
        End If
    Next i
End Function

Function getTotalVolume(ticker As String, column As Range, sheet As Worksheet, lastTickerRow As Long)
    Dim firstRowOfTicker As Long
    Dim totalVolume As Double
    
    firstRowOfTicker = getFirstRowOfTicker(ticker, column)
    
    For i = firstRowOfTicker To lastTickerRow
        If (sheet.Cells(i, 7).Value > 0) Then
            totalVolume = totalVolume + sheet.Cells(i, 7).Value
        End If
    Next i
    
    getTotalVolume = totalVolume
End Function

Function getYearlyChange(ticker As String, tickerColumn As Range, sheet As Worksheet, lastTickerRow As Long)
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    
    startingPrice = getStartingPrice(ticker, tickerColumn, sheet)
    endingPrice = getEndingPrice(ticker, tickerColumn, sheet, lastTickerRow)
    
    getYearlyChange = endingPrice - startingPrice
End Function

Function getStartingPrice(ticker As String, tickerColumn As Range, sheet As Worksheet)
    Dim firstTickerRow As Long
    
    firstTickerRow = getFirstRowOfTicker(ticker, tickerColumn)
    getStartingPrice = sheet.Cells(firstTickerRow, 3).Value
End Function

Function getEndingPrice(ticker As String, tickerColumn As Range, sheet As Worksheet, lastTickerRow As Long)
    getEndingPrice = sheet.Cells(lastTickerRow, 6).Value
End Function

Function getLastRow(column As Range)
    getLastRow = WorksheetFunction.CountA(column)
End Function

