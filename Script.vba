Sub SummarizeAndHighlightStockData()

    Dim ws As Worksheet
    Dim summaryRow As Integer
    Dim ticker As String
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim i As Long
    Dim lastRow As Long
    Dim sortRange As Range
   
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String

    ' Set the worksheet to work on
    Set ws = ActiveSheet

    ' Define the range to sort (assuming data starts at row 2 and goes till the last row)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set sortRange = ws.Range("A1:G" & lastRow)
   
    ' Sort data by ticker and then by date
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Initialize the summary row counter
    summaryRow = 2
   
    ' Add headers for the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
   
    ' Initialize max values
    maxIncrease = -1E+100
    maxDecrease = 1E+100
    maxVolume = 0
   
    ' Initialize total volume and opening price
    totalVolume = 0
    openingPrice = 0
   
    ' Loop through all rows of data
    For i = 2 To lastRow + 1
   
        ' Check if we're still on the same ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And i > 2 Then
       
            ' Calculate yearly change and percent change
            yearlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If
           
            ' Write summary data for previous ticker
            With ws
                .Cells(summaryRow, 9).Value = ticker
                .Cells(summaryRow, 10).Value = yearlyChange
                .Cells(summaryRow, 11).Value = percentChange
                .Cells(summaryRow, 11).NumberFormat = "0.00%"
                .Cells(summaryRow, 12).Value = totalVolume
            End With
           
            ' Apply conditional formatting for yearly change
            With ws.Cells(summaryRow, 10).FormatConditions
                .Delete
                .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                .Item(1).Interior.Color = RGB(0, 255, 0)
                .Item(2).Interior.Color = RGB(255, 0, 0)
            End With
           
            ' Check for maximum and minimum percent changes and volumes
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                tickerIncrease = ticker
            End If
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                tickerDecrease = ticker
            End If
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                tickerVolume = ticker
            End If
           
            ' Reset variables for the next ticker
            summaryRow = summaryRow + 1
            totalVolume = 0
            openingPrice = ws.Cells(i, 3).Value
           
        End If
       
        ' Update variables for current row
        If i <= lastRow Then
            If totalVolume = 0 Then
                openingPrice = ws.Cells(i, 3).Value
            End If
            ticker = ws.Cells(i, 1).Value
            closingPrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
       
    Next i
   
    ' Add headers and values for greatest increase, decrease, and volume
    With ws
        .Cells(1, 15).Value = "Category"
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
       
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(2, 16).Value = tickerIncrease
        .Cells(2, 17).Value = maxIncrease
        .Cells(2, 17).NumberFormat = "0.00%"
       
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(3, 16).Value = tickerDecrease
        .Cells(3, 17).Value = maxDecrease
        .Cells(3, 17).NumberFormat = "0.00%"
       
        .Cells(4, 15).Value = "Greatest Total Volume"
        .Cells(4, 16).Value = tickerVolume
        .Cells(4, 17).Value = maxVolume
    End With

End Sub

