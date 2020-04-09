Sub AllStocksAnalysis()

    Dim rowCount As Long
    Dim mainSheet As String
    Dim tickers() As String
    Dim tickerIndex As Integer
    Dim totalVolume(11) As Double
    Dim startingPrice(11) As Double
    Dim endingPrice(11) As Long
    Dim yearValue As String

    'Record users response for which year they want analyzed
    yearValue = InputBox("What year would you like to run the analysis on? (format is YYYY)")

    'Define primary worksheet where the output is going to go
    mainSheet = "All Stocks Analysis"

    'Focus on the primary worksheet
    Worksheets(mainSheet).Activate

    'Start Setting up the worksheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Loading the Array with the tickers I want to analyze
    tickers = Split("AY,CSIQ,DQ,ENPH,FSLR,HASI,JKS,RUN,SEDG,SPWR,TERP,VSLR", ",")


    Worksheets(yearValue).Activate

    'Find the last row with data and set first row of data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    rowStart = 2

    'Sort the data so the follow loops will work no matter what the starting order of the stock data
    With Worksheets(yearValue)
        With .Cells(1, "A").CurrentRegion
            .Cells.Sort Key1:=.Range("A1"), Order1:=xlAscending, _
                        Key2:=.Range("B1"), Order2:=xlAscending, _
                        Orientation:=xlTopToBottom, Header:=xlYes
        End With
    End With

    'Inialize the ticker index at zero
    tickerIndex = 0
    ticker = tickers(tickerIndex)
    totalVolume(tickerIndex) = 0

    'Set a for loop to cycle through each row of data
    For i = rowStart To rowCount

        If Cells(i, 1).Value = ticker Then

            'incrementally increase totalVolume if the ticker matches
            totalVolume = totalVolume + Cells(i, 8).Value

            'look for first occurance of the ticker and find first closing price
            If Cells(i - 1, 1).Value <> ticker Then
                startingPrice(tickerIndex) = Cells(i, 6).Value
            End If

            'look for last occurance of the ticker and record end price and totalVolume
            If Cells(i + 1, 1).Value <> ticker Then
                endingPrice(tickerIndex) = Cells(i, 6).Value
                totalVolume(tickerIndex)=totalVolume

                'Setup tickerIndex, Ticker, and totalVoume for the next ticker
                tickerIndex = tickerIndex + 1
                ticker = tickers(tickerIndex)
                totalVolume(tickerIndex) = 0
                totalVolume = 0
            End If
        End If
    Next i

        'Output Result into main output worksheet
    With Worksheets(mainSheet)
        for i = 0 to tickerIndex
          .Cells(i + 4, 1).Value = tickers(i)
          .Cells(i + 4, 2).Value = totalVolume(i)
          .Cells(i + 4, 3).Value = endingPrice(i) / startingPrice(i) - 1
        next i
    end With

            'Format cell interior based on the return
            If .Cells(i + 4, 3).Value > 0 Then
                .Cells(i + 4, 3).Interior.Color = vbGreen
            ElseIf .Cells(i + 4, 3).Value < 0 Then
                .Cells(i + 4, 3).Interior.Color = vbRed
            Else
                .Cells(i + 4, 4).Interior.Color = xlNone
            End If

        End With

    Next i


    'Formating to make everything look better
    Worksheets(mainSheet).Activate
        Range("A1:C3").Font.Bold = True
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlDouble
        Range("b4:b" & i + 3).NumberFormat = "#,##0"
        Range("c4:c" & i + 3).NumberFormat = "0.0%"
        Range("a4:c" & i + 3).Borders.LineStyle = xlContinuous
        Rows("3:" & i + 3).Columns.AutoFit


End Sub
