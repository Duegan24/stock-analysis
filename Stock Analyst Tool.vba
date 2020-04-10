Sub AllStocksAnalysis()

    Dim rowCount As Long
    Dim mainSheet As String
    Dim tickers() As String
    Dim tickerIndex As Integer
    Dim totalVolume(11) As Double
    Dim startingPrice(11) As Double
    Dim endingPrice(11) As Long
    Dim yearValue As String
    Dim numTickersElements As Long


    'Record users response for which year they want analyzed
    yearValue = InputBox("What year would you like to run the analysis on? (format is YYYY)")

    'Loading the Array with the tickers I want to analyze
    tickers = Split("AY,CSIQ,DQ,ENPH,FSLR,HASI,JKS,RUN,SEDG,SPWR,TERP,VSLR", ",")

    'Find the index value for the last ticker to be used to end the for statement later
    numTickersElements = UBound(tickers) - LBound(tickers)

'    Worksheets(yearValue).Activate

    'Sort the data so the follow loops will work no matter what the starting order of the stock data
    With Worksheets(yearValue)
        With .Cells(1, "A").CurrentRegion
            .Cells.Sort Key1:=.Range("A1"), Order1:=xlAscending, _
                        Key2:=.Range("B1"), Order2:=xlAscending, _
                        Orientation:=xlTopToBottom, Header:=xlYes
        End With

        'Find the last row with data and set first row of data
        rowCount = .Cells(Rows.Count, "A").End(xlUp).Row
        rowStart = 2

        'Inialize the ticker index at zero
        tickerIndex = 0
        ticker = tickers(tickerIndex)
        totalVolume(tickerIndex) = 0

        'Set a for loop to cycle through each row of data
        For i = rowStart To rowCount

            If .Cells(i, 1).Value = ticker Then

                'incrementally increase totalVolume if the ticker matches
                totalVolume(tickerIndex) = totalVolume(tickerIndex) + .Cells(i, 8).Value

                'look for first occurance of the ticker and find first closing price
                If .Cells(i - 1, 1).Value <> ticker Then
                    startingPrice(tickerIndex) = .Cells(i, 6).Value
                End If

                'look for last occurance of the ticker and record end price and totalVolume
                If .Cells(i + 1, 1).Value <> ticker Then
                    endingPrice(tickerIndex) = .Cells(i, 6).Value

                    'Once the values for the last ticker is found, exit the loop
                    If tickerIndex = numTickersElements Then Exit For

                    'Setup tickerIndex, Ticker, and totalVoume for the next ticker
                    tickerIndex = tickerIndex + 1
                    ticker = tickers(tickerIndex)
                    totalVolume(tickerIndex) = 0
                End If
            End If
        Next i
    End With



    'Define primary worksheet where the output is going to go
'    mainSheet = "All Stocks Analysis"

    'Focus on sheet where the results will be placed
'    Worksheets(mainSheet).Activate


    'Start Setting up the worksheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'loop throough all the tickers and output and format the analysis collected
    For i = 0 To tickerIndex
      'Output Result into main output worksheet
      Cells(i + 4, 1).Value = tickers(i)
      Cells(i + 4, 2).Value = totalVolume(i)
      Cells(i + 4, 3).Value = endingPrice(i) / startingPrice(i) - 1


      'Format cell interior based on the return value for each stock
      If Cells(i + 4, 3).Value > 0 Then
          Cells(i + 4, 3).Interior.Color = vbGreen
      ElseIf Cells(i + 4, 3).Value < 0 Then
          Cells(i + 4, 3).Interior.Color = vbRed
      Else
          Cells(i + 4, 4).Interior.Color = xlNone
      End If
    Next i



    'Formating to make everything look better

        Range("A1:C3").Font.Bold = True
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlDouble
        Range("b4:b" & tickerIndex + 3).NumberFormat = "#,##0"
        Range("c4:c" & tickerIndex + 3).NumberFormat = "0.0%"
        Range("a4:c" & tickerIndex + 3).Borders.LineStyle = xlContinuous
        Rows("3:" & tickerIndex + 3).Columns.AutoFit


End Sub
