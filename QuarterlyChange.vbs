Sub CalculateQuarterlyChanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim resultRow As Long
    Dim i As Long
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String

    ' Set the current worksheet
    Set ws = ActiveSheet

    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Add headers to the result columns
    With ws
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Quarterly Change"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Total Stock Volume"
    End With
    resultRow = 2

    ' Initialize variables for max values
    maxIncrease = -99999
    maxDecrease = 99999
    maxVolume = 0

    ' Loop through each ticker symbol
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = lastRow Then
            If i > 2 Then
                If i = lastRow Then
                    closePrice = ws.Cells(i, 6).Value
                Else
                    closePrice = ws.Cells(i - 1, 6).Value
                End If
                quarterlyChange = closePrice - openPrice
                percentChange = (quarterlyChange / openPrice)

                ws.Cells(resultRow, 9).Value = ticker
                ws.Cells(resultRow, 10).Value = quarterlyChange
                ws.Cells(resultRow, 11).Value = Format(percentChange, "0.00%")
                ws.Cells(resultRow, 12).Value = totalVolume

                If quarterlyChange > 0 Then
                    ws.Cells(resultRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(resultRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If

                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If

                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If

                resultRow = resultRow + 1
            End If

            If i < lastRow Then
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
        End If
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    Next i

    ' Format the columns
    With ws
        .Columns("I:L").AutoFit
    End With

    ' Add summary results for max values
    With ws
        .Cells(1, 15).Value = "Greatest % Increase"
        .Cells(1, 16).Value = maxIncreaseTicker
        .Cells(1, 17).Value = Format(maxIncrease, "0.00%")

        .Cells(2, 15).Value = "Greatest % Decrease"
        .Cells(2, 16).Value = maxDecreaseTicker
        .Cells(2, 17).Value = Format(maxDecrease, "0.00%")

        .Cells(3, 15).Value = "Greatest Total Volume"
        .Cells(3, 16).Value = maxVolumeTicker
        .Cells(3, 17).Value = Format(maxVolume, "#,##0")
    End With

    ' Format the summary columns
    ws.Columns("O:Q").AutoFit
End Sub

