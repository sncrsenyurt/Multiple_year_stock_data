Attribute VB_Name = "Module1"
Sub irfan_senyurt_assignment()

    Dim tickername As String
    Dim summary_ticker_row As Integer
    summary_ticker_row = 2

    Dim openingPrice As Double
    openingPrice = Cells(2, 3).Value

    Dim closingPrice As Double
    closingPrice = Cells(2, 6).Value

    Dim yearlyChange As Double
    Dim percentChange As Double

    Dim totalVolume As Double
    totalVolume = 0

    Dim lastRow As Long
    Dim i As Long

    Dim sheetname As Variant
    Dim ws As Worksheet

    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String

    Dim maxTotalVolumeTicker As String

    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double

    Dim maxTotalVolume As Double

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("Q2").NumberFormat = "0.00%"

    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0

    For Each sheetname In Array("2018", "2019", "2020")

        Set ws = ThisWorkbook.Worksheets(sheetname)
        ws.Activate

        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                tickername = Cells(i, 1).Value

                totalVolume = totalVolume + Cells(i, 7).Value

                Range("I" & summary_ticker_row).Value = tickername

                Range("L" & summary_ticker_row).Value = totalVolume

                closingPrice = Cells(i, 6).Value

                yearlyChange = (closingPrice - openingPrice)

                Range("J" & summary_ticker_row).Value = yearlyChange

                If yearlyChange > 0 Then

                    Range("J" & summary_ticker_row).Interior.Color = vbGreen

                ElseIf yearlyChange < 0 Then

                    Range("J" & summary_ticker_row).Interior.Color = vbRed

                End If

                If (openingPrice = 0) Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openingPrice
                End If

                Range("K" & summary_ticker_row).Value = percentChange
                Range("K" & summary_ticker_row).NumberFormat = "0.00%"

                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = tickername
                End If

                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = tickername
                End If

                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = tickername
                End If

                summary_ticker_row = summary_ticker_row + 1
                totalVolume = 0
                openingPrice = Cells(i + 1, 3).Value

            Else
                totalVolume = totalVolume + Cells(i, 7).Value

            End If

        Next i

        summary_ticker_row = 2
        openingPrice = Cells(2, 3).Value
        closingPrice = Cells(2, 6).Value

    Next sheetname

    Range("P2").Value = maxPercentIncreaseTicker
    Range("P3").Value = maxPercentDecreaseTicker
    Range("P4").Value = maxTotalVolumeTicker

    Range("Q2").Value = maxPercentIncrease
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = maxPercentDecrease
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").Value = maxTotalVolume

End Sub

                    


