Attribute VB_Name = "Module1"
Sub irfan_senyurt_assignment()

    Dim tickername As String
    Dim summary_ticker_row As Integer
    summary_ticker_row = 2

    Dim openingPrice As Double
    Dim closingPrice As Double
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

    For Each sheetname In Array("2018", "2019", "2020")

        Set ws = ThisWorkbook.Worksheets(sheetname)
        ws.Activate

        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        openingPrice = Cells(2, 3).Value
        closingPrice = Cells(2, 6).Value

        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0

        For i = 2 To lastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                tickername = Cells(i, 1).Value
                totalVolume = totalVolume + Cells(i, 7).Value

                closingPrice = Cells(i, 6).Value
                yearlyChange = (closingPrice - openingPrice)
                percentChange = IIf(openingPrice = 0, 0, yearlyChange / openingPrice)

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

                totalVolume = 0
                openingPrice = Cells(i + 1, 3).Value

            Else
                totalVolume = totalVolume + Cells(i, 7).Value

            End If

        Next i

        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("Q2").NumberFormat = "0.00%"

        Range("P2").Value = maxPercentIncreaseTicker
        Range("P3").Value = maxPercentDecreaseTicker
        Range("P4").Value = maxTotalVolumeTicker

        Range("Q2").Value = maxPercentIncrease
        Range("Q2").NumberFormat = "0.00%"
                Range("Q3").Value = maxPercentDecrease
        Range("Q3").NumberFormat = "0.00%"
        Range("Q4").Value = maxTotalVolume

    Next sheetname

End Sub


