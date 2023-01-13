' Create a script that loops through all the stocks for one year and outputs the following information:
    ' The ticker symbol
    ' Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.

' Add functionality to your script to return the stock with the
' "Greatest % increase", "Greatest % decrease", and "Greatest
' total volume". 

Sub challenge2():
    
    Dim yearlyChange As Double
    Dim percentChange As Double

    Dim totalTickerVolume As LongLong
    totalTickerVolume = 0

    Dim summaryRow As Long
    summaryRow = 2
    
    Dim openPrice As Double
    openPrice = Cells(2, 3).Value

    Dim closePrice As Double

    Dim maxPercent As Double
    maxPercent = 0
    Dim minPercent As Double
    minPercent = 0
    Dim maxVol As LongLong
    maxVol = 0
    Dim maxPercentTicker As String
    Dim minPercentTicker As String
    Dim maxVolTicker As String

    ' Naming the columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    ' Looping through each row
    ' If the row is neither the start or the end of the ticker data, just add the stock volume.
    ' Else, compute variables.
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            totalTickerVolume = totalTickerVolume + Cells(i, 7).Value
        Else:
            Range("I" & summaryRow).Value = Cells(i, 1).Value

            closePrice = Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            Range("J" & summaryRow).Value = yearlyChange
            If yearlyChange >= 0 Then
                Range("J" & summaryRow).Interior.ColorIndex = 4
            Else
                Range("J" & summaryRow).Interior.ColorIndex = 3
            End If

            percentChange = yearlyChange / openPrice
            Range("K" & summaryRow).Value = percentChange
            Range("K" & summaryRow).NumberFormat = "0.00%"
            If percentChange > maxPercent Then
                maxPercent = percentChange
                maxPercentTicker = Range("I" & summaryRow).Value
            End If
            If percentChange < minPercent Then
                minPercent = percentChange
                minPercentTicker = Range("I" & summaryRow).Value
            End If

            totalTickerVolume = totalTickerVolume + Cells(i, 7).Value
            Range("L" & summaryRow).Value = totalTickerVolume

            If totalTickerVolume > maxVol Then
                maxVol = totalTickerVolume
                maxVolTicker = Range("I" & summaryRow).Value
            End If

            ' Setting up next summary
            totalTickerVolume = 0
            openPrice = Cells(i + 1, 3).Value
            summaryRow = summaryRow + 1
        End If
    Next i

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P2").Value = maxPercentTicker
    Range("P3").Value = minPercentTicker
    Range("P4").Value = maxVolTicker
    Range("Q2").Value = maxPercent
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = minPercent
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").Value = maxVol
End Sub