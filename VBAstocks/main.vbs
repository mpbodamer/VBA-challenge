Sub loopThroughWorksheets()
    Dim xSh As Worksheet
    For Each xSh In Worksheets
        xSh.Select
        Call ticker()
    Next
End Sub


Sub ticker()
    Dim summaryTableRow As Double
    Dim lastTicker As String
    Dim yearOpenPrice As Double
    Dim yearClosePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double

    totalVolume = 0
    summaryTableRow = 2
        
    'Print Header:
    Cells(1,9).Value = "Ticker"
    Cells(1,10).Value = "Yearly Change"
    Cells(1,11).Value = "Percent Change"
    Cells(1,12).Value = "Total Stock Volume"

    'Find yearly open price for the first ticker
    yearOpenPrice = Cells(2, 3).Value 

    ' Determine the Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    lastTicker = Range("A2")

    For I = 2 To LastRow
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then 'If next ticker is not the same as the current one, then

            'Set ticker name
            lastTicker = Cells(I, 1).Value
            
            'add the last volume onto the totaVolume
            totalVolume = totalVolume + Cells(I, 7).Value

            'Find the yearly closing price
            yearClosePrice = Cells(I, 6).Value

            'Find the yearly change
            yearlyChange = yearClosePrice - yearOpenPrice

            'Find Percent Change
            if yearOpenPrice = 0 Then
                Range("K" & summaryTableRow).Value = "No Opening Price"
            Else
                percentChange = yearlyChange / yearOpenPrice
                Range("K" & summaryTableRow).Value = percentChange
            End If


            'Print results
            Range("I" & summaryTableRow).Value = lastTicker
            Range("L" & summaryTableRow).Value = totalVolume
            Range("J" & summaryTableRow).Value = yearlyChange

            'Move down the summary Table
            summaryTableRow = summaryTableRow + 1
    
            'Reset volume total
            totalVolume = 0

            'Find opening price for next ticker
            yearOpenPrice = Cells(I + 1, 3).Value

        Else 'Same ticker as before

            ' Add row volume to volume total
            totalVolume = totalVolume + Cells(I, 7).Value

        End If
    Next I
    

End Sub
