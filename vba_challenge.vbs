Option Explicit

' Challenge 2 Code

Sub TotalAllStocksInAllSheets()

    ' Loop over all worksheets in workbook
    Dim ws As Worksheet
    For Each ws In Worksheets
        ' Use With ws so we don't have to write ws.abc for so many lines
        With ws
            'Add Header Row for totals
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            
            'Declare variables
            Dim currentOpenValue As Currency
            Dim currentVolumeTotal As LongLong
            Dim totalsCurrentRow As Integer
            
            currentOpenValue = .Range("C2").Value   ' First open value in the sheet
            currentVolumeTotal = 0                  ' Starts with 0 total volume
            totalsCurrentRow = 2                    ' Total table starts at row 2
            
            ' Loop over all rows in worksheet
            Dim rowIndex As Long
            For rowIndex = 2 To .UsedRange.Rows.Count
                If .Cells(rowIndex + 1, "A").Value <> .Cells(rowIndex, "A").Value Then
                    ' Ticker has changed, calculate values, update total table
                    
                    Dim yearlyChange As Double
                    yearlyChange = .Cells(rowIndex, "F") - currentOpenValue
                    
                    ' Set the values in the total table
                    .Range("I" & totalsCurrentRow).Value = .Cells(rowIndex, "A")
                    .Range("J" & totalsCurrentRow).Value = yearlyChange
                    .Range("K" & totalsCurrentRow).Value = (.Cells(rowIndex, "F") / currentOpenValue) - 1
                    .Range("L" & totalsCurrentRow).Value = currentVolumeTotal + .Cells(rowIndex, "G")
                    
                    ' Format the new cells in Totals table
                    .Range("J" & totalsCurrentRow).NumberFormat = "#,##0.00"
                    .Range("K" & totalsCurrentRow).NumberFormat = "0.00%"
                    ' Conditional formatting (Green or Red) (0.00 values are not colored)
                    If yearlyChange > 0 Then
                        .Range("J" & totalsCurrentRow).Interior.Color = RGB(0, 255, 0)
                    ElseIf yearlyChange < 0 Then
                        .Range("J" & totalsCurrentRow).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    ' Reset/Update values for next stock
                    currentOpenValue = .Cells(rowIndex + 1, "C")
                    currentVolumeTotal = 0
                    totalsCurrentRow = totalsCurrentRow + 1
                Else
                    ' Next Ticker is the same stock, just add to VolumeTotal
                    currentVolumeTotal = currentVolumeTotal + .Cells(rowIndex, "G")
                End If
            Next rowIndex
            
            
            ' Bonus ----------------------------
            ' Setup labels
            .Range("N2").Value = "Greatest % Increase"
            .Range("N3").Value = "Greatest % Decrease"
            .Range("N4").Value = "Greatest Total Volume"
            .Range("O1").Value = "Ticker"
            .Range("P1").Value = "Value"
            ' Setup variables
            Dim lastRowTotalTable As Long
            Dim maxPercentIncrease As Double
            Dim maxPercentDecrease As Double
            Dim maxTotalVolume As LongLong
            lastRowTotalTable = .Cells(Rows.Count, "K").End(xlUp).row
            
            ' Get Max values in Totals table
            maxPercentIncrease = Application.Max(.Range("K2:K" & lastRowTotalTable))
            maxPercentDecrease = Application.Min(.Range("K2:K" & lastRowTotalTable))
            maxTotalVolume = Application.Max(.Range("L2:L" & lastRowTotalTable))
            
            ' Loop through Totals table to find the Ticker for the different max values
            Dim rowIndexTotals As Long
            For rowIndexTotals = 2 To lastRowTotalTable
            
                If .Cells(rowIndexTotals, "K").Value = maxPercentIncrease Then
                    .Range("O2").Value = .Range("I" & rowIndexTotals).Value
                ElseIf .Cells(rowIndexTotals, "K").Value = maxPercentDecrease Then
                    .Range("O3").Value = .Range("I" & rowIndexTotals).Value
                End If
                ' max Total Volume could be the same as one of the % changes
                If .Cells(rowIndexTotals, "L").Value = maxTotalVolume Then
                    .Range("O4").Value = .Range("I" & rowIndexTotals).Value
                End If
            
            Next rowIndexTotals
            
            ' Set the values in the Bonus table
            .Range("P2").Value = maxPercentIncrease
            .Range("P3").Value = maxPercentDecrease
            .Range("P4").Value = maxTotalVolume
            .Range("P2:P3").NumberFormat = "0.00%"
            .Range("P4").NumberFormat = "#,##0"
        
            ' Adjust the new column widths to fit
            .Columns("I:P").AutoFit
            
        End With
        
    Next ws
End Sub
