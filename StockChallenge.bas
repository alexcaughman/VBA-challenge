Attribute VB_Name = "Module1"
Sub VBAChallenge()

    Dim ticker As String
    Dim totalVolume As Variant
    Dim outputRow As Integer, lastRow As Long
    Dim yearOpen As Double, yearClose As Double
    Dim yearChange As Double, percentChange As Double
    Dim previousTicker As String, currentTicker As String, nextTicker As String
    
    Dim greatestTotalVolumeValue As Variant, greatestTotalVolumeTicker As String
    Dim greatestPercentIncreaseValue As Double, greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseValue As Double, greatestPercentDecreaseTicker As String
    
    
    For Each ws In Worksheets
        ' Set starting values
        outputRow = 2
        yearOpen = ws.Range("C2").Value
        greatestTotalVolumeValue = 0
        greatestPercentIncreaseValue = 0
        greatestPercentDecreaseValue = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        ' Set labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
    
        For i = 2 To lastRow
            If i > 2 Then
                previousTicker = ws.Range("A" & (i - 1)).Value
            Else
                previousTicker = ""
            End If
            
            currentTicker = ws.Range("A" & i).Value
            nextTicker = ws.Range("A" & (i + 1)).Value
            
            If previousTicker <> "" And previousTicker <> currentTicker Then
                yearClose = ws.Range("F" & (i - 1)).Value
                yearChange = (yearClose - yearOpen)
                If yearOpen > 0 Then
                    percentChange = (yearClose - yearOpen) / yearOpen
                Else
                    percentChange = 0
                End If
                ws.Range("I" & outputRow).Value = previousTicker
                ws.Range("J" & outputRow).Value = yearChange
                ws.Range("K" & outputRow).Value = percentChange
                ws.Range("L" & outputRow).Value = totalVolume
                
                ' track min/max ticker(s)
                If percentChange > greatestPercentIncreaseValue Then
                    greatestPercentIncreaseValue = percentChange
                    greatestPercentIncreaseTicker = previousTicker
                End If
                
                If percentChange < greatestPercentDecreaseValue Then
                    greatestPercentDecreaseValue = percentChange
                    greatestPercentDecreaseTicker = previousTicker
                End If
            
                If totalVolume > greatestTotalVolumeValue Then
                    greatestTotalVolumeValue = totalVolume
                    greatestTotalVolumeTicker = previousTicker
                End If
                
                ' Set cell color by value
                If yearChange >= 0 Then
                    ws.Range("J" & outputRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & outputRow).Interior.ColorIndex = 3
                End If
                
                ' set correct cell format for %
                ws.Range("K" & outputRow).NumberFormat = "0.00%"
            
                outputRow = outputRow + 1
                yearOpen = ws.Range("C" & i)
                totalVolume = ws.Range("G" & i).Value
            ElseIf currentTicker = nextTicker Then
                totalVolume = totalVolume + ws.Range("G" & i).Value
            End If
        Next i
        
        ' finally, report out min/max tickers
        ws.Range("O2").Value = greatestPercentIncreaseTicker
        ws.Range("P2").Value = greatestPercentIncreaseValue
        ws.Range("P2").NumberFormat = "0.00%"
        
        ws.Range("O3").Value = greatestPercentDecreaseTicker
        ws.Range("P3").Value = greatestPercentDecreaseValue
        ws.Range("P3").NumberFormat = "0.00%"
        
        ws.Range("O4").Value = greatestTotalVolumeTicker
        ws.Range("P4").Value = greatestTotalVolumeValue
    Next
End Sub





