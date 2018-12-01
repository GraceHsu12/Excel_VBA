Option Explicit

Sub Easy():
    
   
    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
    
    
        Dim iRow As Long
        Dim rowCount As Long
        rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row 

    
        Dim currentTicker As String
        currentTicker = ws.Range("A2").Value
        
        ws.Range("I2").Value = currentTicker

        Dim currentTickerRow As Long
        currentTickerRow = 2


        Dim stockVol As Variant

        ' Create forloop to look in Column A for unique values
        For iRow = 2 To rowCount
            If ws.Range("A" & iRow).Value = currentTicker Then
                stockVol = stockVol + ws.Range("G" & iRow).Value
                ws.Range("J" & currentTickerRow).Value = stockVol
                
            Else
                stockVol = ws.Range("G"&iRow).Value
                currentTickerRow = currentTickerRow + 1
                currentTicker = ws.Range("A" & iRow).Value
                ws.Range("I" & currentTickerRow).Value = currentTicker
                ws.Range("J"&currentTickerRow).Value = stockVol
            End If
        Next iRow

    Next ws

End Sub



