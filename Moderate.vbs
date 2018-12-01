Option Explicit

Sub Moderate():
    
   
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

        ' Create forloop to look in Column A for unique tickers and add up total stock volume
        For iRow = 2 To rowCount
            If ws.Range("A" & iRow).Value = currentTicker Then
                stockVol = stockVol + ws.Range("G" & iRow).Value
                ws.Range("J" & currentTickerRow).Value = stockVol
                
            Else
                stockVol = ws.Range("G" & iRow).Value
                currentTickerRow = currentTickerRow + 1
                currentTicker = ws.Range("A" & iRow).Value
                ws.Range("I" & currentTickerRow).Value = currentTicker
                ws.Range("J" & currentTickerRow).Value = stockVol
            End If
        Next iRow

        ' Insert 2 columns after Ticker column
        ws.Columns("J:K").EntireColumn.Insert

        ' label new columns
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"

        'Reset current Ticker and current Ticker Row
        currentTicker = ws.Range("A2").Value
        currentTickerRow = 2

        Dim firstTickerRow As Long
        Dim lastTickerRow As Long

        firstTickerRow = 2

        Dim openPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double

        openPrice = ws.Range("C2").Value
        
        ' loop through columns again finidng the first and
        ' last ticker entry for each unique ticker
        ' Then calculate yearly change and percent change
        For iRow = 2 To rowCount
            If iRow = rowCount Then
                closingPrice = ws.Range("F" & iRow).Value
                yearlyChange = closingPrice - openPrice
                ws.Range("J" & currentTickerRow).Value = yearlyChange

                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If

                ws.Range("K" & currentTickerRow).Value = percentChange

            ElseIf currentTicker <> ws.Range("A" & iRow).Value Then

                lastTickerRow = iRow - 1
                closingPrice = ws.Range("F" & lastTickerRow).Value

                yearlyChange = closingPrice - openPrice
                ws.Range("J" & currentTickerRow).Value = yearlyChange
                
                'Check if openPrice = 0 so you dont divide by 0
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If
                ws.Range("K" & currentTickerRow).Value = percentChange
                currentTicker = ws.Range("A" & iRow).Value
                currentTickerRow = currentTickerRow + 1
                openPrice = ws.Range("C" & iRow).Value
                
            End If
        Next iRow
        
        ' format percent change column to percentage with 2 decimals
        ws.Range("K2", ws.Range("K2").End(xlDown)).NumberFormat = "0.00%"

        'define rules for conditional formatting
        Dim cond1 As FormatCondition
        Dim cond2 As FormatCondition

        Set cond1 = ws.Range("J2", ws.Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlGreater, "0")
        Set cond2 = ws.Range("J2", ws.Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlLess, "0")

        'set cond1 as greater than 0, highlight green
        'set cond2 as less than 0, highlight red
        With cond1
            .Interior.Color = vbGreen
        End With

        With cond2
            .Interior.Color = vbRed
        End With


    Next ws

End Sub


