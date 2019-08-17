Sub StockData():
Dim i As LongLong
Dim Stock As String
Dim TSV As LongLong
Dim YC As Double
Dim PC As Double

TSV = 0
YC = 0
PC = 0

ticker = 2
firstopen = 0

For Each ws In Worksheets
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Stock"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume (TSV)"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    For i = 2 To lastrow
        If firstopen = 0 Then
            firstopen = ws.Cells(i, 3).Value
        
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Stock = ws.Cells(i, 1).Value
            TSV = TSV + ws.Cells(i, 7).Value
            YC = (ws.Cells(i, 6).Value - firstopen)
            If firstopen = 0 Then
                PC = 0
            
            Else
            PC = (YC) / firstopen
            
            End If
             
            ws.Range("I" & ticker).Value = Stock
            ws.Range("J" & ticker).Value = YC
            ws.Range("K" & ticker).Value = PC
            ws.Range("L" & ticker).Value = TSV
            
            ticker = ticker + 1
            TSV = 0
            YC = 0
            firstopen = 0
            PC = 0

        Else:
            TSV = TSV + ws.Cells(i, 7).Value
            YC = YC + ws.Cells(i, 6).Value

        End If
    Next i
    TSV = 0
    YC = 0
    ticker = 2
    firstopen = 0
    PC = 0
        
        lastcalculatedrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        maxpcpos = 0
        maxpcneg = 0
        maxTSV = 0
        maxpercticker = ""
        minpercticker = ""
        maxvolticker = ""
        For i = 2 To lastcalculatedrow
        
            If ws.Cells(i, "L").Value >= maxTSV Then
                maxTSV = ws.Cells(i, "L").Value
                maxvolticker = ws.Cells(i, "I").Value
            End If
        
             If ws.Cells(i, "K").Value >= 0 Then
                 ws.Range("K" & i).Interior.ColorIndex = 4
                 If ws.Cells(i, "K").Value >= maxpcpos Then
                    maxpcpos = ws.Cells(i, "K").Value
                    maxpercticker = ws.Cells(i, "I").Value
                End If
            Else:
                ws.Range("K" & i).Interior.ColorIndex = 3
                If ws.Cells(i, "K").Value <= maxpcneg Then
                    maxpcneg = ws.Cells(i, "K").Value
                    minpercticker = ws.Cells(i, "I").Value
                End If
            
            End If
        Next i
        ws.Cells(4, "P").Value = maxTSV
        ws.Cells(2, "P").Value = maxpcpos
        ws.Cells(3, "P").Value = maxpcneg
        ws.Cells(4, "O").Value = maxvolticker
        ws.Cells(3, "O").Value = minpercticker
        ws.Cells(2, "O").Value = maxpercticker
Next ws

End Sub
