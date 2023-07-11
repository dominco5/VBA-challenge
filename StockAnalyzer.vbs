Sub StockAnalyzer()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    'creating table1
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'creating table2
    ws.Range("Q1") = "Ticker"
    ws.Range("R1") = "Value"
    ws.Range("P2") = "Greatest % Increase"
    ws.Range("P3") = "Greatest % Decrease"
    ws.Range("P4") = "Greatest Total Volume"
    
    'find last row of dataset
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ticker = ws.Range("A2")
    Dim totalVol As Double
    totalVol = 0
    
    yearOpen = ws.Range("C2")
    Dim yearClose As Single
    yearClose = 0
    Dim yearChange, percentChange As Single
    
    Dim percentIncTicker, percentDecTicker, greatestVolTicker As String
    Dim percentInc, percentDec, greatestVol As Single
    
    summaryTableRow = 2
    
    'data for table1
    For i = 2 To lastRow:
        If ticker <> ws.Cells(i + 1, 1) Then
            totalVol = totalVol + ws.Cells(i, 7).Value
            
            yearClose = ws.Cells(i, 6)
            yearChange = yearClose - yearOpen
            percentChange = yearChange / yearOpen
            
            'print values for table1
            ws.Cells(summaryTableRow, 9) = ws.Cells(i, 1).Value
            ws.Cells(summaryTableRow, 10) = yearChange
                If yearChange > 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                ElseIf yearChange < 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                End If
            ws.Cells(summaryTableRow, 11) = percentChange
            ws.Cells(summaryTableRow, 12) = totalVol
            
            'find greatest % increase and decrease for table2
            If percentChange > percentInc Then
                percentIncTicker = ticker
                percentInc = percentChange
            ElseIf percentChange < percentDec Then
                percentDecTicker = ticker
                percentDec = percentChange
            End If
            
            If totalVol > greatestVol Then
                greatestVol = totalVol
                greatestVolTicker = ticker
            End If
            
            ticker = ws.Cells(i + 1, 1).Value
            totalVol = 0
            'reassign yearOpen
            yearOpen = ws.Cells(i + 1, 3)
            summaryTableRow = summaryTableRow + 1
        Else
            'add volume for day to total volume
            totalVol = totalVol + ws.Cells(i, 7).Value
        End If
        
    Next i
    
    'print values for table2
    ws.Range("Q2") = percentIncTicker
    ws.Range("R2") = percentInc
    ws.Range("Q3") = percentDecTicker
    ws.Range("R3") = percentDec
    ws.Range("Q4") = greatestVolTicker
    ws.Range("R4") = greatestVol
    
    'formatting
    lastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To lastRow2:
        ws.Cells(i, 10).Value = Round(ws.Cells(i, 10).Value, 2)
        ws.Cells(i, 11).Value = FormatPercent(ws.Cells(i, 11).Value, 2)
    Next i
    
    ws.Range("R2").Value = FormatPercent(ws.Range("R2").Value, 2)
    ws.Range("R3").Value = FormatPercent(ws.Range("R3").Value, 2)
    
    ws.UsedRange.EntireColumn.AutoFit

    
    Next
    
End Sub



