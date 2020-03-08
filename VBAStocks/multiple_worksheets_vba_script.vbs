Sub stocktracker():
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total stock volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim i As Long
        Dim j As Integer
        Dim lrow As Long
        Dim sumrow As Long
        Dim ticker As String
        Dim openprice As Double
        Dim closingprice As Double
        
        
        sumrow = 2
        
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        openprice = ws.Cells(2, 3).Value
        
        For i = 2 To lrow
            ws.Cells(sumrow, 12).Value = ws.Cells(sumrow, 12).Value + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ticker = ws.Cells(i, 1).Value
                    ws.Cells(sumrow, 9).Value = ticker
                    
                    closingprice = ws.Cells(i, 6).Value
                    ws.Cells(sumrow, 10).Value = closingprice - openprice
                    If openprice <> 0 Then
                        ws.Cells(sumrow, 11).Value = (closingprice - openprice) / openprice
                    Else
                        ws.Cells(sumrow, 11).Value = 0
                    End If
                    
                    If ws.Cells(sumrow, 10).Value >= 0 Then
                        ws.Cells(sumrow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(sumrow, 10).Interior.ColorIndex = 3
                    End If
                        
                    sumrow = sumrow + 1
                    
                    openprice = ws.Cells(i + 1, 3).Value
                End If
        Next i
        
        ws.Range("K2:K" & sumrow).NumberFormat = "0.00%"
        
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & sumrow))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & sumrow))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & sumrow))
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'I got this part from https://www.wallstreetmojo.com/vba-max/
        ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I2:I" & sumrow), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & sumrow), 0))
        ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I2:I" & sumrow), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & sumrow), 0))
        ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I2:I" & sumrow), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & sumrow), 0))
        
        ws.Columns.AutoFit
    Next ws
End Sub

