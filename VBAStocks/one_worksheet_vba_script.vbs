Attribute VB_Name = "Module1"
Sub stocktracker():
    Columns(9).Clear
    Columns(10).Clear
    Columns(11).Clear
    Columns(12).Clear
    Columns(15).Clear
    Columns(16).Clear
    Columns(17).Clear
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total stock volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim i As Long
    Dim j As Integer
    Dim lrow As Long
    Dim sumrow As Long
    Dim ticker As String
    Dim openprice As Double
    Dim closingprice As Double
    
    
    sumrow = 2
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    openprice = Cells(2, 3).Value
    
    For i = 2 To lrow
        Cells(sumrow, 12).Value = Cells(sumrow, 12).Value + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker = Cells(i, 1).Value
                Cells(sumrow, 9).Value = ticker
                
                closingprice = Cells(i, 6).Value
                Cells(sumrow, 10).Value = closingprice - openprice
                If openprice <> 0 Then
                    Cells(sumrow, 11).Value = (closingprice - openprice) / openprice
                Else
                    Cells(sumrow, 11).Value = 0
                End If
                
                If Cells(sumrow, 10).Value >= 0 Then
                    Cells(sumrow, 10).Interior.ColorIndex = 4
                Else
                    Cells(sumrow, 10).Interior.ColorIndex = 3
                End If
                    
                sumrow = sumrow + 1
                
                openprice = Cells(i + 1, 3).Value
            End If
    Next i
    
    Range("K2:K" & sumrow).NumberFormat = "0.00%"
    
    Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & sumrow))
    Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & sumrow))
    Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & sumrow))
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'I got this part from https://www.wallstreetmojo.com/vba-max/
    Range("P2").Value = WorksheetFunction.Index(Range("I2:I" & sumrow), WorksheetFunction.Match(Range("Q2").Value, Range("K2:K" & sumrow), 0))
    Range("P3").Value = WorksheetFunction.Index(Range("I2:I" & sumrow), WorksheetFunction.Match(Range("Q3").Value, Range("K2:K" & sumrow), 0))
    Range("P4").Value = WorksheetFunction.Index(Range("I2:I" & sumrow), WorksheetFunction.Match(Range("Q4").Value, Range("L2:L" & sumrow), 0))
    
    Columns.AutoFit
End Sub
