Attribute VB_Name = "Module1"

Sub StockTicker():
    Dim tickercode As String
    Dim summaryrow As Integer
    Dim yearlychange As Double
    Dim Datevalue As Integer
    Dim stockvolumetotal As Long
    
    summaryrow = 2
    stockvolumetotal = 0
    'Create For Loop
    '71266
    For i = 2 To 71266
        
        'Start Conditional
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tickercode = Cells(i, 1).Value
                stockvolumetotal = stockvolumetotal + Cells(i, 7).Value
                        
                Cells(summaryrow, 9).Value = tickercode
                Cells(summaryrow, 12).Value = stockvolumetotal


                summaryrow = summaryrow + 1
                
                stockvolumetotal = 0
        
        Else
        
        stockvolumetotal = stockvolumetotal + Cells(i, 7).Value
               
                
        End If
        
        
    
        Next i
                
End Sub
