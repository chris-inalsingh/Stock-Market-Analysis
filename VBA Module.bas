Attribute VB_Name = "Module1"


Sub StockTicker():

    'Declare variables
    Dim tickercode As String
    Dim summaryrow As Integer
    Dim yearlychange As Double
    Dim Datevalue As Integer
    Dim stockvolumetotal As LongLong
    Dim closeprice As Double
    Dim openprice As Double
    
    'Summary Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total stock volume"
    
    
    
    'Variables
    summaryrow = 2
    stockvolumetotal = 0
    closeprice = 0
    openprice = 0
    yearlychange = 0
    
    
    
    'Create For Loop
    For i = 2 To 71266
        
        'Start Conditional
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tickercode = Cells(i, 1).Value
                stockvolumetotal = stockvolumetotal + Cells(i, 7).Value
                        
                 'Input Ticker code and stockvolume total into Summary table
                Cells(summaryrow, 9).Value = tickercode
                Cells(summaryrow, 12).Value = stockvolumetotal
                summaryrow = summaryrow + 1
                stockvolumetotal = 0
                
                closeprice = Cells(i, 6).Value
                yearlychange = closeprice - openprice

                'Nested conditional for percent change
                If openprice <> 0 Then
                    percentchange = (yearlychange / openprice) * 100
                End If
                
        Else
        
        stockvolumetotal = stockvolumetotal + Cells(i, 7).Value
               
                
        End If
        
        
    
        Next i
                
End Sub

        
    
        Next i
                
End Sub
