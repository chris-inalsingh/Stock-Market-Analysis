
Sub Challenge():
'Declare worksheet
    Dim Ws As Worksheet
    
    'Variables
    Dim percentchange As Double
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As LongLong
    Dim tickersymbol As String
    
     'Loop worksheet
    For Each Ws In Worksheets
    
    Ws.Cells(1, 16).Value = "Ticker"
    Ws.Cells(1, 17).Value = "Value"
    Ws.Cells(2, 15).Value = "Greatest % Increase"
    Ws.Cells(3, 15).Value = "Greatest % Decrease"
    Ws.Cells(4, 15).Value = "Greatest stock volume"
    
    'Assign variables to inital values
    greatestincrease = Ws.Cells(2, 11)
    greatestdecrease = Ws.Cells(2, 11)
    greatestvolume = Ws.Cells(2, 12)
    
    
    'Add variable for last row
    Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'For Loop
    For y = 2 To Lastrow
    
    'Conditional to sort for table values
    If greatestincrease < Ws.Cells(y, 11) Then
        greatestincrease = Ws.Cells(y, 11)
        tickersymbol = Ws.Cells(y, 11).Offset(0, -2)
        Ws.Cells(2, 16).Value = tickersymbol
        End If
    If greatestdecrease > Ws.Cells(y, 11) Then
       greatestdecrease = Ws.Cells(y, 11)
       tickersymbol = Ws.Cells(y, 11).Offset(0, -2)
        Ws.Cells(3, 16).Value = tickersymbol
        End If
    If greatestvolume < Ws.Cells(y, 12) Then
        greatestvolume = Ws.Cells(y, 12)
        tickersymbol = Ws.Cells(y, 11).Offset(0, -2)
        Ws.Cells(4, 16).Value = tickersymbol
        End If
    Next y
    
    'Display values in table
    Ws.Cells(2, 17).Value = greatestincrease
    Ws.Cells(3, 17).Value = greatestdecrease
    Ws.Cells(4, 17).Value = greatestvolume
    
    
    'Format percentages
    Ws.Range("Q2:Q3").Style = "Percent"
    Ws.Range("Q2:Q3").NumberFormat = "0.0000%"

    Next Ws
    
End Sub
