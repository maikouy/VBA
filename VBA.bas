Attribute VB_Name = "Module3"
Sub DataStock()


'--------------------------------------------------
'LOOP THROUGH ALL SHEETS
'--------------------------------------------------
    For Each WS In Worksheets
    '---------------------------------------------

         'Headings
        WS.Range("I1") = "Ticker"
        WS.Range("J1") = "Yearly Change"
        WS.Range("K1") = "Percent Change"
        WS.Range("L1") = "Total Stock Volume"
    
    'Add the Variables Values
        Dim Worksheet As String
        Dim Yearly_Change As Double
        Dim Open_price As Double
        Dim Close_price As Double
        Dim Ticker As String
        Dim Per_Change As Double
        Dim Total_Volume As Double
        Line = 2
        Column = 1
        Volume = 0
        Dim i As Long
     
      'State the last row equation
        For i = 2 To WS.Cells(Rows.Count, "A").End(xlUp).Row
        
      'State the If and Then Statement... If the cell letter below does not match that letter..
        If WS.Cells(i, "A") <> WS.Cells(i + 1, "A") Then
        
        'Set ticker Name
            WS.Cells(Line, "I") = WS.Cells(i, "A")
            Line = Line + 1
        
        'Set close & open price to determine Yearly Change
            Close_price = WS.Cells(i, "F")
            Open_price = WS.Cells(i, "C")
            Yearly_Change = Close_price - Open_price
            WS.Cells(Line - 1, "J") = Yearly_Change
            
        'Set percent change
            Per_Change = ((Open_price - Close_price) / Close_price) * 100
            WS.Cells(Line - 1, "K") = Per_Change
            
        'Set Total Stock Volume
           Total_Volume = Total_Volume + WS.Cells(i, 7)
           WS.Cells(Line - 1, "L") = Total_Volume
           Total_Volume = Total_Volume + 1
        
        'If the cell immediately following a row is the same brand..
        Else
        'Add to the Total
            Total_Volume = Total_Volume + WS.Cells(i, 7)
           
        
        End If
    
       Next i

       Next WS
       
        
        
End Sub
