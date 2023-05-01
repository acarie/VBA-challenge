Sub Mod2Challenge()

    'Applies code to all worksheets
    For Each ws In Worksheets
    
    'Declare variables needed
    
        
        Dim i As Long
        Dim j As Long
        Dim tixcounter As Long
        Dim lastA As Long
        Dim lastI As Long
        Dim percentchange As Double
        Dim increase As Double
        Dim decrease As Double
        Dim total As Double
        Dim GIT As String
        Dim GDT As String
        Dim GTT As String
        Dim WSName As String
            
        WSName = ws.Name
        
        'Label Columns
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        tixcounter = 2
        
        'Finds last A
        
        lastA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To lastA
            
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    tickerSymbol = ws.Cells(i, 1).Value
                    vol = ws.Cells(i, 7).Value
                    yopen = ws.Cells(i, 3).Value
                     
                     
                     
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    
                    
                    
                   
                    yclose = ws.Cells(i, 6).Value
                    ychange = yclose - yopen
                    PerChange = (yclose - yopen) / yopen
                    vol = vol + ws.Cells(i, 7).Value
                    
                    ws.Cells(tixcounter, 9).Value = tickerSymbol
                    ws.Cells(tixcounter, 10).Value = ychange
                    ws.Cells(tixcounter, 11).Value = PerChange
                    ws.Cells(tixcounter, 12).Value = vol
                    
                    'Assigns color

                    If ychange >= 0 Then
                        ws.Cells(tixcounter, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(tixcounter, 10).Interior.ColorIndex = 3
                    End If
                        
                    tixcounter = tixcounter + 1
                Else
                    vol = vol + ws.Cells(i, 7).Value
                    
                End If
                               
                
            Next i
            
            'Find last I
            lastI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            increase = 0
            decrease = 0
            
            For i = 2 To lastI
                
                If ws.Cells(i, 11).Value > increase Then
                    increase = ws.Cells(i, 11).Value
                    GIT = ws.Cells(i, 9).Value
                    
                End If
            
                If ws.Cells(i, 11).Value < decrease Then
                    decrease = ws.Cells(i, 11).Value
                    GDT = ws.Cells(i, 9).Value
                    
                End If
                
                If ws.Cells(i, 12).Value > total Then
                    total = ws.Cells(i, 12).Value
                    GTT = ws.Cells(i, 9).Value
                    
                End If
            
            Next i
            ws.Cells(2, 14).Value = "Greatest Increase"
            ws.Cells(3, 14).Value = "Greatest Decrease"
            ws.Cells(4, 14).Value = "Greatest Volume"
            ws.Cells(2, 15).Value = GIT
            ws.Cells(3, 15).Value = GDT
            ws.Cells(4, 15).Value = GTT
            ws.Cells(2, 16).Value = GreatInc
            ws.Cells(3, 16).Value = GreatDec
            ws.Cells(4, 16).Value = GreatTot
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            
        Next ws
              
        
    

End Sub

