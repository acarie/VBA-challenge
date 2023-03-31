# VBA-challenge

I would like to start by saying I struggled very badly with this assignment.
I know that I did not complete it in the slightest and I wanted to submit something in order to resubmit in the future. 

I tried to write pseudo code to walk me through the process but I could never understand the full concept of loops and
how to incorporate them with so many different variables in this case. 

My VBA script contains some of that pseudo code in bits and pieces. I just wanted to try to explain my process before submission. 
It would not let me save the file and upload it as it says it is too big. I am going to copy my code here. Not sure if that suffices but I don't want to miss this assignment. 

 Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' find last row
    
    lastrow = Cells(Rows.Count, 1).End(x1Up).Row

    Dim tickerRow As Integer
    tickerRow = 2
    

    ' declare vol counter variable
    
    Dim volCounter As Integer
    
    ' declare year variable
    
    Dim Year As Integer

    
    ' declare openPrice, closePrice, yearlyChange, percentChange,variables
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim tickerSymbol As String
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerSymbol = Cells(i, 1).Value
            tickerVOL = 0
            tickerVOL = tickerVOL + Cells(i, 7).Value
        
        Else
        
            tickerVOL = tickerVOL
        
        End If
    Next i

End Sub
        

    
    
        
            ' calculate yearlyChange and insert into column J
            ''Insert into column J using the code below
            
            Cells(tickerRow, 10).Value = yearlyChange
            
            
            ' if yearlyChange is positive, color the yearlyChange cell green
            
            
            ' if yearlyChange is negative, color the yearlyChange cell red
            
        
            ' save the ticker symbol - insert into column I
            
            
            ' insert volCounter into column L
            
            
            ' set volumn counter to 0
            
            ' calculate % increase = yearlyChange / Original Number * 100.
            ''format perChg as percentage
            ''and insert % increase to column K
            
   This is for the code to work on all worksheets:
   
   Sub stockticker_ws()

For Each ws In Worksheets

        Dim tickerSymbol As String
        Dim tickerVOL As Double
        
        tickerVOL = 0
        
        Dim tickerRow As Integer
        tickerRow = 2
        
    lastrow = ws.Cells(Rows, Count, 1).End(x1Up).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
End Sub
