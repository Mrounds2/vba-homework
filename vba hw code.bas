Attribute VB_Name = "Module1"
Sub ticker()

'Declare and set worksheet
    Dim ws As Worksheet

'Loop through all stocks for one year
    For Each ws In Worksheets


'Create the column headings
    ws.Range("I1").value = "Ticker"
    ws.Range("J1").value = "Yearly Change"
    ws.Range("K1").value = "Percent Change"
    ws.Range("L1").value = "Total Stock Volume"

    ws.Range("P1").value = "Ticker"
    ws.Range("Q1").value = "Value"
    ws.Range("O2").value = "Greatest % Increase"
    ws.Range("O3").value = "Greatest % Decrease"
    ws.Range("O4").value = "Greatest Total Volume"

'Create loop for ticker column
    Dim ticker As String
    Dim M As Integer
    Dim yearOpen As Double
    Dim yearlyChange As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim tickervalue As Double
    Dim gtvValue As Double
    M = 1
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    tickerIncrease = ws.Range("P2").value
    tickervalue = ws.Range("Q2").value
    gtvValue = ws.Range("q4").value
    
        'loop through ticker column
        For i = 2 To lastrow
            If ticker <> ws.Cells(i, 1).value Then
                
                M = M + 1
                ws.Cells(M, 9).value = ws.Cells(i, 1).value
                ticker = ws.Cells(i, 1).value
                
         'place starting volume in total stock volume
                ws.Cells(M, 12).value = ws.Cells(i, 7).value
                
         'place open stock in new column to later be subtracted from
                ws.Cells(M, 10).value = ws.Cells(i, 3).value
            
                yearOpen = ws.Cells(i, 3).value
                yearlyChange = ws.Cells(M, 10).value
                
             ElseIf Cells(i + 1, 1).value <> ticker Then
                ws.Cells(M, 10).value = ws.Cells(i, 6).value - yearOpen
                                    
                yearlyChange = ws.Cells(M, 10).value
                ws.Cells(M, 11).value = yearlyChange / yearOpen
             
             Else
                ws.Cells(M, 12).value = ws.Cells(M, 12).value + ws.Cells(i, 7).value
                                       
             End If
             
                           
                PercentChange = ws.Cells(M, 11).value
                
             If PercentChange > tickervalue Then
                ws.Range("q2") = PercentChange
                ws.Range("p2") = ws.Cells(M, 9).value
                
             ElseIf PercentChange < tickervalue Then
                ws.Range("q3") = PercentChange
                ws.Range("p3") = ws.Cells(M, 9).value
                                                                               
             End If
                
                             
         Next i
         
         If ws.Cells(M, 12).value > gtvValue Then
                ws.Range("q4").value = ws.Cells(M, 12).value
                ws.Range("p4").value = ws.Cells(M, 9).value
             
         End If
                                  
    Next ws
    
End Sub

            
        
    
