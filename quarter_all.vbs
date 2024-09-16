Attribute VB_Name = "QtrAll"
Sub Challenge2()

' Loop through all worksheets in spreadsheet file
' ------------------------------

For Each ws In Worksheets

    ' SUMMARY TABLE
    ' ------------------------------
    
    ' Set variables and clear starting values
    ' ------------------------------
    
    Dim Ticker As String
    
    Dim Daily As Date
    
    Dim OpenValue As Double
        OpenValue = 0
    
    Dim CloseValue As Double
        CloseValue = 0
    
    Dim Vol As Double
        Vol = 0
    
    Dim VolTotal As Double
        VolTotal = 0
    
    Dim QtrChangeNum As Double
        QtrChangeNum = 0
    
    Dim QtrChangePct As String
        QtrChangePct = 0
    
    ' Scan data array in ticker column to set iteration range for summary table
    ' ------------------------------
    
    Dim iStart As Double
        iStart = 2
    
    Dim iEnd As Double
    
        With ActiveSheet
        iEnd = ws.Cells(.Rows.Count, "A").End(xlUp).Row
        End With
    
    ' Set starting row for each ticker in summary table
    ' ------------------------------
    
    Dim SummaryRow As Integer
        SummaryRow = 1
    
    ' Set headers for summary table
    ' ------------------------------
    
    ws.Range("I" & SummaryRow).Value = "Ticker"
            
    ws.Range("J" & SummaryRow).Value = "Quarterly Change"
    
    ws.Range("K" & SummaryRow).Value = "Percent Change"
            
    ws.Range("L" & SummaryRow).Value = "Total Stock Volume"
    
    SummaryRow = SummaryRow + 1
    
    ' Loop through array for all stock data
    ' ------------------------------
    
    For i = iStart To iEnd
        
        If ws.Cells(i - 1, 1) = ws.Cells(i, 1) Then
        
        Else
        
            OpenValue = ws.Cells(i, 3)
        
        End If
        
    ' Iterate to build summary for each ticker
    ' ------------------------------
            
        ' Build summary for ticker at end of quarter
        ' ------------------------------
            
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1)
                          
            CloseValue = ws.Cells(i, 6)
            
            QtrChangeNum = CloseValue - OpenValue
                              
            QtrChangePct = QtrChangeNum / OpenValue
        
            VolTotal = VolTotal + ws.Cells(i, 7)
            
            ws.Range("I" & SummaryRow).Value = Ticker
                           
            ws.Range("J" & SummaryRow).Value = QtrChangeNum
            
            ws.Range("J" & SummaryRow).NumberFormat = "0.00"
                                              
            ws.Range("K" & SummaryRow).Value = QtrChangePct
            
            ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
            
            ws.Range("L" & SummaryRow).Value = VolTotal
           
            ' Color the Quarterly Change cell for change
            ' ------------------------------
        
                If QtrChangeNum > 0 Then
                
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                    
                ElseIf QtrChangeNum < 0 Then
                
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                
                Else
                
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = xlNone
                
                End If
            
            ' Move to next ticker summary and reset all values
            ' ------------------------------
            
            SummaryRow = SummaryRow + 1
         
            OpenValue = 0
            
            CloseValue = 0
            
            QtrChangeNum = 0
            
            QtrChangePct = 0
            
            VolTotal = 0
                    
        Else
                   
            ' Add to VolTotal within same ticker
            ' ------------------------------
                           
            VolTotal = VolTotal + ws.Cells(i, 7)
        
        End If
    
    Next i
        
    ' HIGHLIGHT TABLE
    ' ------------------------------
    
    ' Set variables and clear starting values
    ' ------------------------------
    
    Dim SumTickerHigh As String
    
    Dim SumHigh As Double
           
    Dim SumTickerLow As String
    
    Dim SumLow As Double
    
    Dim SumTickerVol As String
    
    Dim SumVol As Double
    
    ' Set starting row for each ticker in summary table
    ' ------------------------------
    
    Dim HighlightRow As Integer
        HighlightRow = 1
    
    ' Set labels for highlight table
    ' ------------------------------
    
    ws.Range("N" & HighlightRow + 1).Value = "Greatest % Increase"
    
    ws.Range("N" & HighlightRow + 2).Value = "Greatest % Decrease"
    
    ws.Range("N" & HighlightRow + 3).Value = "Greatest Total Volume"
    
    ws.Range("O" & HighlightRow).Value = "Ticker"
    
    ws.Range("P" & HighlightRow).Value = "Value"
    
    ' Scan data array in ticker column to set iteration range for summary table
    ' ------------------------------
    
    Dim jStart As Double
        jStart = 2
    
    Dim jEnd As Double
    
        With ActiveSheet
        jEnd = ws.Cells(.Rows.Count, "I").End(xlUp).Row
        End With
    
    ' Set starting values for highlight table
    ' ------------------------------
    
    SumHigh = ws.Cells(2, 11)
    
    SumLow = ws.Cells(2, 11)
    
    SumVol = ws.Cells(2, 12)
    
    ' Loop through summary array for highlight values
    ' ------------------------------
    
    For j = jStart To (jEnd - 1)
      
        If ws.Cells(j + 1, 11) > SumHigh Then
               
            SumHigh = ws.Cells(j + 1, 11)
                        
            SumTickerHigh = ws.Cells(j + 1, 9)
                        
        ElseIf ws.Cells(j + 1, 11) < SumLow Then
               
            SumLow = ws.Cells(j + 1, 11)
                        
            SumTickerLow = ws.Cells(j + 1, 9)
                        
        ElseIf ws.Cells(j + 1, 12) > SumVol Then
               
            SumVol = ws.Cells(j + 1, 12)
                        
            SumTickerVol = ws.Cells(j + 1, 9)
                            
        End If
        
    Next j
    
    ' Display values for highlight table
    ' ------------------------------
        
        ws.Range("O2").Value = SumTickerHigh
        
        ws.Range("P2").Value = SumHigh
    
        ws.Range("P2").NumberFormat = "0.00%"
        
        ws.Range("O3").Value = SumTickerLow
        
        ws.Range("P3").Value = SumLow
    
        ws.Range("P3").NumberFormat = "0.00%"
        
        ws.Range("O4").Value = SumTickerVol
        
        ws.Range("P4").Value = SumVol
    
        ' Range("P4").NumberFormat = "0.00%"

Next ws

End Sub

