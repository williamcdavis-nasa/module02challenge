Attribute VB_Name = "Qtr1"
Sub Challenge2()

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
    iEnd = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

' Set starting row for each ticker in summary table
' ------------------------------

Dim SummaryRow As Integer
    SummaryRow = 1

' Set headers for summary table
' ------------------------------

Range("I" & SummaryRow).Value = "Ticker"
        
Range("J" & SummaryRow).Value = "Quarterly Change"

Range("K" & SummaryRow).Value = "Percent Change"
        
Range("L" & SummaryRow).Value = "Total Stock Volume"

SummaryRow = SummaryRow + 1

' Loop through array for all stock data
' ------------------------------

For i = iStart To iEnd
    
    If Cells(i - 1, 1) = Cells(i, 1) Then
    
    Else
    
        OpenValue = Cells(i, 3)
    
    End If
    
' Iterate to build summary for each ticker
' ------------------------------
        
    ' Build summary for ticker at end of quarter
    ' ------------------------------
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1)
                      
        CloseValue = Cells(i, 6)
        
        QtrChangeNum = CloseValue - OpenValue
                          
        QtrChangePct = QtrChangeNum / OpenValue
    
        VolTotal = VolTotal + Cells(i, 7)
        
        Range("I" & SummaryRow).Value = Ticker
                       
        Range("J" & SummaryRow).Value = QtrChangeNum
        
        Range("J" & SummaryRow).NumberFormat = "0.00"
                                          
        Range("K" & SummaryRow).Value = QtrChangePct
        
        Range("K" & SummaryRow).NumberFormat = "0.00%"
        
        Range("L" & SummaryRow).Value = VolTotal
       
        ' Color the Quarterly Change cell for change
        ' ------------------------------
    
            If QtrChangeNum > 0 Then
            
                Range("J" & SummaryRow).Interior.ColorIndex = 4
                
            ElseIf QtrChangeNum < 0 Then
            
                Range("J" & SummaryRow).Interior.ColorIndex = 3
            
            Else
            
                Range("J" & SummaryRow).Interior.ColorIndex = xlNone
            
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
                       
        VolTotal = VolTotal + Cells(i, 7)
    
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

Range("N" & HighlightRow + 1).Value = "Greatest % Increase"

Range("N" & HighlightRow + 2).Value = "Greatest % Decrease"

Range("N" & HighlightRow + 3).Value = "Greatest Total Volume"

Range("O" & HighlightRow).Value = "Ticker"

Range("P" & HighlightRow).Value = "Value"

' Scan data array in ticker column to set iteration range for summary table
' ------------------------------

Dim jStart As Double
    jStart = 2

Dim jEnd As Double

    With ActiveSheet
    jEnd = .Cells(.Rows.Count, "I").End(xlUp).Row
    End With

' Set starting values for highlight table
' ------------------------------

SumHigh = Cells(2, 11)

SumLow = Cells(2, 11)

SumVol = Cells(2, 12)

' Loop through summary array for highlight values
' ------------------------------

For j = jStart To (jEnd - 1)
  
    If Cells(j + 1, 11) > SumHigh Then
           
        SumHigh = Cells(j + 1, 11)
                    
        SumTickerHigh = Cells(j + 1, 9)
                    
    ElseIf Cells(j + 1, 11) < SumLow Then
           
        SumLow = Cells(j + 1, 11)
                    
        SumTickerLow = Cells(j + 1, 9)
                    
    ElseIf Cells(j + 1, 12) > SumVol Then
           
        SumVol = Cells(j + 1, 12)
                    
        SumTickerVol = Cells(j + 1, 9)
                        
    End If
    
Next j

' Display values for highlight table
' ------------------------------
    
    Range("O2").Value = SumTickerHigh
    
    Range("P2").Value = SumHigh

    Range("P2").NumberFormat = "0.00%"
    
    Range("O3").Value = SumTickerLow
    
    Range("P3").Value = SumLow

    Range("P3").NumberFormat = "0.00%"
    
    Range("O4").Value = SumTickerVol
    
    Range("P4").Value = SumVol

    ' Range("P4").NumberFormat = "0.00%"

End Sub
