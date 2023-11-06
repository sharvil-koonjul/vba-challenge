Sub TickerReport()
    
    For Each ws In Worksheets
    
    'All row count'
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Dim
      Dim TickerSymb As String
      Dim TickerVolCount As Double
      Dim Closing As Double
      Dim Opening As Double
      Dim Counter As Integer
      Dim Summary_Table_Row As Integer
            
    'Counters'
      TickerVolCount = 0
      Counter = 0
      Summary_Table_Row = 2
            
    'Constants/Headers'
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change($)"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"
      ws.Cells(1, 16).Value = "Ticker"
      ws.Cells(1, 17).Value = "Value"
      ws.Cells(2, 15).Value = "Greatest % Increase"
      ws.Cells(3, 15).Value = "Greatest % Decrease"
      ws.Cells(4, 15).Value = "Greatest Total Volume"
       
  For i = 2 To LastRow
     If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        Counter = Counter + 1
        TickerVolCount = TickerVolCount + ws.Cells(i, 7).Value
     Else
        TickerVolCount = TickerVolCount + ws.Cells(i, 7).Value
        TickerSymb = ws.Cells(i, 1).Value
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Opening = ws.Cells(i - Counter, 3).Value
        Counter = 0
        Closing = ws.Cells(i, 6).Value
        
        'Table Summary'
        ws.Range("J" & Summary_Table_Row).Value = (Closing - Opening)
        ws.Range("I" & Summary_Table_Row).Value = TickerSymb
        ws.Range("L" & Summary_Table_Row).Value = TickerVolCount
        TickerVolCount = 0
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Conditional Formatting for Column J
     If ws.Range("J" & Summary_Table_Row - 1).Value > 0 Then
        ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
        
     ElseIf ws.Range("J" & Summary_Table_Row - 1).Value < 0 Then
        ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
     End If
        
        'Percent Change while avoiding numbers dividing by zero
     If Opening <> 0 Then
        PercentChange = (Closing - Opening) / Opening
        ws.Range("K" & Summary_Table_Row - 1).Value = FormatPercent(PercentChange)
     Else
        ws.Range("K" & Summary_Table_Row - 1).Value = 0
     End If
     
        'Conditional Formatting for Column K
     If ws.Range("K" & Summary_Table_Row - 1).Value > 0 Then
        ws.Range("K" & Summary_Table_Row - 1).Interior.ColorIndex = 4
     ElseIf ws.Range("K" & Summary_Table_Row - 1).Value < 0 Then
        ws.Range("K" & Summary_Table_Row - 1).Interior.ColorIndex = 3
     End If
     End If
     End If
     
   Next i
   
        'Row Counts for K and L'
        LastRowK = ws.Cells(Rows.Count, "K").End(xlUp).Row
        LastRowL = ws.Cells(Rows.Count, "L").End(xlUp).Row
                            
        'Minimum Values and Maximum Values searches in K and L
        MaxValueK = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowK))
        MinValueK = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowK))
        MaxValueL = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowL))
         
        ws.Range("Q2").Value = FormatPercent(MaxValueK)
        ws.Range("Q3").Value = FormatPercent(MinValueK)
        
        'Find the Column I value of the same row as the Min and Max Values
        ws.Range("P2").Value = ws.Cells(WorksheetFunction.Match(MaxValueK, ws.Range("K2:K" & LastRowK), 0) + 1, "I").Value
        ws.Range("P3").Value = ws.Cells(WorksheetFunction.Match(MinValueK, ws.Range("K2:K" & LastRowK), 0) + 1, "I").Value
        ws.Range("Q4").Value = MaxValueL
        ws.Range("P4").Value = ws.Cells(WorksheetFunction.Match(MaxValueL, ws.Range("L2:L" & LastRowL), 0) + 1, "I").Value
       
  Next ws
        
End Sub


