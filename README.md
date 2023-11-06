# vba-challenge

'Sharvil's VBA Challenge'

This challenge is for Module 2

3 PNG Files - 3 screenshots for the three different worksheets in the Multiple_Year_Stock_Data_ file.

    Screenshot 1 - Year 2018 Worksheet
    Screenshot 2 - Year 2019 Worksheet
    Screenshot 3 - Year 2020 Worksheet

1 XLSM File - "Multiple_year_stock_data_Complete - Sharvil"
    The completed file after the script has run

1 VB File - "VBA Challenge - Sharvil"
1 Bas File - "VBA Challenge - Sharvil" (added for compatibility)
    The VBA script can be found in this file


Below, you will find a detailed description of my VBA script file if required
-----------------------------

Sub TickerReport()
    
    For Each ws In Worksheets
    
    'Count of all the rows in Column A'
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Dim'
      Dim TickerSymb As String
      Dim TickerVolCount As Double
      Dim Closing As Double
      Dim Opening As Double
      Dim Counter As Integer
      Dim Summary_Table_Row As Integer
            
    'Counters'
      TickerVolCount = 0 'This counter will be used to add up the total of the column G values'
      Counter = 0 'This counter will be used to count the number of cells with the same values in column A'
      Summary_Table_Row = 2 'This counter is the row index for the summary table'
            
    'The values below are constant and/or headers that appear on every sheet in the workbook'
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change($)"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"
      ws.Cells(1, 16).Value = "Ticker"
      ws.Cells(1, 17).Value = "Value"
      ws.Cells(2, 15).Value = "Greatest % Increase"
      ws.Cells(3, 15).Value = "Greatest % Decrease"
      ws.Cells(4, 15).Value = "Greatest Total Volume"

     'For Loop begins'  
  For i = 2 To LastRow
     
     ' As long as the cells in column A have the same values, the counter will add 1 and the ticker count will add up the stock volumes until the value changes and then it will restart and repeat this process' 
     
     If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        Counter = Counter + 1
        TickerVolCount = TickerVolCount + ws.Cells(i, 7).Value
     Else
        TickerVolCount = TickerVolCount + ws.Cells(i, 7).Value
        TickerSymb = ws.Cells(i, 1).Value
     
     'Opening will take the first Column C value of a unique ticker symbol while Closing will grab the last value of Column F'
     
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Opening = ws.Cells(i - Counter, 3).Value
        Counter = 0
        Closing = ws.Cells(i, 6).Value
            
     'Create a new table below the headers'
        ws.Range("J" & Summary_Table_Row).Value = (Closing - Opening)
        ws.Range("I" & Summary_Table_Row).Value = TickerSymb
        ws.Range("L" & Summary_Table_Row).Value = TickerVolCount
        TickerVolCount = 0
      
      'Add the next Row'
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Conditional Formatting for Column J - 4 for Green and 3 for Red from the color index'
     If ws.Range("J" & Summary_Table_Row - 1).Value > 0 Then
        ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
        
     ElseIf ws.Range("J" & Summary_Table_Row - 1).Value < 0 Then
        ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
     End If
        
        'Generating Column K - Percent Change while avoiding numbers dividing by zero. Formatting will be in percentage as well'

     If Opening <> 0 Then
        PercentChange = (Closing - Opening) / Opening
        ws.Range("K" & Summary_Table_Row - 1).Value = FormatPercent(PercentChange)
     Else
        ws.Range("K" & Summary_Table_Row - 1).Value = 0
     End If
     
        'Conditional Formatting for Column K - 4 for Green and 3 for Red from the color index'
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
        
        'Populating the max and min values of Column K into cells Q2 and Q3'
        ws.Range("Q2").Value = FormatPercent(MaxValueK)
        ws.Range("Q3").Value = FormatPercent(MinValueK)
        
        'Q4 will be populating the max value of column L. We are also finding the Column I value of the same row as the Min and Max Values and populating them in cells P2 to P4."
        ws.Range("P2").Value = ws.Cells(WorksheetFunction.Match(MaxValueK, ws.Range("K2:K" & LastRowK), 0) + 1, "I").Value
        ws.Range("P3").Value = ws.Cells(WorksheetFunction.Match(MinValueK, ws.Range("K2:K" & LastRowK), 0) + 1, "I").Value
        ws.Range("Q4").Value = MaxValueL
        ws.Range("P4").Value = ws.Cells(WorksheetFunction.Match(MaxValueL, ws.Range("L2:L" & LastRowL), 0) + 1, "I").Value
       
  Next ws
        
End Sub


