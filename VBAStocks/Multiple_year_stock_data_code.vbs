Sub VBAChallenge1():

'Declare Current worksheet as a variable
Dim Current As Worksheet

'Loop through all of the worksheets in the active workbook
For Each Current In Worksheets

    'Set the variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    Dim FirstTickerRow As Double
    
    'Assign an initial value to total stock volume
    StockVolume = 0
    
    'Keep track of the location for each variable in the Summary Table
    Dim i As Long
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    FirstTickerRow = 2
    
    'Loop through all Stock Market Data
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    For i = 2 To LastRow
    
     'Set the Ticker Value
        Ticker = Cells(i, 1).Value
        
            'Set the Open Value
            If i = FirstTickerRow Then
                OpenValue = Cells(i, 3).Value
                
            End If
                
        'Assign how to calculate yearly change
        YearlyChange = (Cells(i, 6).Value) - OpenValue
        
        'Condition for percent change
        If OpenValue <> 0 Then
    
            'Assign how to calculate percent change
            PercentChange = (YearlyChange) / (OpenValue)
    
            Else
            PercentChange = 0
    
            End If
        
          'Format Percent Change to percentage
          Dim PercentChange2 As String
          PercentChange2 = FormatPercent(PercentChange)
    
        'Add to the Stock Volume total
        StockVolume = StockVolume + Cells(i, 7).Value
    
    'Check if we are within the same Ticker and if not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Print the Ticker to the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
            
        'Print the YearlyChange to the Summary Table
        Range("J" & Summary_Table_Row).Value = YearlyChange
        
            If YearlyChange > 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            ElseIf YearlyChange < 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
        
        'Print the PercentChange to the Summary Table
        Range("K" & Summary_Table_Row).Value = PercentChange2
        
        'Print the Total Stock Volume to the Summary Table
        Range("L" & Summary_Table_Row).Value = StockVolume
                
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
              
        'Reset the Stock Volume
        StockVolume = 0
                
        'Reset the FirstTickerRow
        FirstTickerRow = i + 1
        
        End If
        
        Next i
    
    'Add the title Ticker to Column I
    Cells(1, 9).Value = "Ticker"
        
    'Add the title Yearly Change to Column J
    Cells(1, 10).Value = "Yearly Change"
    
    'Add the title Percent Change to Column K
    Cells(1, 11).Value = "Percent Change"
    
    
    'Add the title Total Stock Volume to Column L
    Cells(1, 12).Value = "Total Stock Volume"

Next

End Sub