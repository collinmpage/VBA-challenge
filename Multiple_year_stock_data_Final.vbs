Sub Multiple_year_stock_data()

For Each ws In Worksheets

    Dim Ticker As String
    
    Dim Yearly_Change As Double
    
    Dim Percent_Change As Double
    'Format the column for 2 decimal places
    ws.Columns("K").NumberFormat = "0.00%"
    
    Dim Stock_Volume As Double
    Stock_Volume = 0

    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2

    Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim CellCounter As Integer
    CellCounter = 0
    
    Dim LastYearlyChangeRow As Double
        LastYearlyChangeRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'Assign Cells Their Headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To LastRow

    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
            Ticker = ws.Cells(i, 1).Value
        
                'to get a value that isn't zero for stocks that did not start until after the start of the year
                If ws.Cells(i - CellCounter, 3).Value = 0 Then
                
                    Do
                        i = i + 1
                        
                    Loop Until ws.Cells(i - CellCounter, 3).Value <> 0
                    
                    'the value will end up being the percent change from the start of the stock
                    Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(i - CellCounter, 3).Value

                    Percent_Change = Yearly_Change / ws.Cells(i - CellCounter, 3).Value
            
                Else
                    
                    Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(i - CellCounter, 3).Value

                    Percent_Change = Yearly_Change / ws.Cells(i - CellCounter, 3).Value
                
                End If
                
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value


            'Set Column values
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume

            Summary_Table_Row = Summary_Table_Row + 1
      
            Stock_Volume = 0

            CellCounter = 0
        Else

      
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
            CellCounter = CellCounter + 1

        End If
       

    Next i
    
    For i = 2 To LastYearlyChangeRow
    
     'Formatting Cell Color for Yearly Change
        If ws.Cells(i, 10).Value < 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(i, 10).Value >= 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        Else
            
            ws.Cells(i, 10).Interior.ColorIndex = 2
            
            
        End If
        
    Next i
    
Next ws
   
 
End Sub
