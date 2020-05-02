Attribute VB_Name = "Module1"
Sub stock():

    ' Worksheet loop
    
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ws.Cells(1, 1) = 1
    
    
        ' what tasks do I need to perform?
    
        '1. Add headers to columns I through L as follows: "Ticker", "Yearly Change", "Percent Change", and "Total Stock Volume"
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
                 
        '2. Loop through data in column A and read the ticker symbol
        
         ' what are the variables?
        
        Dim stock_ticker As String
        Dim stock_open As Double
        Dim stock_close As Double
        Dim stock_volume As Double
        stock_volume = 0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        stock_open = Cells(2, 3).Value
    
        For i = 2 To LastRow
            
            ' If stock_ticker symbol is different...
                    
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                        ' Capture the stock ticker symbol and stock closing price
                            
                        stock_ticker = Cells(i, 1).Value
                        stock_close = Cells(i, 6).Value
                            
                        ' Capture the Yearly Change (last close of the year minus first open of the year)
                            
                        yearly_change = stock_close - stock_open
                        
                        'Capture the percentage change
                        
                        If stock_open <> 0 Then
                            percentage_change = yearly_change / stock_open
                        Else
                            percentage_change = 0
                            
                        End If
                            
                        'Capture the stock_volume
                        
                        stock_volume = stock_volume + Cells(i, 7).Value
                            
                                                  
                        ' Print the stock_ticker, yearly change, percentage change, and total volume in the Summary Table
                            
                        Range("I" & Summary_Table_Row).Value = stock_ticker
                        Range("J" & Summary_Table_Row).Value = yearly_change
                        Range("K" & Summary_Table_Row).Value = percentage_change
                        Range("L" & Summary_Table_Row).Value = stock_volume
                        
                       
                                                    
                        ' Add one row to the summary table
                        
                        Summary_Table_Row = Summary_Table_Row + 1
                            
                        'reset the stock open and stock volume
                        
                        stock_open = Cells(i, 3).Value
                        stock_volume = 0
                     
                Else
                    
                     stock_volume = stock_volume + Cells(i, 7).Value
                     
                End If
                
        Next i
        
         'add conditional formatting to the yearly change cell based on whether the change was positive or negative
                            
        For j = 2 To LastRow:
        
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            Else
                Cells(j, 10).Interior.ColorIndex = 0
            End If
        
        Next j
        
        
        'BONUS: loop through all worksheets again to find the Greatest % Increase, Greastest % Decrease and Greatest Total Volume and print them in another table in N1:P4
        
        'Add headers to columns O and P as follows: "Ticker", "Value"
        
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        ' Add titles to Cells N2:N5 as follows: "Greatest % Increase, Greatest % Decrease, Greatest Total Volume
        
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        
        Dim k As Double
        
        'For k = 2 To LastRow (I can't get this to work yet)
        
            'Cells(k, 11).Value = application.WorksheetFunction.Max(Range("K2" & ":" & "K" & "LastRow"))
            'Range("O2").Value = Cells(k, 9).Value
            'Range("P2").Value = Cells(k, 11).Value
            
            'Cells(k, 11).Value = WorksheetFunction.Min(Range("K2" & ":" & "K" & "LastRow"))
            'Range("O3").Value = Cells(k, 9).Value
            'Range("P3").Value = Cells(k, 11).Value
            
            'Cells(k, 12).Value = WorksheetFunction.Max(Range("L2" & ":" & "L" & "LastRow"))
            'Range("O4").Value = Cells(k, 9).Value
            'Range("P4").Value = Cells(k, 11).Value
            
        'Next k
            
    Next ws
           
End Sub
