Sub AlphabeticalTest()

    'Loop through all sheets
    For Each ws In Worksheets
    
        ' Set an initial variable for holding the ticker symbol
        Dim Ticker_Name As String
        
        ' Set an initial variable for holding the Percent Change
        Dim Ticker_Yearly_Change As Double
        Ticker_Yearly_Change = 0
        
        ' Set an initial variable for holding the Percent Change
        Dim Ticker_Percent_Change As Double
        Ticker_Percent_Change = 0
        
        ' Set an initial variable for holding the Percent Change
        Dim Ticker_Total As Double
        Ticker_Total = 0
        
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'counts the number of rows
        Dim lastrow  As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Open_Price As Double
        Open_Price = ws.Cells(2, 3).Value
        
        Dim first_row As Double
        first_row = 2
        
        'Set the column titles
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
        
        'Loop through all ticker symbols
        For i = 2 To lastrow
        
            ' Check if we are still within the same ticker symbol, if it is not
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the ticker symbol
                Ticker_Name = ws.Cells(i, 1).Value
                                                
                ' Add to the ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
          
                ' Print the ticker symbol in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Ticker_Name
               
                ' Print the ticker amount to the Summary Table
                ws.Range("N" & Summary_Table_Row).Value = Ticker_Total

                If Open_Price = 0 Then
                    Ticker_Yearly_Change = 0
                    Ticker_Percent_Change = 0

                Else
                    Ticker_Yearly_Change = ws.Cells(i, 6).Value - Open_Price
                    Ticker_Percent_Change = (Ticker_Yearly_Change / Open_Price)

                End If

                ' Print the ticker symbol in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Yearly_Change
                                                      
                ' Print the ticker symbol in the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = Ticker_Percent_Change
                ws.Range("M" & Summary_Table_Row).Style = "Percent"
                ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                  
                Open_Price = ws.Cells(i + 1, 3).Value
                               
                ' Reset the ticker Total
                Ticker_Total = 0
      
            ' If the cell immediately following a row is the same ticker...
            Else
                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            
            End If
                                           
            'Coloring of yearly changes
            If ws.Cells(i, 12).Value > 0 Then
             
                ws.Cells(i, 12).Interior.ColorIndex = 4
              
            ElseIf ws.Cells(i, 12).Value < 0 Then
                            
                ws.Cells(i, 12).Interior.ColorIndex = 3
              
            End If
                                           
        Next i

        ws.columns("A:R").AutoFit
        
    Next ws
    
End Sub