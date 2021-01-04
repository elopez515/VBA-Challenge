Sub StockMarketAnalyst()

'Set initial variable for stock
Dim Ticker As String

    For Each Ws In Worksheets
        
        'Describe the data contents of each colum
        Ws.Cells(1, "I").Value = "Ticker"
        Ws.Cells(1, "J").Value = "Yearly Change"
        Ws.Cells(1, "K").Value = "Percent Change"
        Ws.Cells(1, "L").Value = "Total Stock Volume"
        
        'Format column widths to fit contents
        Ws.Columns("A:L").AutoFit
        
        'Set variables for our data
        Dim LastRow As Long
        Dim Volume As Double
        Volume = 0
        Dim Ticker_Opening_Price As Double
        Dim Ticker_Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Previous_Amount As Long
        Previous_Amount = 2
        Dim Percentage_Change As Double
        
        'Keep track of the location for each stock in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        'Determine Last row
        LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all stocks
        For i = 2 To LastRow
            
            'Add sum of total stock volume
            Volume = Volume + Ws.Cells(i, 7).Value
                
            'Determine if we are still on the same stock
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                
                'Set ticker
                Ticker = Ws.Cells(i, 1).Value
                
                'Put Ticker name in the summary table
                Ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                'Put Volume total in the summary table
                Ws.Range("L" & Summary_Table_Row).Value = Volume
                
                'Reset ticker volume
                Volume = 0
                
                'Determine the values for Ticker_Opening_Price, Ticker_Closing_Price,...
                '... and Yearly_Change
                Ticker_Opening_Price = Ws.Range("C" & Previous_Amount)
                Ticker_Closing_Price = Ws.Range("F" & i)
                Yearly_Change = Ticker_Closing_Price - Ticker_Opening_Price
                
                'Place Yearly_Change in the summary table
                Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Format Yearly Change to show green if (+) and red if (-)
                If Ws.Range("J" & Summary_Table_Row).Value > 0 Then
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                    
                    'Calculate the change in stock price as a percentage
                    If Ticker_Opening_Price = 0 Then
                        Percentage_Change = 0
                    Else
                        Ticker_Opening_Price = Ws.Range("C" & Previous_Amount)
                        Percentage_Change = Yearly_Change / Ticker_Opening_Price
                    End If
                    
                    'Configure the Percentage Change in the summary to include...
                        '...a percent value with two deceimal places
                    Ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    Ws.Range("k" & Summary_Table_Row) = Percentage_Change
                    
                    
                    
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                Previous_Amount = i + 1
                End If
                
            Next i
    
    Next Ws
            

End Sub
