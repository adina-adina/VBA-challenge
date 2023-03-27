Sub Stocks_Adina()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    Dim openprice As Double
    Dim closeprice As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Lastrow As Long
    Dim Pricerow As Long
    Dim ticker As String
    Dim Summary_Table_Row As Long


        
            'label variables
            Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Summary_Table_Row = 2
            Pricerow = 2
            vol = 0
            
        
            'create headers
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Stock Volume"
            ws.Range("P1") = "Ticker"
            ws.Range("Q1") = "Value"
            
            'create row labels
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
            
            
                'Loop through tickers and add total volume
                For i = 2 To Lastrow:
                    
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                        'Find ticker
                        ticker = ws.Cells(i, 1).Value
                        
                        'Find total stock volume
                        vol = vol + ws.Range("G" & i).Value
                        
                        'Print ticker name
                        ws.Range("I" & Summary_Table_Row).Value = ticker
                        
                        'Print total stock volume
                        ws.Range("L" & Summary_Table_Row).Value = vol
                        
                        'Calculate yearly and percent changes
                        
                        openprice = ws.Range("C" & Pricerow).Value
                        closeprice = ws.Range("F" & i).Value
                        Yearly_Change = closeprice - openprice
                            
                            If openprice = 0 Then
                                Percent_Change = 0
                            Else
                                Percent_Change = Yearly_Change / openprice
                                
                            End If
                            
                        'Print yearly and percent changes
                        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        
                        'Conditional formatting
                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                        If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                        
                        
                        'loop through all rows
                        Summary_Table_Row = Summary_Table_Row + 1
                        Pricerow = i + 1
                        
                        'Reset the total stock volume
                        vol = 0
                        
                    'Loop through everything
                    Else
                        vol = vol + ws.Range("G" & i).Value
                    
                    End If
                
                Next i
                
                ''find and print tickers/values for the greatest values
                
                'set initial variables to loop through summary table
                Greatest_Increase = ws.Range("K2").Value
                Greatest_Decrease = ws.Range("K2").Value
                Greatest_Volume = ws.Range("L2").Value
                Lastrow_summary = ws.Cells(Rows.Count, "I").End(xlUp).Row
                
                'Loop through summary table to find greatest values and print results
                For m = 2 To Lastrow_summary:
                    If ws.Range("K" & m + 1).Value > Greatest_Increase Then
                        Greatest_Increase = ws.Range("K" & m + 1).Value
                        ws.Range("P2") = ws.Range("I" & m + 1).Value
                        ws.Range("Q2") = Greatest_Increase
                    ElseIf ws.Range("K" & m + 1).Value < Greatest_Decrease Then
                        Greatest_Decrease = ws.Range("K" & m + 1).Value
                        ws.Range("P3") = ws.Range("I" & m + 1).Value
                        ws.Range("Q3") = Greatest_Decrease
                    ElseIf ws.Range("L" & m + 1).Value > Greatest_Volume Then
                        Greatest_Volume = ws.Range("L" & m + 1).Value
                        ws.Range("P4") = ws.Range("I" & m + 1).Value
                        ws.Range("Q4") = Greatest_Volume
                    End If
                Next m
                    

Next ws

End Sub

