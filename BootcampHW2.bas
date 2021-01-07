Attribute VB_Name = "BootcampHW2"
Sub SumStocks()

    ' Define Vars
    
    Dim ws As Worksheet
    Dim volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim results_table_row_counter As Integer
    Dim current_ticker As String
    Dim next_ticker As String
    Dim lastrow As Double
    Dim colorRed As Integer
    Dim colorGreen As Integer
    Dim colorYellow As Integer
    
    ' Variables for second table
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim increase_ticker As String
    Dim decrease_ticker As String
    Dim max_volume_ticker As String
    Dim lastrow_second_table As Integer
    
    ' Define colors
    colorRed = 3
    colorGreen = 4
    colorYellow = 6
    
    ' Loop through all the Worksheets
    For Each ws In Worksheets
        
        ' Find last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Fill in Summary Table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        ws.Range("I1:L1").Font.Bold = True
        
        ' Set some initial values
        open_price = ws.Cells(2, 3)
        results_table_row_counter = 2
        volume = 0
        
        ' Loop through all rows
        For i = 2 To lastrow
        
            'Define tickers, volume
            current_ticker = ws.Cells(i, 1).Value
            next_ticker = ws.Cells(i + 1, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            
            ' Check if tickers are not the same
            If current_ticker <> next_ticker Then
                
                ' Calculate YoY change
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                
                'Print Ticker
                ws.Cells(results_table_row_counter, 9).Value = current_ticker
                
                'Print Yearly change, format currency, color cells
                ws.Cells(results_table_row_counter, 10).Value = yearly_change
                ws.Cells(results_table_row_counter, 10).NumberFormat = "$#,##0.00"
                
                If yearly_change < 0 Then
                
                    ws.Cells(results_table_row_counter, 10).Interior.ColorIndex = colorRed
                
                Else
                    
                    ws.Cells(results_table_row_counter, 10).Interior.ColorIndex = colorGreen
                
                End If
                
                ' Print Percent Change, format
                ' If statement to avoid divide by 0 error
                If open_price <> 0 Then
                    
                    ws.Cells(results_table_row_counter, 11).Value = yearly_change / open_price
                    ws.Cells(results_table_row_counter, 11).NumberFormat = "0.00%"
                
                Else
                
                    ws.Cells(results_table_row_counter, 11).Value = "0"
                    ws.Cells(results_table_row_counter, 11).Interior.ColorIndex = colorYellow
                
                End If
                    
                'Print volume
                ws.Cells(results_table_row_counter, 12).Value = volume
                
                'reset and iterate
                results_table_row_counter = results_table_row_counter + 1
                volume = 0
                open_price = ws.Cells(i + 1, 3).Value
                
            End If
            
       Next i

       ' Add headers for second table
       ws.Range("N2").Value = "Greatest YOY Increase"
       ws.Range("N3").Value = "Greatest YOY Decrease"
       ws.Range("N4").Value = "Highest Volume"
       ws.Range("O1").Value = "Ticker"
       ws.Range("P1").Value = "Value"
       
       ws.Range("N2:N4").Font.Bold = True
       ws.Range("O1").Font.Bold = True
       ws.Range("P1").Font.Bold = True
       
       'set last row using last row of summary table
       lastrow_second_table = results_table_row_counter - 1
       
       'Set initial values
       max_increase = ws.Cells(2, 11).Value
       max_decrease = ws.Cells(2, 11).Value
       max_volume = ws.Cells(2, 12).Value
       increase_ticker = ws.Cells(2, 9).Value
       decrease_ticker = ws.Cells(2, 9).Value
       max_volume_ticker = ws.Cells(2, 9).Value
       
       ' Loop through summary table
       For i = 2 To lastrow_second_table
        
            ' Test for max or min YoY changes
            If ws.Cells(i, 11).Value > max_increase Then
                
                max_increase = ws.Cells(i, 11).Value
                increase_ticker = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value < max_decrease Then
            
                max_decrease = ws.Cells(i, 11).Value
                decrease_ticker = ws.Cells(i, 9).Value
                
            End If
            
            'Test for highest volume
            
            If ws.Cells(i, 12).Value > max_volume Then
            
                max_volume = ws.Cells(i, 12).Value
                max_volume_ticker = ws.Cells(i, 9).Value
            
            End If
       
       Next i
             
       ' Print results
       
       ws.Range("O2").Value = increase_ticker
       ws.Range("p2").Value = max_increase
       ws.Range("P2").NumberFormat = "0.00%"
       ws.Range("O3").Value = decrease_ticker
       ws.Range("P3").Value = max_decrease
       ws.Range("P3").NumberFormat = "0.00%"
       ws.Range("O4").Value = max_volume_ticker
       ws.Range("P4").Value = max_volume
       
    Next ws
    
End Sub



























Sub Clear()
Attribute Clear.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    ws.Range("I1:L1000").Value = ""
    ws.Range("I1:L1000").Interior.ColorIndex = 2
    ws.Range("N1:P4").Value = ""
    ws.Range("n1:p4").Interior.ColorIndex = 2
    
    Next ws
    
End Sub
