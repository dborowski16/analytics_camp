Attribute VB_Name = "Module12"
Sub Stock_Ticker_Analyze():

'Define Variables
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim volume As Double
Dim summary_row As Integer

'Define variable for multiple worksheets
Dim ws As Worksheet

'Loop to include multiple worksheets
For Each ws In Worksheets

    'Create summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Create increase/decrease/Total Labels & headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Set formatting
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'Initiate summary_row to 2
    summary_row = 2
    
    'Initiate Volume
    volume = 0
    
    'Initiate increase/decrease/total
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim greatest_total_volume As Double
    
    max_increase = 0
    max_decrease = 0
    greatest_total_volume = 0
    
    'Set first year_open price
    year_open = ws.Cells(2, 3).Value
    
        'For Loop to scan stock ticker symbols
        'i=2 because data starts at row 2 to ignore headers
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'Check to see if on same ticker symbol, if not, then print values
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                
                'Set Ticker symbol
                ticker = ws.Cells(i, 1).Value
                ws.Cells(summary_row, 9).Value = ticker
                
                'Set year close and calculate delta and percentage
                year_close = ws.Cells(i, 6).Value
                year_change = year_close - year_open
                    
                    'Checking for divide by zero
                    If year_open = 0 Then
                        If year_close > 0 Then
                            percent_change = 1
                        Else
                            percent_change = 0
                        End If
                    Else
                        percent_change = (year_close - year_open) / year_open
                    End If
                    
                ws.Cells(summary_row, 10).Value = year_change
                ws.Cells(summary_row, 11).Value = percent_change
                
                'Set Total volume
                volume = ws.Cells(i, 7).Value + volume
                ws.Cells(summary_row, 12).Value = volume
                
                    'Conditional Formatting the result
                    If year_change > 0 Then
                        ws.Cells(summary_row, 10).Interior.Color = vbGreen
                    Else
                        ws.Cells(summary_row, 10).Interior.Color = vbRed
                    End If
                
                    'Check for max/min/totals
                    If percent_change > max_increase Then
                        max_increase = percent_change
                        ws.Range("Q2").Value = max_increase
                        ws.Range("P2").Value = ticker
                    End If
                    
                    If percent_change < max_decrease Then
                        max_decrease = percent_change
                        ws.Range("Q3").Value = max_decrease
                        ws.Range("P3").Value = ticker
                    End If
                    
                    If volume > greatest_total_volume Then
                        greatest_total_volume = volume
                        ws.Range("Q4").Value = greatest_total_volume
                        ws.Range("P4").Value = ticker
                    End If
                    
                        
                'Increase row of summary table
                summary_row = summary_row + 1
                
                'Set next open price and reset volume to zero
                year_open = ws.Cells(i + 1, 3).Value
                volume = 0
                
            Else
            
                'If still on same stock symbol, increase volume
                volume = ws.Cells(i, 7).Value + volume
                
            End If
    
        Next i
        
    Next ws
    
End Sub

