Attribute VB_Name = "Module11"
Sub Stock_Ticker_Analyze():

'Define Variables
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim volume As Double
Dim summary_row As Integer


'Create summary table headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Set formatting
Range("K:K").NumberFormat = "0.00%"
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"

'For Loop to scan stock ticker symbols

'Initiate summary_row to 2
summary_row = 2

'Initiate Volume
volume = 0

'Set first year_open price
year_open = Cells(2, 3).Value

    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        'Check to see if on same ticker symbol, if not, then print values
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            
            'Set Ticker symbol
            ticker = Cells(i, 1).Value
            Cells(summary_row, 9).Value = ticker
            
            'Set year close and calculate delta and percentage
            year_close = Cells(i, 6).Value
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
                
            Cells(summary_row, 10).Value = year_change
            Cells(summary_row, 11).Value = percent_change
            
            'Set Total volume
            volume = Cells(i, 7).Value + volume
            Cells(summary_row, 12).Value = volume
            
                'Conditional Formatting the result
                If year_change > 0 Then
                    Cells(summary_row, 10).Interior.Color = vbGreen
                Else
                    Cells(summary_row, 10).Interior.Color = vbRed
                End If
            
            'Increase row of summary table
            summary_row = summary_row + 1
            
            'Set next open price and reset volume to zero
            year_open = Cells(i + 1, 3).Value
            volume = 0
            
        Else
        
            'If still on same stock symbol, increase volume
            volume = Cells(i, 7).Value + volume
            
        End If

    Next i
        
        
End Sub

