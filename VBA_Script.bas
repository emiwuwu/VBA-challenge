Attribute VB_Name = "Module6"
Sub StockChecker_Loop()

    Dim ws As Worksheet
    Dim lr As Long
    Dim brand_name As String
    Dim brand_total As Double
    Dim summary_row As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim greatest_total_volume As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total_ticker As String
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim firstopen As Double
    Dim lastclose As Double
    
    For Each ws In Worksheets
        lr = ws.Range("A1").CurrentRegion.Rows.Count
        brand_total = 0
        summary_row = 2
        yearly_change = 0
        percent_change = 0
        greatest_total_volume = 0
        greatest_increase = 0
        greatest_decrease = 0
        
        firstopen = ws.Cells(2, 3).Value
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        For i = 2 To lr
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                brand_name = ws.Cells(i, 1).Value
                ws.Range("I" & summary_row).Value = brand_name
                ws.Range("I1").Value = "Ticker"
                
                brand_total = brand_total + ws.Cells(i, 7).Value
                ws.Range("L" & summary_row).Value = brand_total
                ws.Range("L1").Value = "Total Stock Volume"
                
                If brand_total > greatest_total_volume Then
                    greatest_total_volume = brand_total
                    greatest_total_ticker = ws.Cells(i, 1).Value
                    ws.Range("O4").Value = "Greatest Total Volume"
                    ws.Range("Q4").Value = greatest_total_volume
                    ws.Range("P4").Value = greatest_total_ticker
                End If
    
                lastclose = ws.Cells(i, 6).Value
                yearly_change = lastclose - firstopen
                ws.Range("J" & summary_row).Value = yearly_change
                ws.Range("J1").Value = "Yearly Change"
        
                If ws.Range("J" & summary_row).Value > 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                End If
                    
                percent_change = (yearly_change / firstopen)
                ws.Range("K" & summary_row).Value = percent_change
                ws.Range("K" & summary_row).NumberFormat = "0.00%"
                ws.Range("K1").Value = "Percent Change"
            
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ws.Cells(i, 1).Value
                    ws.Range("O2").Value = "Greatest % Increase"
                    ws.Range("Q2").Value = greatest_increase
                    ws.Range("Q2").NumberFormat = "0.00%"
                    ws.Range("P2").Value = greatest_increase_ticker
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ws.Cells(i, 1).Value
                    ws.Range("O3").Value = "Greatest % Decrease"
                    ws.Range("Q3").Value = greatest_decrease
                    ws.Range("Q3").NumberFormat = "0.00%"
                    ws.Range("P3").Value = greatest_decrease_ticker
                End If
                
                'Apply conditional formatting
                If ws.Range("K" & summary_row).Value > 0 Then
                    ws.Range("K" & summary_row).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & summary_row).Interior.ColorIndex = 3
                End If
        
                'reset
                summary_row = summary_row + 1
                firstopen = ws.Cells(i + 1, 3).Value
                brand_total = 0
                yearly_change = 0
                percent_change = 0
            
            Else
                brand_total = brand_total + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub

