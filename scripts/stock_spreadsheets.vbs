Sub stock_ticker()
    ' sub author: Andrew Ryan McKinney - 11/9/2019; Revision 0
    ' this sub creates output (out_arr) columns right of the input stock data and the corresponding output per ticker. The sub also returns the ticker with the maximum in each of the output's categories.
    ' required data inputs are in order across row 1, col A-G: ticker, date, open, high, low, close, vol



    ' dimming variables
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As LongLong
    
    Dim row_index As Long
    Dim row_index_out As Long
    Dim head_index As Integer
    Dim out_arr() As Variant
    
    Dim year_open_value As Double
    Dim year_close_value As Double
    
    Dim greatest_percent_increase As Double
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_decrease As Double
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_total_volume As LongLong
    Dim greatest_total_volume_ticker As String
    
    
    
    
    ' multi-worksheet loop
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' checking for output columns if previously made, then creating output columns if not
        head_index = 8
        out_arr = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        If Cells(1, head_index).Value = "" Then
            For Each op In out_arr
                Cells(1, head_index).Value = op
                head_index = head_index + 1
            Next
        End If
    
        ' sorting by date and then ticker for clean data
        Range("A:G").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
        Range("A:G").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
        
        
        ' reiterate through all of column 1 until empty
        row_index = 2
        row_index_out = 2
        Do Until Cells(row_index, 1).Value = ""
    
            ' checking if current row ticker is same as previous
            If Cells(row_index, 1).Value <> Cells(row_index - 1, 1).Value Then
    
                Cells(row_index_out, 8).Value = Cells(row_index, 1).Value
                year_open_value = Cells(row_index, 3).Value
                total_stock_volume = 0
    
                Do Until Cells(row_index, 1).Value <> Cells(row_index + 1, 1).Value
    
                    total_stock_volume = total_stock_volume + Cells(row_index, 7).Value
                    row_index = row_index + 1
    
                Loop
    
                year_close_value = Cells(row_index, 6).Value
    

                ' outputting standard board
    
                    ' calculating outputs
                    yearly_change = year_close_value - year_open_value
                    
                    If year_open_value = 0 Then
                        percent_change = 0
                    Else
                        percent_change = yearly_change / year_open_value
                    End If
                    
                    
                    ' checking for leaderboard update
                    If percent_change > greatest_percent_increase Then
                    
                        greatest_percent_increase = percent_change
                        greatest_percent_increase_ticker = Cells(row_index, 1).Value
                        
                    ElseIf percent_change < greatest_percent_decrease Then
                    
                        greatest_percent_decrease = percent_change
                        greatest_percent_decrease_ticker = Cells(row_index, 1).Value
                    
                    ElseIf total_stock_volume > greatest_total_volume Then
                    
                        greatest_total_volume = total_stock_volume
                        greatest_total_volume_ticker = Cells(row_index, 1).Value
                    
                    End If
        
                    ' outputting standard output board
                    Cells(row_index_out, 9).Value = yearly_change
                    Cells(row_index_out, 10).Value = FormatPercent(percent_change, 2)
                    Cells(row_index_out, 11).Value = total_stock_volume
                    
                    ' assigning interior color for yearly change
                    If yearly_change > 0 Then
                        Cells(row_index_out, 9).Interior.ColorIndex = 4
                    ElseIf yearly_change < 0 Then
                        Cells(row_index_out, 9).Interior.ColorIndex = 3
                    End If
                
    
                row_index_out = row_index_out + 1
    
    
            End If
    
            row_index = row_index + 1
    
        Loop
        
        ' outputting learderboard
        
            ' leaderboard header & rows
            Cells(1, 15).Value = "Ticker"
            Cells(1, 16).Value = "Value"
            Cells(2, 14).Value = "Greatest % Increase"
            Cells(3, 14).Value = "Greatest % Decrease"
            Cells(4, 14).Value = "Greatest Total Volume"
            
            ' leader board data
            Cells(2, 15).Value = greatest_percent_increase_ticker
            Cells(3, 15).Value = greatest_percent_decrease_ticker
            Cells(4, 15).Value = greatest_total_volume_ticker
            Cells(2, 16).Value = FormatPercent(greatest_percent_increase, 2)
            Cells(3, 16).Value = FormatPercent(greatest_percent_decrease, 2)
            Cells(4, 16).Value = Format(greatest_total_volume, "scientific")
        
        ' autofitting columns in current sheet
        ThisWorkbook.ActiveSheet.Cells.EntireColumn.AutoFit
        
        ' cleaning leaderboard for next sheet
        greatest_percent_increase_ticker = ""
        greatest_percent_decrease_ticker = ""
        greatest_total_volume_ticker = ""
        greatest_percent_increase = 0
        greatest_percent_decrease = 0
        greatest_total_volume = 0
    
    
    Next ws

End Sub