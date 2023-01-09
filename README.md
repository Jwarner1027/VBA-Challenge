# VBA-Challenge

Sub MultipleYearStockData()

    'Loop through each worksheet
    For Each ws In Worksheets
        
        'Defining variables
        'First one is for my loop and the other two are to hold my place on the sheet
        Dim i As Long
        Dim summary_table_row As Integer
        summary_table_row = 2
        Dim startrow As Long
        startrow = 2
        
        'Hold the worksheet name
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        'Defining the last row of a worksheet so it knows when to stop the loop
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Defining my variables for the second summary table
        Dim percentChange As Double
        Dim largest_increase As Double
        Dim largest_decrease As Double
        Dim largest_volume As Double
        
        'Defining other variables so it is more clear in my loop what values I have
        Dim start_price As Double
        start_price = ws.Cells(startrow, 3).Value
        Dim end_price As Double
        
        'Writing headers for my summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Begin the for loop
        For i = 2 To LastRow
            
            'Verify the ticker names are the different
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Write in the name of the ticker in my first column of my summary table
            ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
            
            'FInding the yearly change and putting it in second column of summary table
            ws.Cells(summary_table_row, 10).Value = ws.Cells(i, 6).Value - start_price
            
                'Changing the color of the cells
                If ws.Cells(summary_table_row, 10).Value < 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                Else
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                End If
                
                'Make sure not dividing by zero then calculate the percent change
                If start_price <> 0 Then
                percentChange = (ws.Cells(i, 6).Value - start_price) / start_price
                ws.Cells(summary_table_row, 11) = Format(percentChange, "Percent")
                Else
                ws.Cells(summary_table_row, 11).Value = Format(0, "Percent")
                End If
                
            'For total volume, find the sum of the range of values
            ws.Cells(summary_table_row, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(startrow, 7), ws.Cells(i, 7)))
            
            'Update my placeholders on the worksheet
            summary_table_row = summary_table_row + 1
            startrow = i + 1
            start_price = ws.Cells(startrow, 3).Value
            
            End If
        
        Next i
    'Define new last row for the summary to find greatest values, start with the very first entry
    Dim Last_summary_row As Integer
    Last_summary_row = ws.Cells(Rows.Count, 9).End(xlUp).row
    largest_increase = ws.Cells(2, 11).Value
    largest_decrease = ws.Cells(2, 11).Value
    largest_volume = ws.Cells(2, 12).Value
    
        'New loop to go through data in summary table
        For i = 2 To Last_summary_row
            'Find largest increase
            If ws.Cells(i, 11).Value > largest_increase Then
            largest_increase = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else
            largest_increase = largest_increase
            End If
            'Enter new value
            ws.Cells(2, 17).Value = Format(largest_increase, "Percent")
            
            'Find largest decrease
            If ws.Cells(i, 11).Value < largest_decrease Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            largest_decrease = ws.Cells(i, 11).Value
            Else
            largest_decrease = largest_decrease
            End If
            'Enter new value
            ws.Cells(3, 17).Value = Format(largest_decrease, "Percent")
            
            'Find largest_volume
            If ws.Cells(i, 12).Value > largest_volume Then
            largest_volume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
            largest_volume = largest_volume
            End If
            'Enter new value
            ws.Cells(4, 17).Value = largest_volume
            
        Next i
        
    Next ws
            
End Sub
