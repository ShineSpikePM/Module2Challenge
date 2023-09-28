Attribute VB_Name = "Module1"
Sub tickerData()
'this just will not work, i think i needed to reference it in the Cells and Ranges but will give an error.
'its even declared for each variable but still not running throughout the sheet
'ThisWorkBook.Worksheets still didnt fix it
'i think this has something to do with where ws is referenced like ws.Cells and range
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

'declare everything needed as what it's used as for the sheet
Dim tickername As String
Dim tickervolume As Double
Dim tickerRow As Double
Dim openPrice As Double
Dim closePrice As Double
Dim yearChange As Double
Dim percentChange As Double
        
    'set values for volume and row for the loop to go through
    'set value of open price to the beginning of the sheet
    tickervolume = 0
    tickerRow = 2
    openPrice = Cells(2, 3).Value
        
    'put the text for the table to display before the loop starts
    'this is for the table beside the data
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    


        'start the loop from the second row until the last row, using lastRow declared above to go through the sheet
        For Row = 2 To lastRow
            
            'set the condition for the loop to go through as long as Row does not equal the last number it checked
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
        
              'populate the columns with the values of the names, volumes, percentage, and stock volumes
              tickername = ws.Cells(Row, 1).Value
              tickervolume = tickervolume + ws.Cells(Row, 7).Value
              Range("I" & tickerRow).Value = tickername
              Range("L" & tickerRow).Value = tickervolume
              closePrice = ws.Cells(Row, 6).Value
              yearChange = (closePrice - openPrice)
              ws.Range("J" & tickerRow).Value = yearChange
                
               'if statement to check if the price is 0 for it to not mess with the percentChange calculation
                If (openPrice = 0) Then

                    percentChange = 0

                'if it is not a 0 then calculate the percentage change
                Else
                    
                    percentChange = yearChange / openPrice
                
                End If

              'put the value of the percent change in column K, and convert it to a percentage
              Range("K" & tickerRow).Value = percentChange
              Range("K" & tickerRow).NumberFormat = "0.00%"
   
              'reset the loop values to then go to the next ticker
              tickerRow = tickerRow + 1
              tickervolume = 0
              
              'set the value of the open price to the first open of the next ticker
              openPrice = ws.Cells(Row + 1, 3)
            
            Else
                            'ticker volume sets itself to the value of the row its on
                            tickervolume = tickervolume + ws.Cells(Row, 7).Value
            End If
        
        Next Row

        'make it look nice and auto fit the text in the columns
        'entirecolumn.autofit
        'makes the text in the column fit correctly
        ws.Range("J1").EntireColumn.AutoFit
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("L1").EntireColumn.AutoFit
        
        'set the value for the colors in year change
        colorTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'go through the year change and set the colors for red and green
    For c = 2 To colorTable
            If Cells(c, 10).Value > 0 Then
                Cells(c, 10).Interior.ColorIndex = 4
            Else
                Cells(c, 10).Interior.ColorIndex = 3
            End If
    Next c

        'set the text in the sheet for the extra stuff
        'autofit the text to the column for it
        ws.Cells(2, 15).Value = "Greatest Increase"
        ws.Cells(3, 15).Value = "Greatest Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("O1").EntireColumn.AutoFit
        ws.Range("P1").EntireColumn.AutoFit
        ws.Range("Q1").EntireColumn.AutoFit

        'this part might be wrong check back on it
        'loop for the min max and volume total
        For m = 2 To extraTable

            If ws.Cells(m, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & extraTable)) Then
                ws.Cells(2, 16).Value = ws.Cells(m, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(m, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf ws.Cells(m, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & extraTable)) Then
                ws.Cells(3, 16).Value = ws.Cells(m, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(m, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(m, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & extraTable)) Then
                ws.Cells(4, 16).Value = ws.Cells(m, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(m, 12).Value
            
            End If
        
        Next m

Next ws
End Sub

