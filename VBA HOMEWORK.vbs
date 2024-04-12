Option Explicit
Sub Stock()
'Declare and set worksheet
Dim ws As Worksheet

'Scan through all Worksheet in the file
For Each ws In Worksheets

' Declare all required variables.
    Dim i, Lastrow, count As Long
    Dim sum, yearlyChange, percentMin, percentMax, volMax, closePrice, openPrice As Double
    Dim priceFlag As Boolean
    Dim percentMinTicker, percentMaxTicker, volMaxTicker As String
 
    'Create the column headings
        ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("I:Q").Columns.AutoFit

    ' Initialize variables before for loop.
    Lastrow = Cells(Rows.count, 1).End(xlUp).Row 'LastRow will contain the row number of the last row with data in column 1 of the worksheet.
    count = 2 'Count start from 2 because of the header at row 1
    sum = 0
    priceFlag = True
    percentMin = 1E+99 'Initialize the percentMin and percentMax to largest available value
    percentMax = -1E+99 '1 multiplied by 10 raised to the power of 99 (1X10^99)
    volMax = -1E+99
    For i = 2 To Lastrow 'Count start from 2 because of the header at row 1 till Last available Row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Save unique ticker symbol in column I.
            ws.Cells(count, 9).Value = ws.Cells(i, 1).Value 'copy the value from one cell to another
            
            ' Calculate Yearly Change and save in column J. Also, highlight cell red (negative) or green (positive).
            closePrice = ws.Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            ws.Cells(count, 10).Value = yearlyChange 'Add first value of yearly change on the 2nd Row if 10th Column
            If yearlyChange < 0 Then
                ws.Cells(count, 10).Interior.ColorIndex = 3 '3 is RED color
                ws.Cells(count, 11).Interior.ColorIndex = 3
            ElseIf yearlyChange > 0 Then
                ws.Cells(count, 10).Interior.ColorIndex = 4 '4 is GREEN color
                ws.Cells(count, 11).Interior.ColorIndex = 4
            End If
            ' Calculate percent change and save in column K. Careful when dividing by zero!
            If yearlyChange = 0 Or openPrice = 0 Then
                ws.Cells(count, 11).Value = 0
            Else
                ws.Cells(count, 11).Value = Format(yearlyChange / openPrice, "#.##%")
            End If
            ' Save Total Volume in column L.
            sum = sum + ws.Cells(i, 7).Value
            ws.Cells(count, 12).Value = sum
            ' Find the values for greatest decrease/increase and greatest volume.
            If ws.Cells(count, 11).Value > percentMax Then
                If ws.Cells(count, 11).Value = ".%" Then
                Else
                    percentMax = ws.Cells(count, 11).Value
                    percentMaxTicker = ws.Cells(count, 9).Value
                End If
            ElseIf ws.Cells(count, 11).Value < percentMin Then
                percentMin = ws.Cells(count, 11).Value
                percentMinTicker = ws.Cells(count, 9).Value
            ElseIf ws.Cells(count, 12).Value > volMax Then
                volMax = ws.Cells(count, 12).Value
                volMaxTicker = ws.Cells(count, 9).Value
            End If
            ' Reset variables and go to next ticker symbol.
            count = count + 1 'increment count by 1
            sum = 0 'Reset the value of sum for new sum calculation
            priceFlag = True 'Reset the price flag to True
        Else
            ' Use flag to save the open price value at the start of the year.
            If priceFlag Then
                openPrice = Cells(i, 3).Value
                priceFlag = False
            End If
            ' If adjacent ticker symbols are the same, then save volume value.
            sum = sum + ws.Cells(i, 7).Value
        End If
        
    ' Value (Q2 - Q4).
    ws.Cells(2, 17).Value = Format(percentMax, "#.##%") 'Value of Q2
    ws.Cells(3, 17).Value = Format(percentMin, "#.##%") 'Value of Q3
    ws.Cells(4, 17).Value = volMax 'Value of Q4

    ' Ticker symbol (Q2-Q4)
    ws.Cells(2, 16).Value = percentMaxTicker 'Ticker for Greatest % Increase - Q2
    ws.Cells(3, 16).Value = percentMinTicker 'Ticker for Greatest % Decrease - Q3
    ws.Cells(4, 16).Value = volMaxTicker  'Ticker for Greatest Total Volume - Q3
    
    Next i
    Next ws
End Sub