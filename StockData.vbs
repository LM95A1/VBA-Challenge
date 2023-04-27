Sub StockAnalysis()
    
    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    Dim lastRow As Long
    Dim summaryTableRow As Long
    Dim yearBeginRow As Long
    
    For Each ws In Worksheets(Array("2018", "2019", "2020"))
        
        ' Initialize variables
        ticker = ""
        openingPrice = 0
        closingPrice = 0
        yearlyChange = 0
        percentChange = 0
        totalVolume = 0
        summaryTableRow = 2
        
        ' Find the last row of the data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through the rows of the data
        For i = 2 To lastRow
            
            ' Check if we're on a new ticker
            If ws.Cells(i, 1).Value <> ticker Then
                
                ' Save the previous ticker's summary data
                If ticker <> "" Then
                    ws.Cells(summaryTableRow, 9).Value = ticker
                    ws.Cells(summaryTableRow, 10).Value = yearlyChange
                    ws.Cells(summaryTableRow, 11).Value = percentChange
                    ws.Cells(summaryTableRow, 12).Value = totalVolume
                End If
                
                ' Update the ticker and opening price
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                
                ' Reset the summary data
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
                
                ' Find the row of the beginning of the year
                yearBeginRow = i
                
                ' Move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
                
            End If
            
            ' Add to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if we're on the last row for the ticker
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
                
                ' Update the closing price
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                
            End If
            
        Next i
        
    Next ws
    
End Sub
