Sub AnalyzeStocks()
'Initiating conversion sequence. Setting Variables.

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, summaryRow As Integer
    Dim yearlyOpen As Double, yearlyClose As Double
    Dim yearlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    Dim ticker As String
    Dim lastVal As Long
  
    
    ' Worksheet Loop that I can't believe I forgot about in the beginning
    For Each ws In ThisWorkbook.Worksheets
        
        ' Initialize summary table
        summaryRow = 2

        ' Summary Table Headers
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With

        ' Focus last row
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Focus last val
        lastVal = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        ' Initialize variables
        totalVolume = 0
        
        ' Row Loops that I kept f*cking up on
        For i = 2 To lastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearlyOpen = ws.Cells(i, 3).Value
            End If
            
            ' Calculate total stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Run that last row back
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                yearlyClose = ws.Cells(i, 6).Value
                
                ' Calculate yearly change and percentage change
                yearlyChange = yearlyClose - yearlyOpen
                If yearlyOpen <> 0 Then
                    percentChange = (yearlyChange / yearlyOpen) * 100
                Else
                    percentChange = 0
                End If
                
                ' Fill summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange & "%"
                ws.Cells(summaryRow, 12).Value = totalVolume

                ' Colors, bay bay!
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red
                End If
                
                ' Reset total volume and increment summary row
                totalVolume = 0
                summaryRow = summaryRow + 1
            End If
        Next i
        
    'Bonus Table Headers (this whole section was a mess, only cleaned up with Dustin's, Miggy's, and Tatiana's help)
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    'Markdown Greatest % Increase
    MaxValue = ws.Application.Max(ws.Range("K:K"))
    ws.Range("Q2").Value = MaxValue
    ws.Range("Q2").NumberFormat = "0.00%"
    For k = 2 To lastVal
      If ws.Range("K" & k).Value = MaxValue Then
        ws.Range("P2").Value = ws.Range("I" & k).Value
      End If
    Next k
'Markdown Greatest % Decrease
    MinValue = ws.Application.Min(ws.Range("K:K"))
    ws.Range("Q3").Value = MinValue
    ws.Range("Q3").NumberFormat = "0.00%"
    For k = 2 To lastVal
      If ws.Range("K" & k).Value = MinValue Then
        ws.Range("P3").Value = ws.Range("I" & k).Value
      End If
    Next k
'Markdown Greatest Total Volume
    GTV = ws.Application.Max(ws.Range("L:L"))
    ws.Range("Q4").Value = GTV
    For k = 2 To lastVal
      If ws.Range("L" & k).Value = GTV Then
        ws.Range("P4").Value = ws.Range("I" & k).Value
      End If
    Next k
'Why was this more annoying than the first part of the HW?
'Also, can you tell the little comments are my favorite part of coding?
        
    Next ws
    'Little message box to let me know it's finally done, goddamn.
    MsgBox "Conversion Process Complete."

End Sub
