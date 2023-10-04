Sub AnalyzeStockData()
    Dim sheetCount As Integer
    Dim currentSheet As Integer
    
    sheetCount = ActiveWorkbook.Worksheets.Count
    
    ' Loop through each worksheet in the workbook
    For currentSheet = 1 To sheetCount
        ' Create headers for analysis
        Worksheets(currentSheet).Activate
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim currentRow, targetRow, startRow, lastRow As Long
        Dim change, percentChange, volumeSum As Double
        change = 0
        percentChange = 0
        volumeSum = 0
        startRow = 2
        targetRow = 2
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through the ticker column to get unique ticker symbols
        For currentRow = 2 To lastRow
            ' Check the ticker name
            If Cells(currentRow, 1).Value <> Cells(currentRow + 1, 1).Value Then
                ' Add the unique ticker symbol to the Ticker column if it doesn't already exist
                Cells(targetRow, 9).Value = Cells(currentRow, 1).Value
                ' Calculate the yearly change, percent change, and total stock volume for each stock
                change = Cells(currentRow, 6).Value - Cells(startRow, 3).Value
                percentChange = FormatPercent(change / Cells(startRow, 3).Value)
                volumeSum = volumeSum + Cells(currentRow, 7).Value
                ' Assign values to appropriate cells
                Cells(targetRow, 10).Value = change
                Cells(targetRow, 11).Value = percentChange
                ' Apply conditional formatting to cells in yearly change and percent change columns
                If Cells(targetRow, 10).Value > 0 Then
                    Range(Cells(targetRow, 10), Cells(targetRow, 11)).Interior.Color = vbGreen
                ElseIf Cells(targetRow, 10).Value < 0 Then
                    Range(Cells(targetRow, 10), Cells(targetRow, 11)).Interior.Color = vbRed
                End If
                Cells(targetRow, 12).Value = volumeSum
                ' Reset the volume sum for the next ticker
                volumeSum = 0
                ' Update the starting row for the next ticker
                startRow = currentRow + 1
                ' Move to the next row for the next ticker
                targetRow = targetRow + 1
            Else
                ' Continue adding to the volume sum when the ticker name is the same
                volumeSum = volumeSum + Cells(currentRow, 7).Value
            End If
        Next currentRow
        
        Dim minPercent, maxPercent, maxVolume As Double
        minPercent = Cells(2, 11).Value
        maxPercent = Cells(2, 11).Value
        maxVolume = 0
        
        ' Loop through the percent change column to find the greatest increase and the greatest decrease
        ' Loop through the total volume column to find the greatest volume
        For currentRow = 2 To targetRow - 1
            If Cells(currentRow, 11).Value < minPercent Then
                minPercent = Cells(currentRow, 11).Value
            End If
            If Cells(currentRow, 11).Value > maxPercent Then
                maxPercent = Cells(currentRow, 11).Value
            End If
            If Cells(currentRow, 12).Value > maxVolume Then
                maxVolume = Cells(currentRow, 12).Value
            End If
        Next currentRow
        
        ' Assign values to appropriate cells
        Cells(2, 17).Value = FormatPercent(maxPercent)
        Cells(3, 17).Value = FormatPercent(minPercent)
        Cells(4, 17).Value = maxVolume
        
        ' Find the ticker names for the greatest increase, greatest decrease, and greatest volume
        ' Compare the values in the Percent Change and Total Stock Volume columns
        For targetRow = 2 To targetRow - 1
            If Cells(2, 17).Value = Cells(targetRow, 11).Value Then
                Cells(2, 16).Value = Cells(targetRow, 9).Value
            ElseIf Cells(3, 17).Value = Cells(targetRow, 11).Value Then
                Cells(3, 16).Value = Cells(targetRow, 9).Value
            ElseIf Cells(4, 17).Value = Cells(targetRow, 12).Value Then
                Cells(4, 16).Value = Cells(targetRow, 9).Value
            End If
        Next targetRow
        
        ' Autofit columns and rows in the worksheet
        Worksheets(currentSheet).Cells.EntireColumn.AutoFit
        Worksheets(currentSheet).Cells.EntireRow.AutoFit
    Next currentSheet
End Sub

