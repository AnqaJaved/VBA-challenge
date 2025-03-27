Attribute VB_Name = "Module1"
Sub StockAnalysis()
    ' Loop through all the worksheets in the workbook (each sheet)
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        ' First, clear any previous results before writing the new ones
        ws.Range("I:Q").Clear
        ws.Range("I:Q").Interior.ColorIndex = xlNone
        
        ' Adding the headers for the summary table where the results will be displayed
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Adding the headers for the bonus table to calculate the biggest changes
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Initializing variables for the calculations
        Dim ticker As String
        Dim lastRow As Long
        Dim summaryRow As Long
        Dim openPrice As Double
        Dim closePrice As Double
        Dim quarterlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        Dim firstRow As Long
        
        ' Starting values for variables
        totalVolume = 0
        summaryRow = 2 ' Start writing summary at row 2
        firstRow = 2 ' Set the first row to start checking data
        
        ' Finding the last row of data to ensure we loop through all rows
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Looping through all rows in the sheet to process the stock data
        Dim i As Long
        For i = 2 To lastRow
            ' Checking if the next row is a different ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Store the ticker name in a variable
                ticker = ws.Cells(i, 1).Value
                
                ' Add up the stock volume for this ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Get the opening and closing prices to calculate changes
                openPrice = ws.Cells(firstRow, 3).Value
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate the quarterly change in price
                quarterlyChange = closePrice - openPrice
                
                ' Calculate the percentage change, checking to avoid dividing by zero
                If openPrice <> 0 Then
                    percentChange = quarterlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Now, write the results to the summary table
                ws.Range("I" & summaryRow).Value = ticker
                ws.Range("J" & summaryRow).Value = quarterlyChange
                ws.Range("K" & summaryRow).Value = percentChange
                ws.Range("K" & summaryRow).NumberFormat = "0.00%" ' Formatting percentage change
                ws.Range("L" & summaryRow).Value = totalVolume
                ws.Range("L" & summaryRow).NumberFormat = "#,##0" ' Formatting Total Stock Volume as a regular number
                
                ' Adding color: Green if positive change, Red if negative change
                If quarterlyChange > 0 Then
                    ws.Range("J" & summaryRow).Interior.Color = RGB(0, 255, 0) ' Green for positive
                    ws.Range("K" & summaryRow).Interior.Color = RGB(0, 255, 0) ' Green for positive
                ElseIf quarterlyChange < 0 Then
                    ws.Range("J" & summaryRow).Interior.Color = RGB(255, 0, 0) ' Red for negative
                    ws.Range("K" & summaryRow).Interior.Color = RGB(255, 0, 0) ' Red for negative
                End If
                
                ' Reset the variables for the next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                firstRow = i + 1
            Else
                ' If it's the same ticker, add the volume to the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Finding the last row of the summary table to calculate max values
        Dim lastSummaryRow As Long
        lastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Initialize variables for calculating the max/min values
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim maxIncreaseTicker As String
        Dim maxDecreaseTicker As String
        Dim maxVolumeTicker As String
        
        ' Start by assuming the first row contains the max values
        maxIncrease = ws.Range("K2").Value
        maxDecrease = ws.Range("K2").Value
        maxVolume = ws.Range("L2").Value
        maxIncreaseTicker = ws.Range("I2").Value
        maxDecreaseTicker = ws.Range("I2").Value
        maxVolumeTicker = ws.Range("I2").Value
        
        ' Loop through the summary table to find the greatest values
        For i = 2 To lastSummaryRow
            ' Check for the greatest % increase
            If ws.Range("K" & i).Value > maxIncrease Then
                maxIncrease = ws.Range("K" & i).Value
                maxIncreaseTicker = ws.Range("I" & i).Value
            End If
            
            ' Check for the greatest % decrease
            If ws.Range("K" & i).Value < maxDecrease Then
                maxDecrease = ws.Range("K" & i).Value
                maxDecreaseTicker = ws.Range("I" & i).Value
            End If
            
            ' Check for the greatest total volume
            If ws.Range("L" & i).Value > maxVolume Then
                maxVolume = ws.Range("L" & i).Value
                maxVolumeTicker = ws.Range("I" & i).Value
            End If
        Next i
        
        ' Write the results of the bonus calculations to the bonus table
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("Q2").Value = maxIncrease
        ws.Range("Q2").NumberFormat = "0.00%" ' Formatting as percentage
        
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("Q3").Value = maxDecrease
        ws.Range("Q3").NumberFormat = "0.00%" ' Formatting as percentage
        
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = maxVolume
        ws.Range("Q4").NumberFormat = "0.00E+00" ' Formatting the volume in scientific notation
        
        ' Autofit the columns to make the sheet look nice
        ws.Columns("I:Q").AutoFit
    Next ws
    
    ' Show a message when the analysis is complete
    MsgBox "Analysis complete for all worksheets!"
    
End Sub

