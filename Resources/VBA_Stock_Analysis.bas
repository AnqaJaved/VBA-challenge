Attribute VB_Name = "Module1"
Sub AnalyzeAllSheets()

    ' Declare variables to store data for each row.
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim volumeValue As Variant

    ' This loop will go through all the sheets in the workbook.
    For Each ws In ThisWorkbook.Sheets

        ' Skip the sheet named "Sheet1".
        If ws.Name <> "Sheet1" Then

            ' Find the last row with data in column A (ticker column).
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' This loop goes through each row, starting from row 2 (ignoring the header in row 1).
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value ' Get the ticker symbol from column A
                openPrice = ws.Cells(i, 3).Value ' Get the opening price from column C
                closePrice = ws.Cells(i, 6).Value ' Get the closing price from column F

                ' Get the volume from column G (this is column 7).
                volumeValue = ws.Cells(i, 7).Value

                ' Check if the volume is a valid number and set to 0 if not.
                If IsNumeric(volumeValue) Then
                    volume = volumeValue
                Else
                    volume = 0
                End If

                ' Calculate the quarterly change (close price - open price).
                quarterlyChange = closePrice - openPrice

                ' Calculate the percentage change: (quarterly change / open price) * 100.
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice) * 100
                Else
                    percentageChange = 0 ' Avoid dividing by zero.
                End If

                ' Round the values to 2 decimal places to clean up the output.
                quarterlyChange = Round(quarterlyChange, 2)
                percentageChange = Round(percentageChange, 2)

                ' Print the results to the Immediate Window.
                Debug.Print "Row " & i & ": " & ticker & ", Open: " & openPrice & ", Close: " & closePrice & ", Volume: " & volume & ", Quarterly Change: " & quarterlyChange & ", Percentage Change: " & percentageChange & "%"
            Next i
        End If
    Next ws

End Sub

' **Code to Calculate Greatest Percentage Change, Greatest Volume, and Greatest Decrease**

Sub AnalyzeAllSheetsWithResults_v2()

    ' Declare variables for tracking the greatest values
    Dim greatestPctIncrease As Double
    Dim greatestPctDecrease As Double
    Dim greatestVolume As Double
    Dim tickerForPctIncrease As String
    Dim tickerForPctDecrease As String
    Dim tickerForGreatestVolume As String

    ' Initialize the variables with very low or very high values
    greatestPctIncrease = -1E+30 ' Start with a very small number
    greatestPctDecrease = 1E+30 ' Start with a very large number
    greatestVolume = 0 ' Start with zero

    ' Loop through the "Analysis Results" sheet to find the greatest values
    For i = 2 To resultRow - 1
        ' Check for greatest percentage increase
        If analysisSheet.Cells(i, 6).Value > greatestPctIncrease Then
            greatestPctIncrease = analysisSheet.Cells(i, 6).Value
            tickerForPctIncrease = analysisSheet.Cells(i, 1).Value ' Ticker symbol for greatest % increase

        ' Check for greatest percentage decrease
        ElseIf analysisSheet.Cells(i, 6).Value < greatestPctDecrease Then
            greatestPctDecrease = analysisSheet.Cells(i, 6).Value
            tickerForPctDecrease = analysisSheet.Cells(i, 1).Value ' Ticker symbol for greatest % decrease

        ' Check for greatest total volume
        If analysisSheet.Cells(i, 4).Value > greatestVolume Then
            greatestVolume = analysisSheet.Cells(i, 4).Value
            tickerForGreatestVolume = analysisSheet.Cells(i, 1).Value ' Ticker symbol for greatest volume
        End If
    Next i

    ' Output the results in the Immediate Window
    Debug.Print "Greatest Percentage Increase: " & tickerForPctIncrease & " - " & greatestPctIncrease & "%"
    Debug.Print "Greatest Percentage Decrease: " & tickerForPctDecrease & " - " & greatestPctDecrease & "%"
    Debug.Print "Greatest Total Volume: " & tickerForGreatestVolume & " - " & greatestVolume

End Sub

' **Code to Calculate Greatest Percentage Increase, Greatest Percentage Decrease, and Greatest Total Volume**

Sub AnalyzeAllSheetsWithResults_v3()

    ' Declare variables to store data for each row.
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim volumeValue As Variant
    Dim analysisSheet As Worksheet
    Dim resultRow As Long

    ' Check if "Analysis Results" sheet exists, and delete it if it does
    On Error Resume Next
    Set analysisSheet = ThisWorkbook.Sheets("Analysis Results")
    On Error GoTo 0

    If Not analysisSheet Is Nothing Then
        Application.DisplayAlerts = False
        analysisSheet.Delete
        Application.DisplayAlerts = True
    End If

    ' Create a new worksheet called "Analysis Results"
    Set analysisSheet = ThisWorkbook.Sheets.Add
    analysisSheet.Name = "Analysis Results"

    ' Write the headers in the first row of the new sheet
    analysisSheet.Cells(1, 1).Value = "Ticker"
    analysisSheet.Cells(1, 2).Value = "Open Price"
    analysisSheet.Cells(1, 3).Value = "Close Price"
    analysisSheet.Cells(1, 4).Value = "Volume"
    analysisSheet.Cells(1, 5).Value = "Quarterly Change"
    analysisSheet.Cells(1, 6).Value = "Percentage Change"

    ' Start writing results from row 2
    resultRow = 2

    ' This loop will go through all the sheets in the workbook.
    For Each ws In ThisWorkbook.Sheets

        ' Skip the sheet named "Sheet1"
        If ws.Name <> "Sheet1" Then

            ' Find the last row with data in column A (ticker column)
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Loop through each row starting from row 2 (ignoring the header in row 1)
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value ' Get the ticker symbol from column A
                openPrice = ws.Cells(i, 3).Value ' Get the opening price from column C
                closePrice = ws.Cells(i, 6).Value ' Get the closing price from column F

                ' Get the volume from column G (this is column 7)
                volumeValue = ws.Cells(i, 7).Value

                ' Check if the volume is a number and not blank
                If IsNumeric(volumeValue) Then
                    volume = volumeValue ' If volume is valid, use it
                Else
                    volume = 0 ' If volume is invalid, set to 0
                End If

                ' Calculate quarterly change (close price - open price)
                quarterlyChange = closePrice - openPrice

                ' Calculate percentage change (quarterly change / open price * 100)
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice) * 100
                Else
                    percentageChange = 0 ' If open price is 0, set percentage change to 0
                End If

                ' Write the results in the new sheet
                analysisSheet.Cells(resultRow, 1).Value = ticker
                analysisSheet.Cells(resultRow, 2).Value = openPrice
                analysisSheet.Cells(resultRow, 3).Value = closePrice
                analysisSheet.Cells(resultRow, 4).Value = volume
                analysisSheet.Cells(resultRow, 5).Value = quarterlyChange
                analysisSheet.Cells(resultRow, 6).Value = percentageChange

                ' Move to the next row in the results sheet
                resultRow = resultRow + 1
            Next i ' This Next corresponds to For i = 2 To lastRow

        End If ' This End If corresponds to If ws.Name <> "Sheet1"
    Next ws ' This Next corresponds to For Each ws In ThisWorkbook.Sheets

    ' Now calculate Greatest Percentage Increase, Greatest Percentage Decrease, and Greatest Total Volume
    Dim greatestPctIncrease As Double
    Dim greatestPctDecrease As Double
    Dim greatestVolume As Double
    Dim tickerForPctIncrease As String
    Dim tickerForPctDecrease As String
    Dim tickerForGreatestVolume As String

    ' Initialize the variables with very low values
    greatestPctIncrease = -1E+30 ' Start with a very small number
    greatestPctDecrease = 1E+30 ' Start with a very large number
    greatestVolume = 0 ' Start with zero

    ' Loop through the "Analysis Results" sheet to find greatest values
    For i = 2 To resultRow - 1 ' We are looking for values in the Analysis Results sheet
        ' Check for greatest percentage increase
        If analysisSheet.Cells(i, 6).Value > greatestPctIncrease Then
            greatestPctIncrease = analysisSheet.Cells(i, 6).Value
            tickerForPctIncrease = analysisSheet.Cells(i, 1).Value ' Ticker symbol for the greatest % increase
        End If

        ' Check for greatest percentage decrease
        If analysisSheet.Cells(i, 6).Value < greatestPctDecrease Then
            greatestPctDecrease = analysisSheet.Cells(i, 6).Value
            tickerForPctDecrease = analysisSheet.Cells(i, 1).Value ' Ticker symbol for the greatest % decrease
        End If

        ' Check for greatest total volume
        If analysisSheet.Cells(i, 4).Value > greatestVolume Then
            greatestVolume = analysisSheet.Cells(i, 4).Value
            tickerForGreatestVolume = analysisSheet.Cells(i, 1).Value ' Ticker symbol for the greatest volume
        End If
    Next i ' This Next corresponds to For i = 2 To resultRow - 1

    ' Output the results in the Immediate Window (or to the worksheet if preferred)
    Debug.Print "Greatest Percentage Increase: " & tickerForPctIncrease & " - " & greatestPctIncrease & "%"
    Debug.Print "Greatest Percentage Decrease: " & tickerForPctDecrease & " - " & greatestPctDecrease & "%"
    Debug.Print "Greatest Total Volume: " & tickerForGreatestVolume & " - " & greatestVolume

    ' Output the greatest percentage increase, decrease, and volume to the worksheet.
    With analysisSheet
        .Cells(resultRow, 1).Value = "Greatest Percentage Increase"
        .Cells(resultRow, 2).Value = tickerForPctIncrease
        .Cells(resultRow, 3).Value = greatestPctIncrease & "%"

        resultRow = resultRow + 1

        .Cells(resultRow, 1).Value = "Greatest Percentage Decrease"
        .Cells(resultRow, 2).Value = tickerForPctDecrease
        .Cells(resultRow, 3).Value = greatestPctDecrease & "%"

        resultRow = resultRow + 1

        .Cells(resultRow, 1).Value = "Greatest Total Volume"
        .Cells(resultRow, 2).Value = tickerForGreatestVolume
        .Cells(resultRow, 3).Value = greatestVolume
    End With

End Sub

' Code to Format the "Analysis Results" Sheet

Sub FormatAnalysisResultsSheet()

    Dim analysisSheet As Worksheet

    ' Set the reference to the "Analysis Results" sheet
    Set analysisSheet = ThisWorkbook.Sheets("Analysis Results")

    ' Apply borders to the entire used range
    With analysisSheet.UsedRange.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0 ' Black color
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Auto-adjust column widths
    analysisSheet.Columns("A:F").AutoFit

    ' Highlight the header row
    With analysisSheet.Rows(1).Interior
        .ColorIndex = 36 ' Light yellow
        .Pattern = xlSolid
    End With

    ' Make the text in the header bold
    analysisSheet.Rows(1).Font.Bold = True

End Sub


Sub ApplyConditionalFormatting()

    Dim analysisSheet As Worksheet
    Set analysisSheet = ThisWorkbook.Sheets("Analysis Results")

    ' Apply conditional formatting to the Quarterly Change column (Column E)
    With analysisSheet.Range("E2:E" & analysisSheet.Cells(analysisSheet.Rows.Count, "E").End(xlUp).Row)
        
        ' Clear any existing conditional formats
        .FormatConditions.Delete

        ' Apply a green format for positive values (greater than 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green

        ' Apply a red format for negative values (less than 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red

    End With

End Sub



Sub OutputGreatestResultsToSheet()

    Dim analysisSheet As Worksheet
    Set analysisSheet = ThisWorkbook.Sheets("Analysis Results")
    
    Dim resultRow As Long
    resultRow = analysisSheet.Cells(analysisSheet.Rows.Count, "A").End(xlUp).Row + 1 ' Get the next available row

    ' Output the greatest percentage increase, decrease, and volume to the worksheet.
    analysisSheet.Cells(resultRow, 1).Value = "Greatest Percentage Increase"
    analysisSheet.Cells(resultRow, 2).Value = tickerForPctIncrease
    analysisSheet.Cells(resultRow, 3).Value = greatestPctIncrease & "%"

    resultRow = resultRow + 1

    analysisSheet.Cells(resultRow, 1).Value = "Greatest Percentage Decrease"
    analysisSheet.Cells(resultRow, 2).Value = tickerForPctDecrease
    analysisSheet.Cells(resultRow, 3).Value = greatestPctDecrease & "%"

    resultRow = resultRow + 1

    analysisSheet.Cells(resultRow, 1).Value = "Greatest Total Volume"
    analysisSheet.Cells(resultRow, 2).Value = tickerForGreatestVolume
    analysisSheet.Cells(resultRow, 3).Value = greatestVolume

End Sub


Sub OutputGreatestResultsToSheet_v2()

    Dim analysisSheet As Worksheet
    Set analysisSheet = ThisWorkbook.Sheets("Analysis Results")
    
    ' Variables to track the greatest values
    Dim greatestPctIncrease As Double
    Dim greatestPctDecrease As Double
    Dim greatestVolume As Double
    Dim tickerForPctIncrease As String
    Dim tickerForPctDecrease As String
    Dim tickerForGreatestVolume As String

    ' Initialize variables with extreme values
    greatestPctIncrease = -1E+30 ' Very small number
    greatestPctDecrease = 1E+30 ' Very large number
    greatestVolume = 0 ' Start with zero

    ' Find the last row in the analysis sheet
    Dim lastRow As Long
    lastRow = analysisSheet.Cells(analysisSheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows of the "Analysis Results" sheet
    Dim i As Long
    For i = 2 To lastRow
        ' Check for greatest percentage increase
        If analysisSheet.Cells(i, 6).Value > greatestPctIncrease Then
            greatestPctIncrease = analysisSheet.Cells(i, 6).Value
            tickerForPctIncrease = analysisSheet.Cells(i, 1).Value ' Ticker for greatest % increase
        End If
        
        ' Check for greatest percentage decrease
        If analysisSheet.Cells(i, 6).Value < greatestPctDecrease Then
            greatestPctDecrease = analysisSheet.Cells(i, 6).Value
            tickerForPctDecrease = analysisSheet.Cells(i, 1).Value ' Ticker for greatest % decrease
        End If
        
        ' Check for greatest total volume
        If analysisSheet.Cells(i, 4).Value > greatestVolume Then
            greatestVolume = analysisSheet.Cells(i, 4).Value
            tickerForGreatestVolume = analysisSheet.Cells(i, 1).Value ' Ticker for greatest volume
        End If
    Next i

    ' Find the next available row for output
    Dim resultRow As Long
    resultRow = lastRow + 2 ' One row below the data for the results

    ' Output the results to the worksheet
    analysisSheet.Cells(resultRow, 1).Value = "Greatest Percentage Increase"
    analysisSheet.Cells(resultRow, 2).Value = tickerForPctIncrease
    analysisSheet.Cells(resultRow, 3).Value = greatestPctIncrease & "%"

    resultRow = resultRow + 1

    analysisSheet.Cells(resultRow, 1).Value = "Greatest Percentage Decrease"
    analysisSheet.Cells(resultRow, 2).Value = tickerForPctDecrease
    analysisSheet.Cells(resultRow, 3).Value = greatestPctDecrease & "%"

    resultRow = resultRow + 1

    analysisSheet.Cells(resultRow, 1).Value = "Greatest Total Volume"
    analysisSheet.Cells(resultRow, 2).Value = tickerForGreatestVolume
    analysisSheet.Cells(resultRow, 3).Value = greatestVolume

End Sub

