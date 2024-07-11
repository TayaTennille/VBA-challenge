Sub CalculateQuarterlyChangesForAllWorksheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim dateValue As Date
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim stockVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim resultsRow As Long
   
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       
        ' Initialize the row to start writing results
        resultsRow = 2 ' Assuming the results start from row 2
       
        ' Loop through each row of data
        For i = 2 To lastRow ' Assuming data starts from row 2
            ' Get data from the current row
            ticker = ws.Cells(i, 1).Value ' Assuming ticker symbol is in column A
           
            ' Convert the date value to a proper date format
            dateValue = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
           
            openingPrice = ws.Cells(i, 3).Value ' Assuming opening price is in column C
            closingPrice = ws.Cells(i, 6).Value ' Assuming closing price is in column F
            stockVolume = ws.Cells(i, 7).Value ' Assuming volume is in column G
           
            ' Calculate quarterly change and percentage change
            quarterlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentageChange = (quarterlyChange / openingPrice) * 100
            Else
                percentageChange = 0 ' Avoid division by zero
            End If
           
            ' Write the results to the next available row, one column over
            ws.Cells(resultsRow, 10).Value = ticker ' Results start from column I
            ws.Cells(resultsRow, 11).Value = quarterlyChange ' Results start from column K
            ws.Cells(resultsRow, 12).Value = percentageChange ' Results start from column L
            ws.Cells(resultsRow, 13).Value = stockVolume ' Results start from column M
           
            ' Move to the next row for results
            resultsRow = resultsRow + 1
        Next i
    Next ws
End Sub

