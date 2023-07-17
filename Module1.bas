Attribute VB_Name = "Module1"
Sub stockanalyzer()

'Declare ws and the variable worksheet
Dim ws As Worksheet

'for loop to loop through each worksheet
For Each ws In ThisWorkbook.Worksheets

    'ADDS EACH COLUMN AND CELL STRING TITLES THAT DO NOT CHANGE
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"

    'FORMATS COLUMN J TO MAKE NEGATIVE NUMBERS RED AND POSITIVE NUMBERS GREEN
   ' Set the format range to column J
    Set formatRange = ws.Range("J:J")
    
    ' Clear any existing conditional formatting
    formatRange.FormatConditions.Delete

    ' Use conditional formatting to make negative in format numbers red
    Set condition1 = formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    condition1.Interior.Color = vbRed
    
    ' Use conditional formatting to make positive numbers in format range green
    Set condition2 = formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    condition2.Interior.Color = vbGreen



    'FORMATS COULMN K TO MAKE NUMBERS SHOW AS PERCENTAGE
    ' Set the format range to column K
    Set formatRange = ws.Range("K:K")
    
    ' Clear any existing conditional formatting
    formatRange.FormatConditions.Delete
    
    ' Change Format range to show numbers as percentages
    formatRange.NumberFormat = "0.00%"
    
    
    
    'REMOVES FORMATING FROM COULMN HEADERs
    ws.Range("I1:L1").FormatConditions.Delete
    
    
    
    'READS THROUGH THE DATA AND FILLS NEW COLUMNS TO SHOW EACH TICKER,
    'YEARLY CHANGE, PERCENTAGE CHANGE, AND TOTAL STOCK VOLUME.
    currentReadingRow = 2
    numticker = 0
    CurrentWritingRow = 2
    'While loop that continues down each row while column A is not empty
    Do While ws.Cells(currentReadingRow, 1) <> ""
    'sets ticker value equal to the name of the ticker in current row'
    ticker = ws.Cells(currentReadingRow, 1)
    'Prints the name of the ticker in the currentWritingRow
    ws.Cells(CurrentWritingRow, 9) = ticker
    'Counts the number of reading rows under the current ticker
    numticker = WorksheetFunction.CountIf(ws.Range("A:A"), ticker)
    'Prints in column M the number of rows under the current ticker (Not included in final code
    'Cells(CurrentWritingRow, 14) = numticker
    'Prints in column N the number of the first row under the current ticker
    'Cells(CurrentWritingRow, 15) = currentReadingRow
    
    'Reads the opening price of the current row
    OpeningPrice = ws.Cells(currentReadingRow, 3)
    'Reads the closing price of the last row of the ticker
        'this is determined by adding the first row of the ticker
        'to the number of rows with that ticker and subtracting 1
    ClosingPrice = ws.Cells(currentReadingRow + numticker - 1, 6)
    'Calculates the price change of the ticker from open to close
    PriceChange = ClosingPrice - OpeningPrice
    'Calculates the percentage change of that ticker from open to close
    PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
    'Prints the price change of the ticker in column J of the current writing row
    ws.Cells(CurrentWritingRow, 10) = PriceChange
    'Prints the percentage change of ticker in column K of the current writing row
    ws.Cells(CurrentWritingRow, 11) = PercentChange
    'Sets a range of volume data from the current row to the last row in coulmn G
    CurrentVolDataRange = ("G" & currentReadingRow & ":G" & currentReadingRow + numticker - 1)
    'Sums the values in the Current Vol Data Range
    Sum = WorksheetFunction.Sum(ws.Range(CurrentVolDataRange))
    'Prints the sum of the volume range in column L of the current writing row
    ws.Cells(CurrentWritingRow, 12) = Sum
    
    'updates the current row to the first row of the next ticker
    currentReadingRow = numticker + currentReadingRow
    'updates the current writing row to the next row
    CurrentWritingRow = CurrentWritingRow + 1
    Loop


    'FINDS THE TICKERS WITH THE MAX % INCREASE, MAX % DECREASE, AND MAX TOTAL VOLUME.
    'PRINTS TICKERS AND VALUES IN NEW CELLS
    'Finds Maximum Percentage Change From Coulmn K
    MaxPercentage = Application.WorksheetFunction.Max(ws.Range("K:K"))
    'Prints Maximum Percentage Change in Cell Q2
    ws.Range("Q2") = MaxPercentage
    'Sets format of Q2 to percentage
    ws.Range("Q2").NumberFormat = "0.00%"
    
    'Finds Minimum Percentage Change From Coulmn K
    MinPercentage = Application.WorksheetFunction.Min(ws.Range("K:K"))
    'Prints Minimum Percentage Change in Cell Q2
    ws.Range("Q3") = MinPercentage
    'Sets format of Q3 to percentage
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'Finds Maximum Total Stock Volume From Coulmn L
    MaxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    'Prints Maximum Total Stock Volume in Cell Q4
    ws.Range("Q4") = MaxTotalVolume


    MaxRow = WorksheetFunction.Match(MaxPercentage, ws.Range("K:K"), 0)
    ws.Range("P2") = ws.Cells(MaxRow, 9)
    MinRow = WorksheetFunction.Match(MinPercentage, ws.Range("K:K"), 0)
    ws.Range("P3") = ws.Cells(MinRow, 9)
    MaxVolRow = WorksheetFunction.Match(MaxTotalVolume, ws.Range("L:L"), 0)
    ws.Range("P4") = ws.Cells(MaxVolRow, 9)
    
'goes to next worksheet
Next ws
End Sub
