Attribute VB_Name = "Module1"
'This sub will loop through all stocks and get volume

Sub Stock():

'define woksheet as a variable
Dim sheet As Worksheet

'loop through each worksheet in this excel
For Each sheet In ThisWorkbook.Worksheets

'define variables
endRow = Cells(Rows.Count, 1).End(xlUp).Row

'define summary table variables
summaryTableRow = 2
tickerName = ""
totalVolume = 0
openingPrice = Cells(2, 3).Value
closingPrice = 0

'define max variables
maxTickerName = ""
maxTickerValueChange = 0
previousRowChange = 0
currentRowChange = 0
previousRowTicker = ""
currentRowTicker = ""

  For Row = 2 To endRow
  
    If (Cells(Row, 1).Value = Cells(Row + 1, 1).Value) Then
      
      'add volume when current row and next row values match
      totalVolume = totalVolume + Cells(Row, 7).Value
    
    Else
    
      'add volume for the last row for each stock
      totalVolume = totalVolume + Cells(Row, 7).Value
      
      'set summary values to variables
      tickerName = Cells(Row, 1).Value
      closingPrice = Cells(Row, 6).Value
      
      'set values to summary table
      Cells(summaryTableRow, 9).Value = tickerName
      Cells(summaryTableRow, 14).Value = totalVolume
      
      Cells(summaryTableRow, 10).Value = openingPrice
      Cells(summaryTableRow, 11).Value = closingPrice
      Cells(summaryTableRow, 12).Value = closingPrice - openingPrice
      
      'identify max % change
      If (openingPrice > 0) Then
        Cells(summaryTableRow, 13).Value = (closingPrice - openingPrice) / openingPrice
        currentRowChange = (closingPrice - openingPrice) / openingPrice
      Else
        Cells(summaryTableRow, 13).Value = 0
        currentRowChange = 0
      End If
        
      Cells(summaryTableRow, 13).NumberFormat = "0.00%"
      
      'if loop to compare previous and current row
      If (previousRowChange < currentRowChange) Then
         maxTickerValueChange = currentRowChange
         maxTickerName = tickerName
 Else
         maxTickerValueChange = previousRowChange
         maxTickerName = previousRowTicker
      End If
      
      'reset values
      totalVolume = 0
      summaryTableRow = summaryTableRow + 1
      closingPrice = 0
      previousRowChange = maxTickerValueChange
      previousRowTicker = maxTickerName
         
      'set opening price for next ticker
      openingPrice = Cells(Row + 1, 3).Value
      
      'set highest % change
      If (Row = endRow) Then
    
        Cells(2, 17).Value = maxTickerName
        Cells(2, 18).Value = maxTickerValueChange
        Cells(2, 18).NumberFormat = "0.00%"
            
      Else
      
      End If
      
    End If
  
    Next Row
    
Next sheet

End Sub


