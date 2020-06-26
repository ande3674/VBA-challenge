Sub Stocks()
  
  ' counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' row counter for stock data
  Dim rowCounter As Integer
  rowCounter = 2
  
  ' Set headers
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
  
  ' Set current Ticker values
  Dim currentTicker As String
  currentTicker = Cells(2, 1).Value
  
  ' Set current opening price at beginning of year
  Dim currenOpeningPrice As Double
  currentOpeningPrice = Cells(2, 3).Value
  
  ' Set a volume counter to zero
  Dim volumeSum As Double
  volumeSum = 0
 
  ' Loop through each row
  For i = 2 To lastrow
    ' While the Ticker remains constant, record first opening and last closing price
    ' Keep a running sum of volume
    volumeSum = volumeSum + Cells(i, 7).Value

    ' When Ticker changes, write/calculate data, reset current Values
    If Cells(i + 1, 1).Value <> currentTicker Then
        ' Get closing price
        Dim closingPrice As Double
        closingPrice = Cells(i, 6).Value
        
        ' Write a result row
        Cells(i, 1).Copy Cells(rowCounter, 9)
        Cells(rowCounter, 10).Value = closingPrice - currentOpeningPrice
        Cells(rowCounter, 11).Value = (closingPrice - currentOpeningPrice) / currentOpeningPrice
        Cells(rowCounter, 12).Value = volumeSum

        ' Color Percent Change Cells
        If Cells(rowCounter, 11).Value < 0 Then Cells(rowCounter, 11).Interior.Color = vbRed
        If Cells(rowCounter, 11).Value >= 0 Then Cells(rowCounter, 11).Interior.Color = vbGreen
        
        ' Reset variables
        rowCounter = rowCounter + 1
        currentTicker = Cells(i + 1, 1).Value
        currentOpeningPrice = Cells(i + 1, 3).Value
        volumeSum = 0
        
    End If

  Next i

End Sub