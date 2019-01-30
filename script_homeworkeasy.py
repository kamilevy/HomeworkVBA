Sub homework()



  ' Set an initial variable for ticker
  Dim Ticker As String

  ' Set an initial variable for stock totAL
  Dim Total As Double
  Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
Range("J1").Value = "Ticker"
Range("K1").Value = "Stock Total"
  ' Loop through all ticker

  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker = Cells(i, 1).Value

      ' Add to the Total
      Total = Total + Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker

      ' Print the Brand Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total
      Total = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the  Total
      Total = Total + Cells(i, 7).Value
      
      End If
      Next i
      End Sub
      


