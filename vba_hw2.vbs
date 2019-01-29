Sub total_stock_volume()
    
  ' Declare Variables
    Dim TickerName As String
    Dim StockVolume As Double
    Dim TotalVolume As Double
    Dim Summary_Table_Row As Integer
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet

For Each ws In Worksheets
    ws.Activate

        'Last row determination and summary table setup
        Summary_Table_Row = 2
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
     
        ' Add the titles to the Column Header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"

        ' Loop through rows in the column
    For i = 2 To LastRow

        ' Searches for when the value of the next cell is different than current cell
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the ticker name
        TickerName = Cells(i, 1).Value
        
        ' Add to the Total stock volume
        TotalVolume = TotalVolume + Cells(i, 7).Value
     
       ' Print the Ticker Name in the Summary Table
        Range("I" & Summary_Table_Row).Value = TickerName
      
        ' Print the Total to the Summary Table
        Range("J" & Summary_Table_Row).Value = TotalVolume
        
        ' Add row to summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the Total stock volume
        TotalVolume = 0

        ' If the cell immediately following a row is the same ticker name
        Else

      ' Add to the Total stock volume
        TotalVolume = TotalVolume + Cells(i, 7).Value

        End If

    Next i
    
Next ws

End Sub
