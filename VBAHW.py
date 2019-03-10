Sub MultipleStockData()
    
    'set an initial variable for holding the ticker name
    Dim Ticker_Name As String
    
    'set an initial variable for holding the total per ticker
    Dim ValueTotal As Long
    Value_Total = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'loop throug all ticker values
    For i = 2 To 797711
    
    'check if we are still within the ame ticker, if we are not....
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
     
    'set the ticker name
    Ticker_Name = Cells(i, 1).Value
    
    'add to the ticker total
    Value_Total = Value_Total + Cells(i, 7).Value
    
    'Print the ticker name in the summary table
    Range("J" & Summary_Table_Row).Value = Ticker_Name
    
    'print the brand amount to the summary table
    Range("K" & Summary_Table_Row).Value = Value_Total
    
    'add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'reset the value total
    ValueTotal = 0
    
    'if the cell immediately following a row is the same ticker....
    Else
    
    'add to the value total
    Value_Total = Value_Total + Cells(i, 7).Value
    
    
    End If
    
    Next i
    
   

End Sub

