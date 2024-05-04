Sub StockCounterHomework()
        
    ' --------------------------------------------
    ' Setting my Variables
    ' --------------------------------------------
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ' Set an initial variable for holding the stock
        Dim Stock As String
        
        ' Set an initial variable for holding the Total Stock Volume, Qtrly Change, and Qtrly Percent Change per Stock
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        Dim Qtrly_Change As Double
        Qtrly_Change = 0
        
        Dim Qtrly_Percent_Change As Double
        Qtrly_Percent_Change = 0
        
        'Keep track of the location for each stock ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Find how many rows there are
        Dim LastRow As Double
        LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        
        Dim qtrlyOpenVal As Double
        qtrlyOpenVal = ws.Cells(2, 3)
        
        Dim qtrlyCloseVal As Double
        qtrlyCloseVal = 0
        
        Dim minPercent As Double, maxPercent As Double, maxVolume As Double, minTicker As String, maxTicker As String, volumeTicker As String
        minPercent = 0
        maxPercent = 0
        maxVolume = 0
        ' --------------------------------------------
        ' Calculate the Total Stock Volume
        ' --------------------------------------------
        
        ' Loop through all stocks
        For i = 2 To LastRow
        
         ' Check if we are still within the same stock, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Stock
          Stock = ws.Cells(i, 1).Value
    
          ' Add to the Stock Volume Total
          Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                      
          ' Print the Stock in the Summary Table
          ws.Range("I" & Summary_Table_Row).Value = Stock
    
          ' Print the Brand Amount to the Summary Table
          ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
          
          'Calculate the Quarterly Change Value
          qtrlyCloseVal = ws.Cells(i, 6).Value
          Qtrly_Change = qtrlyCloseVal - qtrlyOpenVal
          Qtrly_Percent_Change = Qtrly_Change / qtrlyOpenVal
          Qtrly_Percent_Change_Value = FormatPercent(Qtrly_Percent_Change)
                      
          qtrlyOpenVal = ws.Cells(i + 1, 3).Value
                
          ' Print the Quarterly Change Value in the Summary Table and Apply Formatting
          ws.Range("J" & Summary_Table_Row).Value = Qtrly_Change
          
          If (Qtrly_Change > 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
          ElseIf (Qtrly_Change < 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
          End If
            
          'Finding Greatest % Increase Value & Ticker
          If (Qtrly_Percent_Change > maxPercent) Then
            maxPercent = Qtrly_Percent_Change
            maxTicker = Stock
          End If
          
          If (Qtrly_Percent_Change < minPercent) Then
            minPercent = Qtrly_Percent_Change
            minTicker = Stock
          End If
          
          If (Total_Stock_Volume > maxVolume) Then
            maxVolume = Total_Stock_Volume
            volumeTicker = Stock
          End If
        
          'Print the Quarterly % Change Value in the Summary Table
          ws.Range("K" & Summary_Table_Row).Value = Qtrly_Percent_Change_Value
          
          ' Reset the Total Stock Volume
          Total_Stock_Volume = 0
          
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
        ' If the cell immediately following a row is the same stock...
        Else
    
          ' Add to the Stock Volume Total
          Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
        End If
    
      Next i
        
      'Adding Headers to my data
       ws.Range("I1").Value = "Ticker"
       ws.Range("J1").Value = "Quarterly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Stock Volume"
       ws.Range("P1").Value = "Ticker"
       ws.Range("Q1").Value = "Value"
       ws.Range("O2").Value = "Greatest % Increase"
       ws.Range("O3").Value = "Greatest % Decrease"
       ws.Range("O4").Value = "Greatest Total Volume"
       ws.Range("P2").Value = maxTicker
       ws.Range("P3").Value = minTicker
       ws.Range("P4").Value = volumeTicker
       ws.Range("Q2").Value = FormatPercent(maxPercent)
       ws.Range("Q3").Value = FormatPercent(minPercent)
       ws.Range("Q4").Value = maxVolume
      'Autofit to display data
       ws.Columns("A:Q").AutoFit

   Next ws
   
End Sub
