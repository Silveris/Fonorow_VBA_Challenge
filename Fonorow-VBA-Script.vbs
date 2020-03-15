Sub main()
For Each ws In Worksheets
    ' ticker - 1, open - 3, close - 6, vol - 7'
    '(letter, number) or (Col, Row)


    ' read ticker > save first open > add vol to total
    ' cycle through rows untill find new ticker
    ' go to prev row and save final close
    ' create new ticker row for new ticker and repeat process

    'for each ws is Worksheats < loop through all worksheets


    'global vars here'
    ' ticker vars
    Dim prevTick As String

    Dim ticker As String
    Dim curTick As String
    newTicker = False

    ' vol vars
    Dim volume As Double

    'price change vars
    Dim yrlyChng As Double
    Dim pcntChng As Double
    openFound = False
    Dim fstOpen As Double
    Dim lstClose As Double

    ' navagation vars
    tickerIn = 1
    openIn = 3
    closeIn = 6
    volumeIn = 7

    tickerOut = 9
    yrlyChngOut = 10
    pcntChngOut = 11
    totalStockVolOut = 12

    outputRow = 2


       '-------------------------------------------------------------

    
        ws.Cells(1, tickerOut).Value = "Ticker"
        ws.Cells(1, yrlyChngOut).Value = "Yearly Change  "
        ws.Cells(1, pcntChngOut).Value = "Percent Change  "
        ws.Cells(1, totalStockVolOut).Value = "Total Stock Vol   "
        ws.Cells.EntireColumn.AutoFit
                
        Dim lngth As Long
        lngth = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
      For i = 2 To lngth
    
        ' Check if we are still with the same ticker
        If ws.Cells(i + 1, tickerIn).Value <> ws.Cells(i, tickerIn).Value Then
         
          ' Set the ticker
          ticker = ws.Cells(i, tickerIn).Value
    
          ' Add to the vol Total
          volume = volume + ws.Cells(i, volumeIn).Value
    
          ' Print the Ticker in the Summary Table
          ws.Range("I" & outputRow).Value = ticker
    
          ' Print the vol Amount to the Summary Table
          ws.Range("L" & outputRow).Value = volume
          
          lstClose = ws.Cells(i, closeIn).Value
          yrlyChng = (lstClose - fstOpen)
          pcntChng = (yrlyChng / fstOpen)
                
          ws.Range("J" & outputRow).Value = yrlyChng
          If ws.Range("J" & outputRow).Value < 0 Then
            ws.Range("J" & outputRow).Interior.ColorIndex = 3
          ElseIf ws.Range("J" & outputRow).Value >= 0 Then
            ws.Range("J" & outputRow).Interior.ColorIndex = 4
          End If
          
          
          ws.Range("K" & outputRow).Value = pcntChng
          ws.Range("K" & outputRow).NumberFormat = "0.00%"
    
          ' itterate through output rows
          outputRow = outputRow + 1
          
          ' Reset Total vol
          volume = 0
                
    
        ' If the following cell is the same
        Else
    
          ' Add to the total
          volume = volume + ws.Cells(i, volumeIn).Value
          
          If openFound = False Then
                fstOpen = ws.Cells(i, openIn).Value
                openFound = True
          End If
    
        End If
    
      Next i
Next ws
                    
End Sub

























