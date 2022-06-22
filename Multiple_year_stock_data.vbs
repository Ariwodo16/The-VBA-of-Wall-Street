Attribute VB_Name = "Module1"
' Create column headings
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.


Sub stock_volume()
    'Declare and set worksheet
Dim ws As Worksheet

    'loop through all tests for one sheet
For Each ws In Worksheets
ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Tot. Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


' Set a variable to hold the tickers
    Dim ticker As String, volume As Double, perchange As Double, yrchange As Double
    Dim startprice, endprice As Double
    Dim maxpercent, minpercent, maxvolume As Double
    Dim maxpercentvalue, minpercentvalue, maxvolumevalue As String
    Dim lastrow As Long

' initialize values
    volume = 0
    startprice = 0
    endprice = 0
    perchange = 0
    yrchange = 0
    
' Keep track of the location of the tickers in the summary table
    Dim Summarytablerow As Long
    Summarytablerow = 2
    
    
' Count the # of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'assign open date
    startprice = ws.Cells(2, 3).Value
    
' Loop through all tickers
    For i = 2 To lastrow

    ' Check if all still within the same ticker, if it isn't,
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Set the ticker
        ticker = ws.Cells(i, 1).Value
        
    'Find the end prices & change
    endprice = ws.Cells(i, 6).Value
    yrchange = endprice - startprice
    
    'Set condition for percent change
    If startprice <> 0 Then
        perchange = (endprice - startprice) / startprice
        ws.Range("K" & Summarytablerow).Style = "percent"
        
     End If
     
     
    ' Add to the volume
      volume = volume + ws.Cells(i, 7).Value
    
        
    ' print the ticker symbol in the summary table
        ws.Range("I" & Summarytablerow).Value = ticker
        ws.Range("J" & Summarytablerow).Value = yrchange
    
   'conditional format the positive and negative change
    If yrchange > 0 Then
     ws.Range("J" & Summarytablerow).Interior.ColorIndex = 4
     
     Else
     ws.Range("J" & Summarytablerow).Interior.ColorIndex = 3
     
     End If
    
    'print percent change in column K
    ws.Range("K" & Summarytablerow).Value = perchange
    
        
    'print total stock volume in column L
    ws.Range("L" & Summarytablerow).Value = volume
    ws.Range("L" & Summarytablerow).Style = "Comma [0]"


    ' Add one to the summary table row
      Summarytablerow = Summarytablerow + 1
      
    ' New start price
      startprice = ws.Cells(i + 1, 3).Value
      
    
    ' Reset the volume
      volume = 0
      perchange = 0
      

    ' If the cell immediately following a row is the same brand...
    Else
    
    ' Add to the volume
      volume = volume + ws.Cells(i, 7).Value
      
    End If
      'find greatest % increase
    maxpercent = WorksheetFunction.Max(ws.Range("K2:K3001"))
    maxpercentvalue = WorksheetFunction.Match(maxpercent, ws.Range("K2:K3001"), 0)
    ws.Range("P2").Value = ws.Range("I" & maxpercentvalue + 1)
    ws.Range("Q2").Value = maxpercent
    ws.Range("Q2:Q3").Style = "percent"
 
     'find greatest & decrease
    minpercent = WorksheetFunction.Min(ws.Range("K2:K3001"))
    minpercentvalue = WorksheetFunction.Match(minpercent, ws.Range("K2:K3001"), 0)
    ws.Range("P3").Value = ws.Range("I" & minpercentvalue + 1)
    ws.Range("Q3").Value = minpercent
    
    Next i
    
Next ws
  
    For Each ws In ActiveWorkbook.Worksheets
        ws.Columns.AutoFit
    Next ws
    
End Sub






