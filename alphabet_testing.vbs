Attribute VB_Name = "Module1"
' Create column headings
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.


Sub alphabet_testing()
    'Declare and set worksheet
Dim ws As Worksheet

    'loop through all tests for one sheet
For Each ws In Worksheets
ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Tot. Stock Volume"


' Set a variable to hold the tickers
    Dim ticker As String

' Set variable for stock volume
    Dim volume As Double
    volume = 0
    
' Keep track of the location of the tickers in the summary table
    Dim Summarytablerow As Integer
    Summarytablerow = 2
    
' Count the # of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' Loop through all tickers
    For i = 2 To lastrow

    ' Check if all still within the same ticker, if it isn't,
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Set the ticker
        ticker = ws.Cells(i, 1).Value
        
    ' Add to the volume
      volume = volume + ws.Cells(i, 7).Value

        
    ' print the ticker symbol in the summary table
        ws.Range("I" & Summarytablerow).Value = ticker
        ws.Range("L" & Summarytablerow).Value = volume
    
    ' Add one to the summary table row
      Summarytablerow = Summarytablerow + 1
    
    ' Reset the volume
      volume = 0

    ' If the cell immediately following a row is the same brand...
    Else
    
    ' Add to the volume
      volume = volume + ws.Cells(i, 12).Value

      
    End If
    
    'set variable for yearly change
    Dim Yearly As Double
    Yearly = 0
    
    Yearly
    
        
    Next i
    
Next ws

End Sub


