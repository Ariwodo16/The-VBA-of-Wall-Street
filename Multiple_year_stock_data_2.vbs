Attribute VB_Name = "Module2"
Sub greatest_volume()

Dim ws As Worksheet

    'loop through all tests for one sheet
For Each ws In Worksheets
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Dim maxvolume, startvolume As Double
Dim maxvolumevalue As String


lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow

'find greatest volume
    maxvolume = WorksheetFunction.Max(ws.Range("L2:L3001"))
    maxvolumevalue = WorksheetFunction.Match(maxvolume, ws.Range("L2:L3001"), 0)
    ws.Range("P4").Value = ws.Range("I" & maxvolumevalue + 1)
    ws.Range("Q4").Value = maxvolume
    ws.Range("Q4").Style = "Comma [0]"
 
Next i

Next ws


For Each ws In ActiveWorkbook.Worksheets
        ws.Columns.AutoFit
    Next ws
    
End Sub
