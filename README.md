# VBA-challeneg
Sub challenge()
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As String
Dim ws As Worksheet


greatestincrease = Application.WorksheetFunction.Max(ws.Range("d2:d705714"))
ws.Range("M1").Value = (greatestincrease)
greatestdecrease = Application.WorksheetFunction.Min(ws.Range("E2:E705714"))
ws.Range("M2").Value = (greatestdecrease)
greatestvolume = Application.WorksheetFunction.Max(ws.Range("G2:G705714"))
ws.Range("M3").Value = (greatestvolume)





For Each ws In ThisWorkbook.Sheets
Print .debug

ws.Range("h1").Value = "ticker"
ws.Range("I1").Value = "Yearly"
ws.Range("j1").Value = "Percent"
ws.Range("K1").Value = "Total Volume"
ws.Range("L1").Value "Greatest Increase"
ws.Range("L2").Value = "Greatest Decrease"
ws.Range("L3").Value = "Greatest Volume"

For x = 2 To 705714
ws.Cells(x, 8).Value = ws.Cells(x, 1).Value
ws.Cells(x, 9).Value = ws.Cells(x, 3).Value
ws.Cells(x, 10).Value = ws.Cells(x, 5).Value
ws.Cells(x, 11).Value = ws.Cells(x, 7).Value

Next x
Next ws










End Sub
