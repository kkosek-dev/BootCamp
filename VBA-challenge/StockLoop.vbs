Attribute VB_Name = "Module1"
Sub StockLoop()

'Define Variables
    Dim ws As Worksheet
'Open Worksheet Loop
For Each ws In Worksheets
    Dim I As Double
    Dim x As Double
    Dim lastrow As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim SummaryTableRow As Double
SummaryTableRow = 2
'Opening Value
    Dim ov As Double
'Closing Value
    Dim cv As Double
'Ticker Symbol
    Dim ts As String
'Yearly Change
    Dim yc As Double
'Percent Change
    Dim pc As Double
'Total Stock Volume
    Dim tsv As Double
'Greatest Percent Increase
    Dim gpc As Double
'Greatest Percent Decrease
    Dim lpc As Double
'Greatest Total Volume
    Dim gtsv As Double

'Insert Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("I:I").ColumnWidth = 8
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J:J").ColumnWidth = 12
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K:K").ColumnWidth = 13
        ws.Range("L1").Value = "Total Stock Value"
        ws.Range("L:L").ColumnWidth = 20
'Bonus Chart
        ws.Range("O1").Value = "Ticker Symbol"
        ws.Range("O:O").ColumnWidth = 12
        ws.Range("P1").Value = "Value"
        ws.Range("P:P").ColumnWidth = 15
        ws.Range("N2").Value = "Greatest Percent Increase"
        ws.Range("N3").Value = "Greatest Percent Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("N:N").ColumnWidth = 20
        
'Initiate Opening Value
ov = ws.Cells(2, 3).Value

'Open Row Loop
For I = 2 To lastrow
    If ws.Cells((I + 1), 1).Value <> ws.Cells(I, 1).Value Then
        ts = ws.Cells(I, 1).Value
        cv = ws.Cells(I, 6).Value
        yc = cv - ov
        pc = (yc / ov)
        tsv = tsv + ws.Cells(I, 7).Value
        ws.Range("I" & SummaryTableRow).Value = ts
        ws.Range("J" & SummaryTableRow).Value = yc
            If yc > 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
        ws.Range("K" & SummaryTableRow).Value = pc
        ws.Range("L" & SummaryTableRow).Value = tsv
        SummaryTableRow = SummaryTableRow + 1
        tsv = 0
        ov = ws.Cells(I + 1, 3).Value
    Else
        tsv = tsv + ws.Cells(I, 7).Value
    End If
Next I

'Bonus
gpc = 0
For x = 2 To lastrow
    If ws.Cells(x, 11).Value > gpc Then
        gpc = ws.Cells(x, 11).Value
        ws.Range("P2").Value = gpc
        ws.Range("O2").Value = ws.Cells(x, 9).Value
    End If
Next x

lpc = 0
For x = 2 To lastrow
    If ws.Cells(x, 11).Value < lpc Then
        lpc = ws.Cells(x, 11).Value
        ws.Range("P3").Value = lpc
        ws.Range("O3").Value = ws.Cells(x, 9).Value
    End If
Next x

gtsv = 0
For x = 2 To lastrow
    If ws.Cells(x, 12).Value > gtsv Then
        gtsv = ws.Cells(x, 12).Value
        ws.Range("P4").Value = gtsv
        ws.Range("O4").Value = ws.Cells(x, 9).Value
    End If
Next x
    
Next ws

MsgBox ("Updated!")

End Sub
