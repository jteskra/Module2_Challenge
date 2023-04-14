VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockEquation()
    Dim ws As Worksheet
    Dim Yearly_Change As Double
    Dim PercentChange As Double
    Dim Ticker As String
    Dim Close_1 As Double
    Dim Open_1 As Double
    Dim Total As LongLong
    Dim ClosePrice As Double
    Dim OpenPrice As Double
    Dim Summary_Table_Row As Long
    Dim currentString As String
    Dim lastRow As Long
    Dim previousString As String
    Dim currentTotal As LongLong
    Dim endR As Double
    Dim YearlyChange As Double
    Dim tVolume As Long
    Dim GTVolume As Long
    Dim TGTVolume As Long
    

    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"



        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        previousString = ""
        Summary_Table_Row = 2
        OutputR = 2
        GTVolume = 0

        For i = 2 To lastRow + 1
            currentString = ws.Cells(i, 1).Value
            
            If currentString <> previousString Or i = lastRow + 1 Then
            endR = i - 1
            
                ' Populate Ticker in Summary table
                ws.Cells(Summary_Table_Row, 3) = OpenPrice
                ws.Cells(endRow, 6) = ClosePrice
                YearlyChange = ClosePrice - OpenPrice
                
            If OpenPrice <> 0 Then
                PercentChange = (YearlyChange) / OpenPrice * 100
            Else
                PercentChange = 0
            End If
                
                ' MsgBox "This is fun"
                
                tVolume = Application.WorksheetFunction.Sum(ws.Range(w.Cells(Summary_Table_Row, 7), ws.Cells(endRow, 7)))
                
                If tVolume > GVolume Then
                    GTVolume = tVolume
                    TGTVolume = previousString
                End If
                
                wsOutput.Cells(OutputRow, 9).Value = previousString
                wsOutput.Cells(OutputRow, 10).Value = YearlyChange
                wsOutput.Cells(OutputRow, 11).Value = Format(PercentChange) & "%"
                wsOutput.Cells(OuputRow, 12).Value = tVolume
                
                OutputR = OutputR + 1
                
                previousString = currentString
                Summary_Table_Row = i
                End If
                
               
    
                           
                ' Color cell based on Yearly Change
                If YearlyChange > 0 Then
                    ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
                End If
                
                ' Move to next row in Summary table
                Summary_Table_Row = Summary_Table_Row + 1
            
        Next i
    Next ws
End Sub
