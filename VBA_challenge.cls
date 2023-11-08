VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()

Dim ws As Worksheet
For Each ws In Worksheets
   
ws.Cells(1, "I") = "Ticker"
ws.Cells(1, "J") = "Yearly Change"
ws.Cells(1, "K") = "Percent Change"
ws.Cells(1, "L") = "Total Stock Volume"
ws.Cells(1, "P") = "Ticker"
ws.Cells(1, "Q") = "Value"
ws.Cells(2, "O") = "Greatest % Increase"
ws.Cells(3, "O") = "Greatest % Decrease"
ws.Cells(4, "O") = "Greatest Total Volume"

Dim ticker As String
Dim start As Double
Dim openprice As Double
Dim closeprice As Double
Dim vol As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvol  As Double
Dim volticker As String
Dim increaseticker As String
Dim decreaseticker As String

openprice = ws.Cells(2, "C")
vol = 0
start = 2

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
ticker = ws.Cells(i, "A")

vol = vol + ws.Cells(i, "G")

If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
ws.Cells(start, "I") = ticker

closeprice = ws.Cells(i, "F")
ws.Cells(start, "J") = closeprice - openprice

    If ws.Cells(start, "J") > 0 Then
    ws.Cells(start, "J").Interior.ColorIndex = 4
    
    ElseIf ws.Cells(start, "J") < 0 Then
    ws.Cells(start, "J").Interior.ColorIndex = 3
    End If
    
    If ws.Cells(start, "K") > 0 Then
    ws.Cells(start, "K").Interior.ColorIndex = 4
    
    ElseIf ws.Cells(start, "K") < 0 Then
    ws.Cells(start, "K").Interior.ColorIndex = 3
    End If

If openprice <> 0 Then
ws.Cells(start, "K") = FormatPercent((closeprice - openprice) / openprice, 2)

Else: ws.Cells(start, "K") = Null
End If

ws.Cells(start, "L") = vol

start = start + 1
openprice = ws.Cells(i + 1, "C")
vol = 0
End If
Next i

greatestincrease = 0
greatestdecrease = 0
greatestvol = 0

For i = 2 To 3000
    If ws.Cells(i, "L") > greatestvol Then
    greatestvol = ws.Cells(i, "L")
    volticker = ws.Cells(i, "I")
    
    End If
    Next i

ws.Cells(4, "Q") = greatestvol
ws.Cells(4, "P") = volticker

For i = 2 To 3000
    If ws.Cells(i, "K") > greatestincrease Then
    greatestincrease = ws.Cells(i, "K")
    increaseticker = ws.Cells(i, "I")
    
    End If
    Next i
     
ws.Cells(2, "Q") = FormatPercent(greatestincrease)
ws.Cells(2, "P") = increaseticker
    
For i = 2 To 3000
    If ws.Cells(i, "K") < greatestdecrease Then
    greatestdecrease = ws.Cells(i, "K")
    decreaseticker = ws.Cells(i, "I")
    
    End If
    Next i
    
ws.Cells(3, "Q") = FormatPercent(greatestdecrease)
ws.Cells(3, "P") = decreaseticker

Next
End Sub
