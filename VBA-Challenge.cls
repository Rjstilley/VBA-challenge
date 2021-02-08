Sub VbaChallenge()
For Each ws In Worksheets
Dim lrow As LongLong

lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 12) = "Total Stock Value"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"

Dim Ticker As String
Dim Count As LongLong
Dim StockValue As LongLong
Dim YC As Double
Dim Opened As Double
Dim Closed As Double
Dim PC As Double
Dim GPI As Double
Dim GPD As Double
Dim GTV As LongLong
Dim GPITicker As String
Dim GPDTicker As String
Dim GTVTicker As String

GPI = 0
GPD = 0
GTV = 0
StockValue = 0
Count = 2
Opened = ws.Cells(2, 3).Value
'GreaterPercent

    For i = 2 To lrow
    Ticker = ws.Cells(i, 1)
    
    StockValue = ws.Cells(i, 7).Value + StockValue
    
    
    If Ticker <> ws.Cells(i + 1, 1) Then
    
    ws.Cells(Count, 9) = Ticker
    Closed = ws.Cells(i, 6).Value
    YC = Closed - Opened
    ws.Cells(Count, 10) = YC

    
 If Opened = 0 Then
 PC = 0
 Else
    PC = YC / Opened
    Opened = ws.Cells(i + 1, 3).Value
   
    End If
    ws.Cells(Count, 12).Value = StockValue
    'Extra Credit Loops
    If GPI < PC Then
    GPI = PC
    GPITicker = Ticker
    
    
    End If
    If GPD > PC Then
    GPD = PC
    GPDTicker = Ticker
    End If
     
     If GTV < StockValue Then
     GTV = StockValue
     GTVTicker = Ticker
     End If
     
    StockValue = 0
    
   
    
    ws.Cells(Count, 11).Value = PC
    
    Count = Count + 1
    
 
   
   
    End If

    

Next i

For i = 2 To (Count - 1)


 If ws.Cells(i, 10).Value < 0 Then

    
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    

End If

Next i

MsgBox ("Done")

'print

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest Percent Increase"
ws.Cells(3, 15).Value = "Greatest Percent Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(2, 17).Value = GPI
ws.Cells(3, 17).Value = GPD
ws.Cells(4, 17).Value = GTV
ws.Cells(2, 16).Value = GPITicker
ws.Cells(3, 16).Value = GPDTicker
ws.Cells(4, 16).Value = GTVTicker
Next ws
End Sub
