Attribute VB_Name = "Module1"
Sub RealVBAChallenge()

Dim StrTicker As String
StrTicker = "Ticker"

Dim DbYearlyChange, DbPercentChange, DbTSV As Double
DbYearlyChange = "Yearly Change"
DbPercentChange = "Percent Change"
DbTotalStockVolume = "Total Stock Volume"

DbTSV = 0

Range("I1").Value = StrTicker
Range("J1").Value = DbYearlyChange
Range("K1").Value = DbPercentChange
Range("L1").Value = DbTotalStockVolume


Dim SumTable As Integer
SumTable = 2

Dim OpenPrice As Double
Dim ClosePrice As Double

Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow


   If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
   
OpenPrice = Cells(i, 3).Value

End If


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ClosePrice = Cells(i, 6).Value
StrTicker = Cells(i, 1).Value

DbTSV = DbTSV + Cells(i, 7).Value

DbYearlyChange = ClosePrice - OpenPrice
DbPercentChange = DbYearlyChange / OpenPrice

Range("I" & SumTable).Value = StrTicker
Range("L" & SumTable).Value = DbTSV
Range("J" & SumTable).Value = DbYearlyChange
Range("K" & SumTable).Value = DbPercentChange
Range("K" & SumTable).NumberFormat = "0.00%"

SumTable = SumTable + 1
DbTSV = 0
DbYearlyChange = 0
OpenPrice = Cells(i + 1, 3).Value

    Else
    DbTSV = DbTSV + Cells(i, 7).Value
    DbYearlyChange = ClosePrice - OpenPrice
    DbPercentChange = DbYearlyChange / OpenPrice
  
End If

    Next i



For i = 2 To SumTable


    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 10
    
    Else
    Cells(i, 10).Interior.ColorIndex = 3
    
    
End If

    Next i
       

    
End Sub

