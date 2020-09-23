Attribute VB_Name = "Module1"
Sub StockDataScript():
'-----Variable List-----
Dim i As Long
Dim j As Integer
Dim TableRow As Integer
Dim LastRow As Double
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVol As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim GreatestInc As Double
Dim GreatestDec As Double
Dim GreatestTotVol As Double

TableRow = 2
LastRow = Range("A1").End(xlDown).Row
Ticker = Cells(2, 1).Value
TotalStockVol = 0
OpenPrice = Cells(2, 3).Value
GreatestInc = 0
GreatestDec = 0
GreatestTotVol = 0

'-----Headers for Stock Market Tables-----
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


For i = 2 To LastRow
    If Cells(i, 1).Value = Ticker Then
        TotalStockVol = TotalStockVol + Cells(i, 7).Value
    Else
        '-----Calculate Variables for Analyzing Table-----
        ClosePrice = Cells(i - 1, 6).Value
        YearlyChange = ClosePrice - OpenPrice
        If OpenPrice = 0 Then
            For k = i To LastRow
                If Cells(k, 3).Value <> 0 Then
                    OpenPrice = Cells(k, 3).Value
                    PercentChange = YearlyChange / OpenPrice
                    Exit For
                Else
                End If
            Next k
        Else
            PercentChange = YearlyChange / OpenPrice
        End If
        '-----Update Analyzing Table with New Row-----
        Cells(TableRow, 9).Value = Ticker
        Cells(TableRow, 10).Value = YearlyChange
        If YearlyChange < 0 Then
            Cells(TableRow, 10).Interior.ColorIndex = 3
        Else
            Cells(TableRow, 10).Interior.ColorIndex = 4
        End If
        Cells(TableRow, 11) = Format(PercentChange, "Percent")
        Cells(TableRow, 12).Value = TotalStockVol
        '-----Update Greatest Values Variables and Table-----
        If PercentChange > GreatestInc Then
            Cells(2, 16).Value = Ticker
            GreatestInc = PercentChange
            Cells(2, 17).Value = GreatestInc
        Else
        End If
        
        If PercentChange < GreatestDec Then
            Cells(3, 16).Value = Ticker
            GreatestDec = PercentChange
            Cells(3, 17).Value = GreatestDec
        Else
        End If
        
        If TotalStockVol > GreatestTotVol Then
            Cells(4, 16).Value = Ticker
            GreatestTotVol = TotalStockVol
            Cells(4, 17).Value = GreatestTotVol
        Else
        End If
        
        '-----Prepare/Reset Variables for Next Ticker Group-----
        TableRow = TableRow + 1
        Ticker = Cells(i, 1).Value
        TotalStockVol = 0
        OpenPrice = Cells(i, 3).Value

    End If
       
Next i

End Sub

