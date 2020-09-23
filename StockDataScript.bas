Attribute VB_Name = "Module1"
Sub StockDataScript():
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Worksheets
    
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
        LastRow = sh.Range("A1").End(xlDown).Row
        Ticker = sh.Cells(2, 1).Value
        TotalStockVol = 0
        OpenPrice = sh.Cells(2, 3).Value
        GreatestInc = 0
        GreatestDec = 0
        GreatestTotVol = 0
        
        '-----Headers for Stock Market Tables-----
        sh.Cells(1, 9).Value = "Ticker"
        sh.Cells(1, 10).Value = "Yearly Change"
        sh.Cells(1, 11).Value = "Percent Change"
        sh.Cells(1, 12).Value = "Total Stock Volume"
        sh.Cells(1, 16).Value = "Ticker"
        sh.Cells(1, 17).Value = "Value"
        sh.Cells(2, 15).Value = "Greatest % Increase"
        sh.Cells(3, 15).Value = "Greatest % Decrease"
        sh.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        For i = 2 To LastRow
            '-----Calculate Variables for Analyzing Table-----
            If sh.Cells(i, 1).Value = Ticker Then
                TotalStockVol = TotalStockVol + sh.Cells(i, 7).Value
            Else
                ClosePrice = sh.Cells(i - 1, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice = 0 Then
                    For k = i To LastRow
                        If sh.Cells(k, 3).Value <> 0 Then
                            OpenPrice = sh.Cells(k, 3).Value
                            PercentChange = YearlyChange / OpenPrice
                            Exit For
                        Else
                        End If
                    Next k
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
                '-----Update Analyzing Table with New Row-----
                sh.Cells(TableRow, 9).Value = Ticker
                sh.Cells(TableRow, 10).Value = YearlyChange
                If YearlyChange < 0 Then
                    sh.Cells(TableRow, 10).Interior.ColorIndex = 3
                Else
                    sh.Cells(TableRow, 10).Interior.ColorIndex = 4
                End If
                sh.Cells(TableRow, 11) = Format(PercentChange, "Percent")
                sh.Cells(TableRow, 12).Value = TotalStockVol
                '-----Update Greatest Values Variables and Table-----
                If PercentChange > GreatestInc Then
                    sh.Cells(2, 16).Value = Ticker
                    GreatestInc = PercentChange
                    sh.Cells(2, 17) = Format(GreatestInc, "Percent")
                Else
                End If
                
                If PercentChange < GreatestDec Then
                    sh.Cells(3, 16).Value = Ticker
                    GreatestDec = PercentChange
                    sh.Cells(3, 17) = Format(GreatestDec, "Percent")
                Else
                End If
                
                If TotalStockVol > GreatestTotVol Then
                    sh.Cells(4, 16).Value = Ticker
                    GreatestTotVol = TotalStockVol
                    sh.Cells(4, 17).Value = GreatestTotVol
                Else
                End If
                
                '-----Prepare/Reset Variables for Next Ticker Group-----
                TableRow = TableRow + 1
                Ticker = sh.Cells(i, 1).Value
                TotalStockVol = 0
                OpenPrice = sh.Cells(i, 3).Value
        
            End If
               
        Next i
    
    Next sh

    
End Sub


