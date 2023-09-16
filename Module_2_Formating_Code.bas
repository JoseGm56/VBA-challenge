Attribute VB_Name = "Module1"
Sub Formating()
    For Each ws In Worksheets
    
        Dim tickerSym As String
        Dim indicator As Integer
        Dim openingValue As Double
        Dim vol As Double
        Dim indicator2 As Integer
        Dim TotalVol As Double
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        vol = 0
        indicator = 2
        indicator2 = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        openingValue = ws.Cells(2, 3).Value
        For i = 2 To LastRow
            vol = vol + Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Place the correct data on the correct cells
                ws.Cells(indicator, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(indicator, 10).Value = ws.Cells(i, 6).Value - openingValue
                ws.Cells(indicator, 11).Value = ws.Cells(indicator, 10).Value / openingValue
                ws.Cells(indicator, 12).Value = vol
                            
                'Update indicator, opening value and volume
                indicator = indicator + 1
                openingValue = ws.Cells(i + 1, 3).Value
                vol = 0
                LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
            End If
        Next i
        GreatestIncrease = ws.Cells(2, 11)
        GreatestDecrease = ws.Cells(2, 11)
        TotalVol = ws.Cells(2, 12)
        For x = 2 To LastRow2
            
            'Gets the greatest % increase and places it in correct cell
            If ws.Cells(x + 1, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(x + 1, 11).Value
                ws.Cells(2, 17).Value = GreatestIncrease
                ws.Cells(2, 16).Value = ws.Cells(x + 1, 9).Value
            End If
            
            'Gets the greatest % decrease and places it in correct cell
            If ws.Cells(x + 1, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(x + 1, 11).Value
                ws.Cells(3, 17).Value = GreatestDecrease
                ws.Cells(3, 16).Value = ws.Cells(x + 1, 9).Value
            End If
            
            'Gets the total volume and places it in correct cell
            If ws.Cells(x + 1, 12).Value > TotalVol Then
                TotalVol = ws.Cells(x + 1, 12).Value
                ws.Cells(4, 17).Value = TotalVol
                ws.Cells(4, 16).Value = ws.Cells(x + 1, 9).Value
            End If
            
            'Setting the color for the yearly chage
            If ws.Cells(x, 10).Value >= 0 Then
                ws.Cells(x, 10).Interior.Color = RGB(0, 225, 0)
            Else
                ws.Cells(x, 10).Interior.Color = RGB(225, 0, 0)
            End If
            
        Next x
            
        'ws.Cells(3, 17).Style.NumberFormat = "0.00%"
        'ws.Cells(2, 17).Style.NumberFormat = "0.00%"
        
    Next ws
End Sub

