Sub one_run()
    Dim counter As Integer
    Dim lastRow As Long
    Dim change As Double
    Dim ticker()
    Dim open_prices()
    Dim close_prices()
    Dim total_vol()
    Dim temp_vol
    Dim top(3, 2)
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    
           
    
    For i = 0 To ActiveWorkbook.Sheets.Count - 1
        
        lastRow = Sheets(i + 1).Cells(Rows.Count, 1).End(xlUp).Row
        counter = 0
        
        top(1, 2) = 0
        top(2, 2) = 0
        top(3, 2) = 0
        
        Sheets(i + 1).Cells(1, 9).Value = "Ticker"
        Sheets(i + 1).Cells(1, 10).Value = "Yearly Change"
        Sheets(i + 1).Cells(1, 11).Value = "Percet Change"
        Sheets(i + 1).Cells(1, 12).Value = "Total Stock Volume"
        Sheets(i + 1).Cells(1, 15).Value = "Ticker"
        Sheets(i + 1).Cells(1, 16).Value = "Value"
        Sheets(i + 1).Cells(2, 14).Value = "Greatest % Increse"
        Sheets(i + 1).Cells(3, 14).Value = "Greatest % Decrese"
        Sheets(i + 1).Cells(4, 14).Value = "Greatest Total Volume"
        Sheets(i + 1).Activate
        Columns.AutoFit
        
        Sheets(i + 1).Columns(10).FormatConditions.Delete
        Set condition1 = Sheets(i + 1).Columns(10).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set condition2 = Sheets(i + 1).Columns(10).FormatConditions.Add(xlCellValue, xlLess, "=0")
        With condition1
            .Interior.ColorIndex = 4
        End With
        With condition2
            .Interior.ColorIndex = 3
        End With
        
        temp_vol = 0
        
        For j = 2 To lastRow
            
            temp_vol = temp_vol + Sheets(i + 1).Cells(j, 7).Value
            
            If Sheets(i + 1).Cells(j, 1).Value <> Sheets(i + 1).Cells(j + 1, 1).Value Then
    
                counter = counter + 1
                ReDim Preserve ticker(counter)
                ReDim Preserve open_prices(counter + 1)
                ReDim Preserve close_prices(counter)
                ReDim Preserve total_vol(counter)
                ticker(counter) = Sheets(i + 1).Cells(j, 1).Value
                open_prices(counter + 1) = Sheets(i + 1).Cells(j + 1, 3).Value
                close_prices(counter) = Sheets(i + 1).Cells(j, 6).Value
                total_vol(counter) = temp_vol
                temp_vol = 0
                
            End If
        Next j
        For j = 1 To counter
        
            Sheets(i + 1).Cells(j + 1, 9).Value = ticker(j)
            change = open_prices(j) - close_prices(j)
            Sheets(i + 1).Cells(j + 1, 10).Value = change
            If open_prices(j) <> 0 Then
                percentage_change = FormatPercent(change / open_prices(j))
            Else
                percentage_change = FormatPercent(0)
            End If
            Sheets(i + 1).Cells(j + 1, 11).Value = percentage_change
            Sheets(i + 1).Cells(j + 1, 12).Value = total_vol(j)
            
            If top(1, 2) < percentage_change Then
                top(1, 1) = ticker(j)
                top(1, 2) = percentage_change
            End If
            
            If top(2, 2) > percentage_change Then
                top(2, 1) = ticker(j)
                top(2, 2) = percentage_change
            End If
            
            If top(3, 2) < total_vol(j) Then
                top(3, 1) = ticker(j)
                top(3, 2) = total_vol(j)
            End If
            
        Next j
        
        For j = 1 To 3
            Sheets(i + 1).Cells(j + 1, 15).Value = top(j, 1)
            Sheets(i + 1).Cells(j + 1, 16).Value = top(j, 2)
        Next j
                    
    Next i
    
    
End Sub


