Sub credit_card()

    Dim brand As String
    Dim brand_total As Long
    Dim summ_table_row As Integer
    Dim firstValue As Long
    
    summ_table_row = 1
    summ_table_row2 = 1
    brand_total = 0
    
    lastrow = (Cells(Rows.Count, "A").End(xlUp).Row)
    'MsgBox (lastrow)
    
    For i = 2 To lastrow
        brand_total = brand_total + Cells(i, 3).Value
        '-- IF creates summary table --
        '-Fill outs when there is only one record for CCType
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Fill outs First Value column
            summ_table_row2 = summ_table_row2 + 1
            Cells(summ_table_row2, 9).Value = Cells(i, 3).Value
            'Fill outs Total Charged column
            brand = Cells(i, 1).Value
            summ_table_row = summ_table_row + 1
            Cells(summ_table_row, 7).Value = brand
            Cells(summ_table_row, 8).Value = brand_total
            brand_total = 0
            'Get Last Value from Amount
            Cells(summ_table_row, 10).Value = Cells(i, 3).Value
            'Calculate Sum of First Value and Last Value
            Cells(summ_table_row, 11).Value = Cells(i, 3).Value + Cells(i, 3).Value
            
        '-Fill outs when there are multiple records for CCType
        'Fill outs Total Charged column by Brand
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Fill outs Total Charged column
            brand = Cells(i, 1).Value
            summ_table_row = summ_table_row + 1
            Cells(summ_table_row, 7).Value = brand
            Cells(summ_table_row, 8).Value = brand_total
            brand_total = 0
            'Get Last Value from Amount
            Cells(summ_table_row, 10).Value = Cells(i, 3).Value
            'Calculate Sum of First Value and Last Value
            Cells(summ_table_row, 11).Value = Cells(i, 3).Value + firstValue
            
        'Fill outs First Value when there are multiple records for CCType
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            summ_table_row2 = summ_table_row2 + 1
            Cells(summ_table_row2, 9).Value = Cells(i, 3).Value
            firstValue = Cells(i, 3).Value
        End If
    Next i

End Sub
