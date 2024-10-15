Sub StockData()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    
    ws.Activate
    
        'Ultima fila
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row

        'encabezados
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quartery Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
       Dim open_price As Double
       Dim close_price As Double
       Dim quarterly_change As Double
       Dim percent_change As Double
       Dim ticker As String
       
     
        Dim volume As Double
        Dim row As Double
        Dim column As Integer
        
        volume = 0
        row = 2
        column = 1
       
       
        
        open_price = Cells(2, column + 2).Value
        
        
        For i = 2 To last_row
        
         
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            
                'NOMBRES
                ticker = Cells(i, column).Value
                Cells(row, column + 8).Value = ticker
                
                'PRECIO SALIDA
                close_price = Cells(i, column + 5).Value
                
                'CAMBIO EN Q
                quarterly_change = close_price - open_price
                Cells(row, column + 9).Value = quarterly_change
               
                    'PORCENTAJE
                    percent_change = quarterly_change / open_price
                    Cells(row, column + 10).Value = percent_change
                    Cells(row, column + 10).NumberFormat = "0.00%"
               
               
               ' volumen total por Q
                volume = volume + Cells(i, column + 6).Value
                Cells(row, column + 11).Value = volume
                
            
                row = row + 1
                
                ' resetar contadores
                open_price = Cells(i + 1, column + 2)
                
                
                volume = 0
                
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i
        
        
        'encontrar ultima fila quarterly change
        
         quaterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        'colores
         For x = 2 To quaterly_change_last_row
            If (Cells(x, 10).Value > 0 Or Cells(x, 10).Value = 0) Then
                Cells(x, 10).Interior.ColorIndex = 10
            ElseIf Cells(x, 10).Value < 0 Then
                Cells(x, 10).Interior.ColorIndex = 3
            End If
        Next x
        
        'nuevos encabezados
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'encontar maximos, minimos
        
        For x = 2 To quaterly_change_last_row
            If Cells(x, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(2, 16).Value = Cells(x, 9).Value
                Cells(2, 17).Value = Cells(x, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(3, 16).Value = Cells(x, 9).Value
                Cells(3, 17).Value = Cells(x, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quaterly_change_last_row)) Then
                Cells(4, 16).Value = Cells(x, 9).Value
                Cells(4, 17).Value = Cells(x, 12).Value
            End If
        Next x
        
       ActiveSheet.Range("I:Q").EntireColumn.AutoFit
       Worksheets("Q1").Select
        
        
        Next ws
        
End Sub
# VBA-challenge