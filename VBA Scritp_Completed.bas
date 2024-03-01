Attribute VB_Name = "Module1"
Sub stock_data()
Dim ws As Worksheet
Dim lastrow As LongLong
Dim ticker As String
Dim opening As Double
Dim closing As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim total As Double
Dim i As LongLong
Dim j As Integer
Dim change As Double
Dim start As Integer

Set ws = ThisWorkbook.Sheets("2018")
For Each ws In ActiveWorkbook.Sheets

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'headers
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    j = 0
    total = 0
    change = 0
    start = 2
    
        'math
        
For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            total = total + ws.Cells(i, 7).Value
            
            If total = 0 Then
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
                
            Else
            
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
            
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            percentchange = change / ws.Cells(start, 3)
            
            
            
       'input data
       
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = change
            ws.Range("J" & 2 + j).NumberFormat = "0.00"
            ws.Range("K" & 2 + j).Value = percentchange
            ws.Range("K" & 2 + j).NumberFormat = "0.00%"
            ws.Range("L" & 2 + j).Value = total
            
'color
            Select Case change
                Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
                

            End If
            
            total = 0
            change = 0
            j = j + 1
            
            
        Else
            total = total + ws.Cells(i, 7).Value
                    
    
    End If

    Next i

    
Next ws


End Sub

