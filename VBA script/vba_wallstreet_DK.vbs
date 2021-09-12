Sub stockChecker()

Dim i As Long
Dim ticker As String
Dim yearOpen As Double
Dim yearClose As Long
Dim yearChange As Double
Dim percentChange As Variant
Dim totalVolume As LongLong
Dim lastRow As Long
Dim ws As Worksheet
Dim index As Long



For Each ws In Worksheets

    index = 2
    totalVolume = 0

    'insert new headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Columns("L").AutoFit

    'record initial yearOpen
    yearOpen = ws.Cells(2, 3).Value
    
        'determine the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastRow
            
           
            
            'record total stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            
            'find ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                    'get and write ticker
                    ticker = ws.Cells(i, 1).Value
                    ws.Cells(index, 9).Value = ticker
            
                    'get and write calculate yearly change
                    yearClose = ws.Cells(i, 6).Value
                    yearChange = yearClose - yearOpen
                    ws.Cells(index, 10).Value = yearChange
                    
                    'format cells
                    red = 3
                    green = 4
                    If ws.Cells(index, 10).Value >= 0 Then
                        ws.Cells(index, 10).Interior.ColorIndex = green
                        Else
                        ws.Cells(index, 10).Interior.ColorIndex = red
                    End If
                
                                
                    'calculate and write percent change
                    If yearOpen = 0 Then
                        percentChange = "Cannot Calculate"
                        ws.Cells(index, 11).Value = percentChange
                        Else
                        percentChange = yearChange / yearOpen
                        ws.Cells(index, 11).Value = percentChange
                        ws.Cells(index, 11).Style = "Percent"
                    End If
                    
                    'write total stock volume
                    ws.Cells(index, 12).Value = totalVolume
            
                    'record new year open
                    yearOpen = ws.Cells(i + 1, 3)
                
                    'record and write totalvolume
                    ws.Cells(index, 12).Value = totalVolume
                    
                    
                    'reset total stock volume to zero
                    totalVolume = 0
                 
                
                    'increase index
                    index = index + 1
                    
                    'last row
                    
                    
    
                End If
                
        Next i

Next ws

End Sub


