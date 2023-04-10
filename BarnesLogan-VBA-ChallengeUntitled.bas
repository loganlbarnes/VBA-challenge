Attribute VB_Name = "Module1"
Sub stocks():
    
'enable code to work on all tabs of the workbook
For Each ws In Worksheets
    
    'create variables... using Long for the total because it exceeds limits for doubles
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    'make total volume as a longlong variable and set at 0
    Dim Total As LongLong
    Total = 0
    
    'set opening price as the first value in the table
    Dim OpenPrice As Double
    OpenPrice = ws.Range("C2")
    
    Dim ClosePrice As Double
    
    'create rows for the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'add column headers and autofit columns
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Columns("A:Q").AutoFit
        
    'make i able to handle big numbers and set lastRow to the last row of the spreadsheet
    Dim i As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through all ticker values
    For i = 2 To lastRow
        
        'check if we are still within the same ticker name
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set the ticker name and print in I2
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(Summary_Table_Row, 9) = Ticker
            
            'set total value to the total plus i, then print the total value
            Total = Total + ws.Cells(i, 7).Value
            ws.Cells(Summary_Table_Row, 12).Value = Total
            
            'set closing price to the current i value
            ClosePrice = ws.Cells(i, 6).Value
            
            'calculate yearly change
            YearlyChange = ClosePrice - OpenPrice
            
            'calculate percent change
            PercentChange = YearlyChange / OpenPrice
            
            'print yearly change and percent change
            ws.Cells(Summary_Table_Row, 10).Value = YearlyChange
            ws.Cells(Summary_Table_Row, 11).Value = PercentChange
            'format percent change as a percent with 2 decimal points
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "#.##%"
            
             'Format YearlyChange to when greater than 0 it's green, when 0 it's yellow and when less than 0 is red
                If ws.Cells(Summary_Table_Row, 10) > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(Summary_Table_Row, 10) = 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 6
                    Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                End If
            
            'reset the summary table value to the next row and total to 0 and OpenPrice to the first row of the next ticker
            Summary_Table_Row = Summary_Table_Row + 1
            OpenPrice = ws.Cells(i + 1, 3).Value
            Total = 0
            
            Else
            
            'creates a running total for Total volume
            Total = Total + ws.Cells(i, 7).Value
             
        End If
        
    Next i
    
    'set the max, min, and volume as the first value in the summary table
    Max = ws.Range("K2").Value
    Min = ws.Range("K2").Value
    Volume = ws.Range("L2").Value
    
    'loop through second row to last row of the summary table
    lastSummaryRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    For j = 2 To lastSummaryRow
    
        'calculates the greatest % increase by looking at the current value and seeing if it's less than the current value
        If Max < ws.Cells(j + 1, 11).Value Then
            Max = ws.Cells(j + 1, 11).Value
            MaxTicker = ws.Cells(j + 1, 9).Value
        End If
        
       'calculates the greatest % decrease by looking at the current value and seeing if it's greater than the current value
        If Min > ws.Cells(j + 1, 11).Value Then
            Min = ws.Cells(j + 1, 11).Value
            MinTicker = ws.Cells(j + 1, 9).Value
        End If
    
        'greatest total volume
        If Volume < ws.Cells(j + 1, 12).Value Then
            Volume = ws.Cells(j + 1, 12).Value
            VolumeTicker = ws.Cells(j + 1, 9).Value
        End If
        
        'print greatest % increase with format and ticker name
        ws.Range("Q2").Value = Max
        ws.Range("Q2").NumberFormat = "#.##%"
        ws.Range("P2").Value = MaxTicker
        
        'print greatest % decrease with format and ticker name
        ws.Range("Q3").Value = Min
        ws.Range("Q3").NumberFormat = "#.##%"
        ws.Range("P3").Value = MinTicker
        
        'print greatest total volume with format and ticker name
        ws.Range("Q4").Value = Volume
        ws.Range("Q4").NumberFormat = "0.00E+00"
        ws.Range("P4").Value = VolumeTicker
        
    Next j

Next ws

MsgBox ("All Done")

End Sub

