Attribute VB_Name = "Module1"
'create a script that loops through all the stocks for each quarter and outputs the following information:
' the ticker symbol
' quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
' the percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
' the total stock volume of the stock


Sub challenge()

    Dim i As Long
    Dim j As Integer
    Dim wksht As Worksheet
    
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim change As Double
    Dim pchange As Double
    Dim volume As Double
    Dim total As Double
    
    Dim start As Long
    Dim rowcount As Long
    Dim days As Integer
    Dim dailychange As Double
    Dim averagechange As Double
    
    
    'set title row of summary table
    Range("I1") = "Ticker"
    Range("J1") = "Quarterly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'set initial values
    j = 0
    total = 0
    change = 0
    start = 2
    
    'get row # of last row of data
    rowcount = Cells(Rows.Count, 1).End(xlUp).Row
    
    'begin loop
    For i = 2 To rowcount
            
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            total = total + Cells(i, 7).Value
        
            If total = 0 Then
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
            
            Else
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
            
            change = (Cells(i, 6) - Cells(start, 3))
            pchange = change / Cells(start, 3)
            
            'start of next stock ticker
            start = i + 1
            
            'print results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = change
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = pchange
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Range("L" & 2 + j).Value = total
            
            'color-coding positives GREEN, negatives RED
            Select Case change
                Case Is > 0
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
            
        End If
    
        'resetting variables
        total = 0
        change = 0
        j = j + 1
        days = 0
        
    'if ticker is still the same, add results
        Else
            total = total + Cells(i, 7).Value
            
        End If
    
    Next i
    
    'take the max and min, place them in separate part of worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowcount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowcount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowcount))
    
    'returns 1 less because header of row is not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowcount)), Range("K2:K" & rowcount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowcount)), Range("K2:K" & rowcount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowcount)), Range("L2:L" & rowcount), 0)
    
    'final ticker symbol for total, greatest % of increase & decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)
    

End Sub
