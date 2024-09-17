Attribute VB_Name = "Module1"
'create a script that loops through all the stocks for each quarter and outputs the following information:
' the ticker symbol
' quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
' the percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
' the total stock volume of the stock


Sub challenge()
Dim i As Long
Dim wksht As Worksheet

Dim ticker As String
Dim finalrow As Long
Dim openprice As Double
Dim closeprice As Double
Dim qchange As Double
Dim pchange As Double
Dim totalvol As Double
Dim tablerow As Integer

For Each wksht In ThisWorkbook.Worksheets
wksht.Cells(1, 9).Value = "Ticker"
wksht.Cells(1, 10).Value = "Quarterly Change"
wksht.Cells(1, 11).Value = "Percent Change"
wksht.Cells(1, 12).Value = "Total Stock Volume"

wksht.Cells(1, 16).Value = "Ticker"
wksht.Cells(1, 17).Value = "Value"
wksht.Cells(2, 15).Value = "Greatest % Increase"
wksht.Cells(3, 15).Value = "Greatest % Decrease"
wksht.Cells(4, 15).Value = "Greatest Total Volume"

wksht.Range("Q2").NumberFormat = "0.00%"
wksht.Range("Q3").NumberFormat = "0.00%"

wksht.Cells(2, 17).Value = 0
wksht.Cells(3, 17).Value = 0
wksht.Cells(4, 17).Value = 0

i = 2
tablerow = 1
finalrow = wksht.Cells(wksht.Rows.Count, "A").End(xlUp).Row
openprice = wksht.Cells(i, 3).Value

For i = 2 To finalrow
        
    ticker = wksht.Cells(i, 1).Value
    totalvol = totalvol + wksht.Cells(i, 7).Value
    
    If wksht.Cells(i - 1, 1).Value <> wksht.Cells(i, 1).Value Then
    openprice = wksht.Cells(i, 3).Value
    Else
    If wksht.Cells(i + 1, 1).Value <> wksht.Cells(i, 1).Value Then
        closeprice = wksht.Cells(i, 6).Value
        
        qchange = closeprice - openprice
        pchange = qchange / openprice
        
        tablerow = tablerow + 1
        
        wksht.Cells(tablerow, 9).Value = ticker
        wksht.Cells(tablerow, 10).Value = qchange
        wksht.Cells(tablerow, 11).Value = pchange
        wksht.Cells(tablerow, 12).Value = totalvol
        
        wksht.Columns("K").NumberFormat = "0.00%"
        
        If qchange < 0 Then
            wksht.Cells(tablerow, 10).Interior.ColorIndex = 3
        Else
            wksht.Cells(tablerow, 10).Interior.ColorIndex = 4
        End If
        
        If totalvol > wksht.Cells(4, 17).Value Then
            wksht.Cells(4, 17).Value = totalvol
            wksht.Cells(4, 16).Value = wksht.Cells(i, 1)
        End If
        
        If pchange > wksht.Cells(2, 17).Value Then
            wksht.Cells(2, 17).Value = pchange
            wksht.Cells(2, 16).Value = wksht.Cells(i, 1)
        End If
        If pchange < wksht.Cells(3, 17).Value Then
            wksht.Cells(3, 17).Value = pchange
            wksht.Cells(3, 16).Value = wksht.Cells(i, 1)
        End If
        
        totalvol = 0
        
    End If
    End If
Next i

Next wksht
End Sub


