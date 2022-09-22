Attribute VB_Name = "Module1"
Sub runaccrosssheets():
'This code runs vbachallenge accross all the active sheets.
'This code was found on google
'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html

    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call vbaChallenge
    Next
    Application.ScreenUpdating = True
End Sub

Sub vbaChallenge():

'Define some counters for the loops
Dim counter As Double
Dim i As Double
Dim j As Double


'Fill in the headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L01").Value = "Total Stock Volume"

'Get the number of rows in the active sheet

counter = Cells(Rows.Count, 1).End(xlUp).Row


'Initialize ticker, yearChange, and the other variables
j = 2
tickerInitial = Range("A2").Value
yearlychange = Range("F2").Value - Range("C2").Value
percentchange = yearlychange / Range("C2").Value
volume = Range("G2").Value
For i = 2 To counter - 1
    If Cells(i + 1, 1) = tickerInitial Then
        yearc = Cells(i + 1, 6) - Cells(i + 1, 3)
        yearlychange = yearlychange + yearc
        percentC = yearc / Cells(i + 1, 3)
        percentchange = percentchange + percentC
        volume = volume + Cells(i + 1, 7)
    Else
        Cells(j, 9) = tickerInitial
        Cells(j, 10) = yearlychange
        Cells(j, 11) = percentchange
        Cells(j, 12) = volume
        tickerInitial = Cells(i + 1, 1)
        j = j + 1
        yearlychange = Cells(i + 1, 6) - Cells(i + 1, 3)
        percentchange = yearlychange / Cells(i + 1, 3)
        volume = Cells(i + 1, 7)
    End If
Next i

'Greatest percents and total volume
'Headers for the summary table

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest total volume"
percentIncrease = 0
percentDecrease = 0
totalVolume = 0
For i = 2 To j - 1
    If Cells(i, 10) > 0 Then
        Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    Else
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    End If
    
    'Initialize the percent change to verify whether it's the greatest increase or decrease
    
    Cells(i, 11) = FormatPercent(Cells(i, 11))
    
    'If statement to get the greatest perecent changes.
    If Cells(i, 11) > percentIncrease Then
        percentIncrease = Cells(i, 11)
        greatTicker = Cells(i, 9)
    End If
    If Cells(i, 11) < percentDecrease Then
        percentDecrease = Cells(i, 11)
        leastTicker = Cells(i, 9)
    End If
    
    If Cells(i, 12) > totalVolume Then
        totalVolume = Cells(i, 12)
        volumeTicker = Cells(i, 9)
    End If
    
    
Next

'Fill in summary table

Range("O2") = greatTicker
Range("O3") = leastTicker
Range("O4") = volumeTicker
Range("P2") = FormatPercent(percentIncrease)
Range("P3") = FormatPercent(percentDecrease)
Range("P4") = totalVolume

'Autofit columns with long headers or content

Range("J1").EntireColumn.AutoFit
Range("K1").EntireColumn.AutoFit
Range("L1").EntireColumn.AutoFit
Range("N3").EntireColumn.AutoFit
End Sub


