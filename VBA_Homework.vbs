Attribute VB_Name = "Module1"
Sub StockMarket()

'Need to loop though all worksheets
For Each ws In Worksheets

'Declaring Variables
Dim OpenStock As Double
Dim CloseStock As Double
Dim Ticker As String
Dim OutputRow As String
OutputRow = 2
Dim YearlyChange As Double
Dim Percent As Double
Dim TotalStockVolume As Double
TotalStockVolume = 0

'Count Last Row on each worksheet
Dim LastRow As Long
'Placing a Variable for the previous amount
Dim PreviousAmount As Long
PreviousAmount = 2

'To find the LastRow in the Worksheets
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Setting the Headers in the Columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent"
ws.Range("L1").Value = "Total Stock Volume"

    
For i = 2 To LastRow
    
    
    'Finding the Ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(OutputRow, 9).Value = Ticker
        'Adding and resetting the total stock volume
        TotalStockVolume = TotalStockVolume + ws.Range("G" & i).Value
        ws.Cells(OutputRow, 12).Value = TotalStockVolume
        TotalStockVolume = 0
        'Calculating the yearly change
        OpenStock = ws.Range("C" & PreviousAmount)
        CloseStock = ws.Range("F" & i)
        YearlyChange = CloseStock - OpenStock
        ws.Cells(OutputRow, 10).Value = YearlyChange
        
        'Finding the percent change
        If OpenStock = 0 Then
            Percent = 0
        Else
            OpenStock = ws.Range("C" & PreviousAmount)
            Percent = YearlyChange / OpenStock
        End If
        'Formatting the Percentages
        ws.Range("K" & OutputRow).NumberFormat = "0.00%"
        ws.Range("K" & OutputRow).Value = Percent
        'Formatting the Yearly Change to Number format for Conditional Formatting
        ws.Range("J" & OutputRow).NumberFormat = "0.00"
            
        'Conditional Formatting Positive(Green)& Negative(Red)
        If ws.Range("J" & OutputRow).Value >= 0 Then
            ws.Range("J" & OutputRow).Interior.ColorIndex = 4
        Else
           ws.Range("J" & OutputRow).Interior.ColorIndex = 3
        End If
        'Setting up the Output row to add one
        OutputRow = OutputRow + 1
        PreviousAmount = i + 1
        
    Else
        ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value
        TotalStockVolume = TotalStockVolume + ws.Range("G" & i).Value
       
    End If
    
Next i
'CHALLENGE
    
    'Declaring the Greatest Variables & thier default position
    Dim GreatestIncrease As Double
    GreastIncrease = 0
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestTotal As Double
    GreatestTotal = 0
    
    'Setting the Headers for the Columns in the Challenge Part
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Counting the Last Row for the % Increase & Decrease
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
'Give the Greatest % Increase, Decrease, Total Value w/Ticker and Value
For i = 2 To LastRow
        If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
        End If

        If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
           ws.Range("Q4").Value = ws.Cells(i, 12).Value
           ws.Range("P4").Value = ws.Cells(i, 9).Value
        End If
Next i
    
    'Formating to Greatest Percent Value to Number format of %
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'Autofitting each Column
    ws.Columns("I:Q").AutoFit
    
Next ws
End Sub


