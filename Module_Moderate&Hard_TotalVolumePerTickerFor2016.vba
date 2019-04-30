Attribute VB_Name = "Module4"
'VBA Script Homework Moderate And Hard Challange
'Moderate
'Create a script that will loop through all the stocks for one year for each run and take the following information.
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'Hard
'Your solution will include everything from the moderate challenge.
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
'Solution will look as follows

Sub TotalVolumePerTicker2016()

'Create a variable to count the total volume for each ticker
Dim TotalVolumePerYear As Double
Dim i As Long
Dim j As Long
Dim lastrow As Long
Dim OpenStockPrice As Double
Dim ClosedStockPrice As Double
Dim GreatesTotalVolume As Double
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolumeTicker As String
Dim GreatestPercentDecreaseTicker As String
Dim GreatestPercentIncreaseTicker As String


'Counts the number of rows in a sheet
With Sheets("2016")
    If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
        lastrow = .Cells.Find(What:="*", _
                      After:=.Range("A1"), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Row
    Else
        lastrow = 1
    End If
End With


'Initialize parameters
TotalVolumePerYear = 0
OpenStockPrice = Cells(2, 3).Value
ClosedStockPrice = 0
j = 2
GreatesTotalVolume = 0
GreatestPercentIncrease = 0
GreatestPercentDecrease = 0

'Look up through each row and count the total volume per ticket
For i = 2 To lastrow
        'If the next row value for the ticket column is not same as the previous row, then run the last count for Total Volume Per Year and save the  ticket no and total value in two new columns
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            TotalVolumePerYear = TotalVolumePerYear + Cells(i, 7) 'Total Volume Calculation
            If GreatesTotalVolume < TotalVolumePerYear Then 'Greatest Volume Calculation within same loop
                GreatesTotalVolume = TotalVolumePerYear
                GreatestTotalVolumeTicker = Cells(i, 1).Value
            End If
            ClosedStockPrice = Cells(i, 6).Value
            Cells(j, 11).Value = Cells(i, 1).Value
            Cells(j, 12).Value = TotalVolumePerYear
            Cells(j, 13).Value = ClosedStockPrice - OpenStockPrice 'Calculate the difference/change from years day1 Open Price and last day Close Price per ticker
            
            If (Cells(j, 13).Value < 0) Then ' Color index definition if the difference is negetive vs positive
                Cells(j, 13).Interior.ColorIndex = 3
            Else
                Cells(j, 13).Interior.ColorIndex = 4
            End If
            
            Cells(j, 14).Value = (ClosedStockPrice - OpenStockPrice) / OpenStockPrice 'Percentage change calculation of stock ticker price
            If Cells(j, 14).Value < GreatestPercentDecrease Then 'Calculate the Greatest % Decrease
                GreatestPercentDecrease = Cells(j, 14).Value
                GreatestPercentDecreaseTicker = Cells(j, 1).Value
            End If
            
            If Cells(j, 14).Value > GreatestPercentIncrease Then 'Calculate the Greatest % Increase
                GreatestPercentIncrease = Cells(j, 14).Value
                GreatestPercentIncreaseTicker = Cells(j, 1).Value
            End If
    
            TotalVolumePerYear = 0
            j = j + 1
            OpenStockPrice = Cells(i + 1, 3).Value
        'If the next row value for the ticket column is same as previous sum the total volume
        Else
            TotalVolumePerYear = TotalVolumePerYear + Cells(i, 7)
            
        End If
        
Next i

Cells(2, 17).Value = GreatestPercentIncreaseTicker
Cells(2, 18).Value = GreatestPercentIncrease
Cells(3, 17).Value = GreatestPercentIncreaseTicker
Cells(3, 18).Value = GreatestPercentDecrease
Cells(4, 17).Value = GreatestTotalVolumeTicker
Cells(4, 18).Value = GreatesTotalVolume

End Sub












