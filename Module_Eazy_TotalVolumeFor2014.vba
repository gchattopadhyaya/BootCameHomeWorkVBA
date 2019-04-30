Attribute VB_Name = "Module1"
'VBA Script Homework Part1
'Create a script that will loop through one year of stock data for each run and return the total volume each stock had over that year.
'You will also need to display the ticker symbol to coincide with the total stock volume.
'Your result should look as follows (note: all solution images are for 2015 data).

Sub TotalVolumePerTicker()

'Create a variable to count the total volume for each ticker
Dim TotalVolumePerYear As Double
Dim i As Long
Dim j As Long
Dim lastrow As Long


'Counts the number of rows in a sheet
With Sheets("2014")
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


'Initialize counter and j
TotalVolumePerYear = 0
j = 2

'Look up through each row and count the total volume per ticket
For i = 2 To lastrow
        'If the next row value for the ticket column is not same as the previous row, then run the last count for Total Volume Per Year and save the  ticket no and total value in two new columns
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            TotalVolumePerYear = TotalVolumePerYear + Cells(i, 7)
            Cells(j, 11).Value = Cells(i, 1).Value
            Cells(j, 12).Value = TotalVolumePerYear
            TotalVolumePerYear = 0
            j = j + 1
        'If the next row value for the ticket column is same as previous sum the total volume
        Else
            TotalVolumePerYear = TotalVolumePerYear + Cells(i, 7)
            
        End If
        
Next i


End Sub
