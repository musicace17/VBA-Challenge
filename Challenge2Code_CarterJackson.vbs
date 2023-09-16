Attribute VB_Name = "Module1"
Sub ChallengeCode()
    Dim ws As Worksheet
    
    'Loop through all workbook sheets by starting with a for loop
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        
        'Declare the variables to be used
        Dim LastRow As Long
        Dim OpenPrice As Double
        Dim YearlyChange As Double
        Dim SummaryRow As Integer
        Dim PercentChange As Double
        
        
        'Label the output columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change ($)"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        'Find the Last Row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'Predefine some variables before beginning the loop
        Totalvolume = 0
        OpenPrice = Cells(2, 3).Value
        SummaryRow = 2
        
        
        'Begin looping through each ticker and add the volumes as you go
        For i = 2 To LastRow
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                Totalvolume = Totalvolume + Cells(i, 7).Value
                
            'Print the values gained and calculate other summary values
            Else
                Totalvolume = Totalvolume + Cells(i, 7).Value
                ClosePrice = Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                PercentChange = YearlyChange / OpenPrice * 100
                'Print the results to their respective columns
                Cells(SummaryRow, 9).Value = Cells(i, 1).Value
                Cells(SummaryRow, 10).Value = YearlyChange
                Cells(SummaryRow, 11).Value = "%" & PercentChange
                Cells(SummaryRow, 12).Value = Totalvolume
                
            'Now format the Yearly Change cell to be green or red based on value
                If Cells(SummaryRow, 10).Value > 0 Then
                    Cells(SummaryRow, 10).Interior.ColorIndex = 4
                ElseIf Cells(SummaryRow, 10).Value < 0 Then
                    Cells(SummaryRow, 10).Interior.ColorIndex = 3
                End If
            
            'Set up the variables for the next loop
            Totalvolume = 0
            OpenPrice = Cells(i + 1, 3).Value
            SummaryRow = SummaryRow + 1
            
            End If
        Next i
        
        
        'Declare a couple of new variables
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim IncTick As String
        Dim DecTick As String
        Dim VolTick As String
        
        'Predefine the variables
        GreatInc = 0
        GreatDec = 0
        GreatVol = 0
        IncTick = " "
        DecTick = " "
        VolTick = " "
        
        
        
        'Now loop through the results columns to store and print superlative results
  For y = 2 To LastRow
            If Cells(y, 11).Value > GreatInc Then
                GreatInc = Cells(y, 11).Value
                IncTick = Cells(y, 9).Value
            End If
            
            If Cells(y, 11).Value < GreatDec Then
                GreatDec = Cells(y, 11).Value
                DecTick = Cells(y, 9).Value
            End If
            
            If Cells(y, 12).Value > GreatVol Then
                GreatVol = Cells(y, 12).Value
                VolTick = Cells(y, 9).Value
            End If
        Next y
       
            
        'Print the 3 stored variables to results columns (and label those rows/columns)
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells(2, 17).Value = "%" & GreatInc * 100
        Cells(2, 16).Value = IncTick
        Cells(3, 17).Value = "%" & GreatDec * 100
        Cells(3, 16).Value = DecTick
        Cells(4, 17).Value = GreatVol
        Cells(4, 16).Value = VolTick
        
        'Clean up some column sizes by autofitting them!
        Columns(10).AutoFit
        Columns(11).AutoFit
        Columns(12).AutoFit
        Columns(15).AutoFit
        
        
    Next ws
End Sub
Sub ClearResults()

    'Declare variables prior to any commands
    Dim ws As Worksheet
    Dim FinalRow As Long
    Dim rng As Range
    Dim cell As Range
    
    'Run this macro across all sheets with a For loop
    For Each ws In ThisWorkbook.Sheets

        ' Find the last row with contents in columns I to Q
        FinalRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' Define the range from row 1 to the last row in columns I to Q
        Set rng = ws.Range("I1:Q" & FinalRow)
        
        ' Loop through each cell in the range and clear contents and colors
        For Each cell In rng
            cell.ClearContents
            cell.Interior.ColorIndex = xlNone
        Next cell

        'Reset the column widths
        ws.Cells.Columns("I:Q").ColumnWidth = 8
        

    Next ws
End Sub
