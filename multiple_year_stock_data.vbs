Attribute VB_Name = "Module1"
Sub finance()
Dim tickerID As String
Dim ticekrTableRow As Integer
Dim lasRowX, lastRowYearChange As Integer
Dim tickerRowsA As Integer
Dim firstOpen As Double
Dim lastClose As Double
Dim openTest, closeTest, yearChange, closeValue, openValue, percentageChange, maxPercentage As Double
Dim WorksheetName, wsCount As String

wsCount = ActiveWorkbook.Worksheets.Count


'gertting the ticker symbol
For k = 1 To wsCount
Sheets(k).Select                           'selected current page to repeat calculations

'Var----------------------------
lastRowX = Cells(Rows.Count, 1).End(xlUp).Row
tickerTableRow = 2                          'pointer on <ticker> column
openValue = Cells(2, 3).Value
'-------------------------------


    If Range("H1") = 0 Then
        Range("H1") = "Ticker"
        Range("I1") = "Yearly Change"
        Range("J1") = "percent Change"
        Range("K1") = "Total Stock Volume"
        Range("N2") = "Greatest % Increase"
        Range("N3") = "Greatest % Decrease"
        Range("N4") = "Greatest Total Volume"
    End If
    
    
    volume = 0
    
    For i = 2 To lastRowX
    
        volume = volume + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then  'starting from the value below
            tickerID = Cells(i, 1).Value                    'assigning the ticker value to the variable
            Range("H" & tickerTableRow).Value = tickerID    'adding the value to the column
            
            'with this block yearly change is calculated
            closeValue = Cells(i, 6).Value                                    'close value is set to the first item on the for loop
            yearChange = closeValue - openValue                               ' close value is sunstracted from open value
            Range("I" & tickerTableRow).Value = yearChange                    'the result is assing to the variable
            
            'with this block percentage change is calculated
            If openValue <> 0 Then
                percentageChange = (yearChange / openValue)
            Else
                percentageChange = 0
            End If
            
            Range("J" & tickerTableRow).Value = percentageChange
            Range("J" & tickerTableRow).NumberFormat = "0.00%"
            
    
            
            
            
            'with this block the color is assinged to percentage change denpending on its value
            
            If Range("I" & tickerTableRow).Value > 0 Then
                Range("I" & tickerTableRow).Interior.ColorIndex = 4
                Else
                    Range("I" & tickerTableRow).Interior.ColorIndex = 3
                
            End If
    
            Range("K" & tickerTableRow).Value = volume
             volume = 0
             
            openValue = Cells(i + 1, 3).Value
            tickerTableRow = tickerTableRow + 1
            
        End If
    
        
    Next i
    
    'Bonus
    maxPercentage = Application.WorksheetFunction.Max(Range(Cells(2, 10), Cells(Rows.Count, 10)))
    Range("P2").Value = maxPercentage
    Range("P2").NumberFormat = "0.00%"
    
    maxPercentageMin = Application.WorksheetFunction.Min(Range(Cells(2, 10), Cells(Rows.Count, 10)))
    Range("P3").Value = maxPercentageMin
    Range("P3").NumberFormat = "0.00%"
    
    volumeMax = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(Rows.Count, 11)))
    Range("P4").Value = volumeMax
    
     'end of ticket symbol
     MsgBox Worksheets(k).Name
     
     
     
Next k
 

 
End Sub

