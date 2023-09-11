Sub StockData_Ticker()

    ' Declare worksheet variable
    Dim ws As Worksheet
    
    
    '   Looping through sheets one by one
    For Each ws In Worksheets
        
        '   Section 1
        '   Loop through present data to get the Ticker, Yearly Change, Percentage Change and Total Stock Volume of the stock for one year
        
       
	'   Declare variable to loop for every row
        Dim i As Long

 	'   Looping variable depends on the value of i, start row of the Ticker
        Dim j As Long
        
        '   Counter to fill the Ticker row
        Dim TickerCount As Long
        
        '   Declare variable to calculate and store Column K (Percent Change)
        Dim PerChange As Double
        
        '   Declare variable for the Last row in the Ticker (Column A)
        Dim LastRowColumnA As Long
        
        
        
        ' Initialise to start from row number 2
        TickerCount = 2
        
        ' Initialise to start from row number 2
        j = 2
       
        
        ' Set headers for the Stocks Summary Table
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        
        
        
        '   Fetch and store the value of the Last row in the Ticker (Column A)
        LastRowColumnA = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        '   Looping through every row in the current sheet to last row of Column A
        For i = 2 To LastRowColumnA
            '   Compare one row with the next for change of value in Column A
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                '   Apply conditional formatting on Column J, green for > than 0 and red for < 0
                If ws.Cells(TickerCount, 10).Value < 0 Then
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                End If
                
                '    Calculate and show value of Column K (Percent Change)
                If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
                    ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
                Else
                     ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                End If
                
                '   Calculate and show value of Column L (Total Stock Volume)
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                TickerCount = TickerCount + 1
                
                j = i + 1
            End If
        Next i
        
        '   Section 2
        '   This section will populate the fields Greatest % Increase , Greatest % Decrease and Greatest Total Volume according to the
        '   values populated in the table using the Section 1 code
        
        '   Set headers of the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" populated according to data from the Section 1 table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        '   To get/store the value of Last row in the Column I (Ticker) populated by the code in the Section 1
        Dim LastRowColumnI As Long
        
        '   To get/store the value of Greatest % Increase
        Dim GreatIncrease As Double
        
        '   To get/store the value of Greatest % Decrease
        Dim GreatDecrease As Double
        
        '   To get/store the value of Greatest Total Volume
        Dim GreatVolume As Double
        
        '   To get/store the value of Ticker corresponds to Greatest % Increase
        Dim GreatIncreaseTicker As String
        
        '   To get/store the value of Ticker corresponds to Greatest % Decrease
        Dim GreatDecreaseTicker As String
        
        '   To get/store the value of Ticker corresponds to Greatest % Volume
        Dim GreatVolumeTicker As String
        
        
        '   fetch and store the value of the Last row in the newly build Ticker (Column I)
        LastRowColumnI = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        '   Now initialise all declared variables with the first value in the row of newly formed table using Section 1 so that
        '   comparsions can be made by looping through each and every row
        
        
        '   Set an initial value from the Percent Change (Column K),value starts from Row 2
        GreatIncrease = ws.Cells(2, 11).Value
        
        '   Set an initial value from the Percent Change (Column K),value starts from Row 2
        GreatDecrease = ws.Cells(2, 11).Value
        
        '   Set an initial value from the Total Stock Volume (Column L),value starts from Row 2
        GreatVolume = ws.Cells(2, 12).Value
        
        '   If we get the minimum and maximum value in the first row then Ticker will be populated accordingly
        '   Set an initial value from the Ticker (Column I),value starts from Row 2
        GreatIncreaseTicker = ws.Cells(2, 9).Value
        
        '   Set an initial value from the Ticker (Column I),value starts from Row 2
        GreatDecreaseTicker = ws.Cells(2, 9).Value
        
        '   Set an initial value from the Ticker (Column I),value starts from Row 2
        GreatVolumeTicker = ws.Cells(2, 9).Value
        
        
        '   Loop through every row and compare with initialised values using if-else block
        For i = 2 To LastRowColumnI
            ' Match the initialise value with every row to get/set Greatest % Increase
            If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                GreatIncreaseTicker = ws.Cells(i, 9).Value
            Else
                GreatIncreaseTicker = GreatIncreaseTicker
                GreatIncrease = GreatIncrease
            End If
            
             ' Match the initialise value with every row to get/set Greatest % Decrease
            If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                GreatDecreaseTicker = ws.Cells(i, 9).Value
            Else
                GreatDecreaseTicker = GreatDecreaseTicker
                GreatDecrease = GreatDecrease
            End If
            
             ' Match the initialise value with every row to get/set Greatest % Volume
            If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                GreatVolumeTicker = ws.Cells(i, 9).Value
            Else
                GreatVolumeTicker = GreatVolumeTicker
                GreatVolume = GreatVolume
            End If
             
             
            '   Apply formatting to the values shown for Column Q (Greatest % Increase, Greatest % Decrease and Greatest Total Volume)
            '   with percentage sign and exponential

            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            '   To Assign/Show the values of Ticker to the Column P
            
            ws.Cells(2, 16).Value = GreatIncreaseTicker
            ws.Cells(3, 16).Value = GreatDecreaseTicker
            ws.Cells(4, 16).Value = GreatVolumeTicker
        
        Next i
        
        '   Section 3
        '   To adjust the width of every colum of every sheet according to the data length
        '   Declare the variable WorksheetName to get the name of each sheet in the Workbook
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
End Sub
