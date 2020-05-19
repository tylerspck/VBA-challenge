Attribute VB_Name = "Module1"
'VBA Homework

Sub stockticker()
    'Ticker Name
    Dim Ticker As String
    'Make a row counter
    Dim Row_Counter As Long
    'making a counter for the print table row to increase
    Dim Ticker_Table As Integer
    'conter for how many rows for each ticker symbol
    Dim yearlypts As Integer
    'setting the row that contains the opening value
    Dim openvalue As Long
    
    'setting challenges var
    'looking for max increase percent
    Dim Max_increase As Double
    'looking for max decrease percent
    Dim max_decrease As Double
    'looking for largest volume
    Dim max_volume As Double
    'assigning ticker for max increase
    Dim increase_ticker As String
    'assigning ticker for max decrease
    Dim decrease_ticker As String
    'assigning ticker for max volume
    Dim volume_ticker As String
    
    'starting point for rows for each ticker symbol
    yearlypts = 1
    'starting point for where the opening value is located
    openvalue = 0
     
        'Loop for each worksheet
        For Each ws In Worksheets
            'set row counter for each worksheet
            Row_Counter = 1
            'set start of printout table for each worksheet
            Ticker_Table = 2
            'set headers for new tables in each worksheet
            ws.Range("i1").Value = "Ticker"
            ws.Range("j1").Value = "Yearly Change"
            ws.Range("k1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            'reset var for challeenges for each worksheet
            Max_increase = 0
            max_decrease = 0
            max_volume = 0
     
            'Find Last Row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'start on row 2
            For i = 2 To LastRow
                
                'Creating a row counter to increase for each row
                Row_Counter = Row_Counter + 1
                    
                'creating conditional to find how many points are for each ticker (added and conditional to rule out rows that have zeros in them)
                If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i, 3).Value <> 0 Then
                    
                    yearlypts = yearlypts + 1
                
                'testing for yearlypts correct value
                Else
                    'using to see if yearlypts is correct
                    'ws.Range("k" & Ticker_Table).Value = yearlypts
                

                    'create conditional to the change of the stock ticker symbol
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                        'test for reseting ticker name
                        'msgbox(cells(i, 1.value))
                        Ticker = ws.Cells(i, 1).Value
                    
                        'Copying ticker symbol to new row in printout table
                        ws.Range("I" & Ticker_Table).Value = Ticker
                    
                        'using to see if yearlypts is correct
                        'ws.Range("L" & Ticker_Table).Value = yearlypts
                    
                        'create value to use for opening value of each stock ticker symbol
                        openvalue = Row_Counter - yearlypts + 1
                        
                        'calc the difference of the year closing value from the year opening value
                        ws.Range("j" & Ticker_Table).Value = ws.Cells(Row_Counter, 6).Value - ws.Cells(openvalue, 3).Value
                        ws.Range("j" & Ticker_Table).NumberFormat = "0.00"
                        
                        'Conditional formatting for the cells depending if positive or negitive
                        If ws.Range("j" & Ticker_Table).Value > 0 Then
                            ws.Range("j" & Ticker_Table).Interior.ColorIndex = 4
                                
                            ElseIf ws.Range("j" & Ticker_Table).Value < 0 Then
                                ws.Range("j" & Ticker_Table).Interior.ColorIndex = 3
                                
                            Else
                                ws.Range("j" & Ticker_Table).Interior.ColorIndex = 0
                                
                        End If
                        
                            'calc the percent change at year end from year opening value
                            If (ws.Cells(Row_Counter, 6).Value - ws.Cells(openvalue, 3).Value) = 0 Then
                                ws.Range("k" & Ticker_Table).Value = 0
                                ws.Range("k" & Ticker_Table).NumberFormat = "0.00%"
                            
                            'fix for stock starting to trade part way through the year
                            ElseIf (ws.Cells(Row_Counter, 6).Value - ws.Cells(openvalue, 3).Value) <> 0 And ws.Cells(openvalue, 3).Value = 0 Then
                                openvalue = openvalue + 1
                                
                                
                                If ws.Cells(openvalue, 3).Value <> 0 Then
                                    ws.Range("k" & Ticker_Table).Value = (ws.Cells(Row_Counter, 6).Value - ws.Cells(openvalue, 3).Value) / ws.Cells(openvalue, 3).Value
                                    ws.Range("k" & Ticker_Table).NumberFormat = "0.00%"
                                    
                                End If
                            
                            Else
                                
                                ws.Range("k" & Ticker_Table).Value = (ws.Cells(Row_Counter, 6).Value - ws.Cells(openvalue, 3).Value) / ws.Cells(openvalue, 3).Value
                                ws.Range("k" & Ticker_Table).NumberFormat = "0.00%"
                                
                            End If
                        
                       'calc total stock volume
                       ws.Range("L" & Ticker_Table).Value = Application.Sum(Range(ws.Cells(openvalue, 7), ws.Cells(Row_Counter, 7)))
                       ws.Range("L" & Ticker_Table).NumberFormat = "general"
                       
                        '----------------Challenges------------
                        'finding the max increase
                       If ws.Range("k" & Ticker_Table).Value > Max_increase Then
                            Max_increase = ws.Range("k" & Ticker_Table).Value
                            increase_ticker = ws.Range("i" & Ticker_Table).Value
                                
                            'finding the max decrease
                             ElseIf ws.Range("k" & Ticker_Table).Value < max_decrease Then
                                max_decrease = ws.Range("k" & Ticker_Table).Value
                                decrease_ticker = ws.Range("i" & Ticker_Table).Value
                                
                         End If

                        'Finding the max volume
                        If ws.Range("L" & Ticker_Table).Value > max_volume Then
                            max_volume = ws.Range("L" & Ticker_Table).Value
                            volume_ticker = ws.Range("i" & Ticker_Table).Value
                        End If
                         
                         'setting up challenges table
                         ws.Range("o2").Value = "Greatest % Increase"
                         ws.Range("o3").Value = "Greatest % Decrease"
                         ws.Range("o4").Value = "Greatest Total Volume"
                         ws.Range("p1").Value = "Ticker"
                         ws.Range("Q1").Value = "Value"
                         ws.Range("p2").Value = increase_ticker
                         ws.Range("p3").Value = decrease_ticker
                         ws.Range("p4").Value = volume_ticker
                         ws.Range("Q2").Value = Max_increase
                         ws.Range("q2").NumberFormat = "0.00%"
                         ws.Range("Q3").Value = max_decrease
                         ws.Range("q3").NumberFormat = "0.00%"
                         ws.Range("q4").Value = max_volume
                         ws.Range("q4").NumberFormat = "general"


                        'increase the table row
                        Ticker_Table = Ticker_Table + 1
                
                        'reset counters for next ticker
                        yearlypts = 1
                        openvalue = 0
                    End If

                End If

            Next i

        Next ws

End Sub

