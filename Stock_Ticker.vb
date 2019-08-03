
Sub Ticker()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Select

    Dim Last_Row, Vol_Row, Results_Last_Row As Integer
    Dim stock, Stk_Inc, Stk_Dec, Stk_Vol As String
    Dim Volume, Stk_Open, Stk_Close, Percent_Chg, Yearly_Chg, Greatest_Inc, Greatest_Dec, Greatest_Vol As Double

    Last_Row = Cells(Rows.Count, "A").End(xlUp).Row             ' Easy Level. Finds last row for initial spreadsheet.
   
    Vol_Row = 2                                                 ' Easy Level. We need to start on the second row to avoid header.
    Volume = 0
    Greatest_Inc = 0
    Greatest_Dec = 0
    Greatest_Vol = 0

    Cells(1, 9) = "Ticker"                      ' Easy Level
    Columns("I:I").EntireColumn.AutoFit
    Cells(1, 10) = "Yearly Change"              ' Moderate Level
    Columns("J:J").EntireColumn.AutoFit
    Cells(1, 11) = "Percent Change"             ' Moderate Level
    Columns("K:K").EntireColumn.AutoFit
    Cells(1, 12) = "Total Stock Volume"         ' Moderate Level
    Columns("L:L").EntireColumn.AutoFit
    

    Stk_Open = Cells(2, 3)  ' Moderate Level. Initial stock opening price on each worksheet.


    For i = 2 To Last_Row + 1                                   ' Easy Level. Runs one row past the last row for comparison purposes.

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then      ' Easy/Moderate Level. Sets up what to do when the ticker name changes.

            Cells(Vol_Row, 12) = Volume                         ' Easy Level. When the stock ticker changes, the Volume (determined below) is outputted.
            Stk_Close = Cells(i, 6)                             ' Moderate Level. Pulls in closing price when the ticker changes.
            Yearly_Chg = Stk_Close - Stk_Open                   ' Moderate Level. Difference between close and open.
            Cells(Vol_Row, 10).NumberFormat = "0.00000000"
            Cells(Vol_Row, 10) = Yearly_Chg                     ' Moderate Level. Outputs Yearly Change to the appropriate cell.

                If Yearly_Chg > 0 Then                              ' Moderate Level. Determines color of cell based on whether change is above or
                    Cells(Vol_Row, 10).Interior.ColorIndex = 4      ' below zero. Note that zero is excluded, as it is neither a gain nor a loss.
                ElseIf Yearly_Chg < 0 Then
                    Cells(Vol_Row, 10).Interior.ColorIndex = 3
                End If
                                         ' Moderate Level.
                If Stk_Open = 0 Then     ' If all opening prices for a stock are zero (see PLNT in 2014) then this fixes the zero denominator issue.
                    Percent_Chg = 0
                Else
                    Percent_Chg = Yearly_Chg / Stk_Open  ' Moderate Level. Determines Percent Change when stock opening price is above zero. (Every case except PLNT in 2014.)
                End If

            Cells(Vol_Row, 11).NumberFormat = "0.00%"
            Cells(Vol_Row, 11) = Percent_Chg            ' Moderate Level. Outputs Percent Change.
            Cells(Vol_Row, 9) = Cells(i, 1)     ' Easy Level. Pulls Ticker name and outputs it to appropriate cell.
            Vol_Row = Vol_Row + 1       ' Easy Level. Increments Volume Row when the stock ticker changes so that output is sequential.
            Volume = 0                  ' Easy Level. Re-initializes Volume for next stock symbol.
            Stk_Open = Cells(i + 1, 3)  ' Moderate Level. New opening price grabbed from the row below the closing price of previous stock.
        
            Else
                                                    ' Easy Level. Sums Volume so long as ticker remains the same.
                Volume = Volume + Cells(i + 1, 7)
                If Stk_Open = 0 Then                ' Moderate Level. Runs through rows of same stock name to find first non-zero open (PLNT 2015).
                    Stk_Open = Cells(i, 3)
                End If
                
   
        End If

    Next i

i = 2
Results_Last_Row = Cells(Rows.Count, "I").End(xlUp).Row     ' Hard Level. Finds last row of the moderate-level output.


    For i = 2 To Results_Last_Row                   ' Difficult Level. Only examines the rows of output created in earlier For loop.

        If Cells(i, 11).Value > Greatest_Inc Then   ' Difficult Level. Compares Percent Change (put in Greatest_Inc variable) and overwrites
            Greatest_Inc = Cells(i, 11).Value       ' variable whenever the Percent Change is higher than value already in the variable.
            Stk_Inc = Cells(i, 9).Value             ' Difficult Level. Pulls the Ticker name whenever Greatest_Inc is overwritten.
        End If

        If Cells(i, 11).Value < Greatest_Dec Then   ' Difficult Level. Compares Percent Change (put in Greatest_Dec variable) and overwrites
            Greatest_Dec = Cells(i, 11).Value       ' variable whenever the Percent Change is lower than value already in the variable.
            Stk_Dec = Cells(i, 9).Value             ' Difficult Level. Pulls the Ticker name whenever Greatest_Dec is overwritten.
        End If

        If Cells(i, 12).Value > Greatest_Vol Then   ' Difficult Level. Compares each row of Volume and overwrites Greatest_Vol variable whenever
            Greatest_Vol = Cells(i, 12).Value       ' Volume is larger than the value already in the variable.
            Stk_Vol = Cells(i, 9).Value             ' Difficult Level. Pulls the Ticker name whenever Greatest_Vol is overwritten.
        End If
    
    Next i

    ws.Range("P2") = Stk_Inc                           ' Difficult Level. Outputs data for the small table to the right.
    ws.Range("Q2") = Greatest_Inc
    ws.Range("P3") = Stk_Dec
    ws.Range("Q3") = Greatest_Dec
    ws.Range("P4") = Stk_Vol
    ws.Range("Q4") = Greatest_Vol

    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Range("O2") = "Greatest % Increase"         ' Hard Level
    Range("O3") = "Greatest % Decrease"         ' Hard Level
    Range("O4") = "Greatest Total Volume"       ' Hard Level
    Columns("O:O").EntireColumn.AutoFit
    Cells(1, 16) = "Ticker"                     ' Hard Level
    Columns("P:P").EntireColumn.AutoFit
    Cells(1, 17) = "Value"                      ' Hard Level
    Columns("Q:Q").EntireColumn.AutoFit

Next ws

End Sub


