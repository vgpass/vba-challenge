Sub analysis()

'selecting a different worksheet

Dim ws_count As Integer
ws_count = Application.Sheets.Count
Dim ws_sheets As Integer

    For ws_sheets = 1 To ws_count
        Sheets(ws_sheets).Select

    'Subroutine to output ticker, yearly change, percent change, and total stock volume for each year
    'Format spreadsheet

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Range("O2:O4").Font.Bold = True
        Range("P1:Q1").Font.Bold = True
        Range("I1:L1").EntireColumn.AutoFit
        Range("A1:L1").Font.Bold = True
        Range("A1:Q1").HorizontalAlignment = xlCenter
        
        
    'Dim variable to be used in subroutine

        Dim ticker As String
        Dim last As Long
        Dim j, c, i As Long

    'Determine last row in data range/
    'Found on https://www.exceldemy.com/excel-vba-find-last-row-with-data-in-range/
        
        last = Range("A2").End(xlDown).Row
        
    'Format columns for large numbers and percentage

        Range("L1:L" & last).NumberFormat = "#,###"
        Range("K1:K" & last).NumberFormat = "0.00%"
        Range("Q4").NumberFormat = "#,###"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
    'Loop and counter to determine ticker symbols

        c = 2
        i = 2
            For j = 2 To last
            
                If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
                    
                    ticker = Cells(j, 1).Value
                    Cells(i, 9) = ticker
                    Cells(i, 12) = WorksheetFunction.Sum(Range("G" & c & ":G" & j))
                    Cells(i, 10) = Cells(j, "F").Value - Cells(c, "C").Value
                    Cells(i, 11) = Round((Cells(j, "F").Value - Cells(c, "C").Value) / Cells(c, "C").Value, 4)
                    
                        If Cells(i, 10).Value < 0 Then
                            Cells(i, 10).Interior.ColorIndex = 3
                            
                        ElseIf Cells(i, 10).Value > 0 Then
                            Cells(i, 10).Interior.ColorIndex = 4
                            
                        End If
                    
                    i = i + 1
                    c = 1
                    c = c + j
                
                ElseIf Cells(j, 1).Value = Cells(j + 1, 1).Value Then
                
                End If
                
            Next j
        
    'Freezing top row of spreadsheet. https://stackoverflow.com/questions/3232920/how-can-i-programmatically-freeze-the-top-row-of-an-excel-worksheet-in-excel-200

        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
        Cells(1, 1).Select

    'Sorting for greatest % increase, greatest % decrease and greatest total volume
    'Resouce found at https://www.delftstack.com/howto/vba/vba-sort/#:~:text=Sort%20Data%20Range%20by%20Specific%20Column%20Using%20the,is%20included%20in%20the%20sorting%20process%20or%20not.

        last = Range("K2").End(xlDown).Row
            Range("I2:L" & last).Sort Key1:=Range("K1"), _
                Order1:=xlAscending, _
                Header:=xlNo
                
                    Range("Q3") = Range("K2").Value
                    Range("P3") = Range("I2").Value
                    Range("Q2") = Range("K" & last).Value
                    Range("P2") = Range("I" & last).Value
                
            Range("I2:L" & last).Sort Key1:=Range("L1"), _
                Order1:=xlDescending, _
                Header:=xlNo
                
                    Range("Q4") = Range("L2").Value
                    Range("P4") = Range("I2").Value
                
            Range("I2:L" & last).Sort Key1:=Range("I1"), _
                Order1:=xlAscending, _
                Header:=xlNo
                
        Range("O1:Q1").EntireColumn.AutoFit
        
    Next ws_sheets

End Sub