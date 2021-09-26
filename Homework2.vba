Attribute VB_Name = "Module1"

' Proto-code:
    ' Create a script that will loop through all the stocks for one year and output the following information:
    ' The ticker symbol.
    ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.
    ' You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
        ' The code will be a loop that will take the ticker and store it in a variable,
        ' it won't be changed until it is detected that there is a difference
        
        ' The loop will take the open


' Need to account for dividing by zero, or to not use 0 value as the first stock and look for a greater than zero value.

Sub homework2():
    ' Setup the routine to search through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        
        ' Set up the summary table titles
        ' Count the columns for the data file and move the summary table 3 columns over
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        Dim STSC As Integer
        STSC = LastColumn + 2 ' STSC = summary table start column
        STSR = 2
        Cells(1, STSC) = "Ticker"
        Cells(1, STSC + 1) = "Yearly Change"
        Cells(1, STSC + 2) = "Percent Change"
        Cells(1, STSC + 3) = "Total Stock Volume"
        ' AutoFit the Columns
        Cells(1, STSC).Columns.AutoFit
        Cells(1, STSC + 1).Columns.AutoFit
        Cells(1, STSC + 2).Columns.AutoFit
        Cells(1, STSC + 3).Columns.AutoFit
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create a for loop to scan through the data sheet
        Dim TickerTemp As String
        Dim YearStockOpen As Double
        YearStockOpen = 0
        Dim YearStockClose As Double
        Dim StockTemp As Double
        Dim ChangeCounter As Integer
        ChangeCounter = 0
        For Row = 2 To LastRow
            
            If Cells(Row, 3).Value <> 0 And ChangeCounter = 0 Then
                YearStockOpen = Cells(Row, 3).Value
            ElseIf Cells(Row, 3).Value <> 0 And YearStockOpen = 0 Then
                YearStockOpen = Cells(Row, 3).Value
            End If
            
            ' Populate the ticker column the one each of the ticker symbols
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                ' Set the ticker value
                TickerTemp = Cells(Row, 1).Value
                ' Print the ticker value to summary table
                Cells(STSR, STSC).Value = TickerTemp
                ' Add to the brand total
                StockTemp = StockTemp + Cells(Row, 7)
                ' Print total stock volume to summary table
                Cells(STSR, STSC + 3) = StockTemp
                ' Set the years closing stock price
                YearStockClose = Cells(Row, 6)
                ' Print the yearly Change in the Summary Table
                YearlyChange = YearStockClose - YearStockOpen
                Cells(STSR, STSC + 1).Value = YearlyChange
                ' Format the interior color of the cells <0 is red, >0 is green
                If Cells(STSR, STSC + 1) > 0 Then
                    Cells(STSR, STSC + 1).Interior.Color = RGB(0, 255, 0)
                ElseIf Cells(STSR, STSC + 1) < 0 Then
                    Cells(STSR, STSC + 1).Interior.Color = RGB(255, 0, 0)
                End If
                ' Print the yearly percent change
                Cells(STSR, STSC + 2).Value = FormatPercent(YearlyChange / YearStockOpen)
                ' Reset stock volume counter
                StockTemp = 0
                ' Add one to the summary table row counter
                STSR = STSR + 1
                ' Reset the change counter
                ChangeCounter = 0
            Else
                ' Add the stock volume while scanning the sheet
                StockTemp = StockTemp + Cells(Row, 7)
                ChangeCounter = ChangeCounter + 1
            End If
        Next Row
    Next ws


    Sheets(1).Activate ' Resets the worksheet to the first sheet
End Sub
