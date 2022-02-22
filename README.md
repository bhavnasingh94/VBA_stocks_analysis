# VBA_stocks_analysis
#Overview of Project
  The purpose of the project was to to expand the dataset to include the entire stock market over the years 2017 and 2018 to determine whether the stocks are worth investing on or not. The goal was to refactor the solution code to through all the data one time in order to collect the same information.
  # Analysis 
    I began by activating DQAnalysis worksheets, created loops through stock data, use the ticketindex to access the stock ticker index, had the TickerIndex set to equal to zero before looping over the rows. Here are my codes.
    Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

'1)Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2)Initialize an array of all tickers.

    Dim tickers(12) As String
    'Creates and array with 12 elements
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
  
'3a)Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    
'3b)Activate the data worksheet.
    Worksheets(yearValue).Activate

'3c)Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4)Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        'Do stuff with ticker
        totalVolume = 0
    
'5)Loop through rows in the data.
    Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
    'a)Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
    End If
    
    'b)Find the starting price for the current ticker.
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            startingPrice = Cells(j, 6).Value
        End If
        'Determines the beginning of the ticker section
        
    'c)Find the ending price for the current ticker.
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        'Determines the end of the ticker section
        
    Next j
    
'6)Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
 Next i


Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

'Color Formatting
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
    
    If Cells(i, 3) > 0 Then
        Cells(i, 3).Interior.Color = vbGreen
        'Color the cell green
    ElseIf Cells(i, 3) < 0 Then
        Cells(i, 3).Interior.Color = vbRed
        'Color the cell red
    Else
        Cells(i, 3).Interior.Color = xlNone
        'Clear the cell color
    
    End If
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

#Summary
#Pros and Cons
Refactoring the code is intended to improve the design, and implementation of the code while maintaining its functionality. However, further code refactoring activities may require additional effort to regain this knowledge. (Nassif, Matthieu; Robillard, Martin P, 2017) The benefit in the refactoring process of the stocks analysis is that decrease in the running time. For this case, it ran for 0.199 seconds for 2017 stock analysis and 
0.187 seconds for the year 2018 as following
![VBA_Challenge_2017_](https://user-images.githubusercontent.com/98790082/155231576-da1e025f-4c23-4ad2-ae70-31b879b284d1.png)

![VBA_Challenge_2018_](https://user-images.githubusercontent.com/98790082/155231550-99949c3b-9984-49be-bf87-5694f008caf4.png)
