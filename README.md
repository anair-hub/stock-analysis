# stock-analysis

## Overview Of the Project
    Stock analysis module was created to help Steve. In analyzing some stock data,  Steve wanted  to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
    Now one additional step is taken here to see how the program performs if there is increase in source data 

## Results
   
   Refactored Code:
   Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim Totalvolume(12) As Long
    Dim Startingprice(12) As Single
    Dim Endingprice(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        Totalvolume(tickerIndex) = 0
        Startingprice(tickerIndex) = 0
        Endingprice(tickerIndex) = 0
       
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
         If Cells(i, 1).Value = tickers(tickerIndex) Then
                Totalvolume(tickerIndex) = Totalvolume(tickerIndex) + Cells(i, 8).Value
         End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                Startingprice(tickerIndex) = Cells(i, 6).Value
                
          End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
          If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                Endingprice(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
        'End If
        End If
    Next i
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = Totalvolume(i)
        Cells(4 + i, 3).Value = Endingprice(i) / Startingprice(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

   ### 2017
   
   Original Stock Analysis Timings:

   ![Module_Stock_Analysis_2017](Resources/Module_Stock_Analysis_2017.png)
   

   Refactored Timings:

   ![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)

   
   As reflected , refactored code runs faster by making use of array data structures 

   ### 2018

   Original Stock Analysis Timings:

   ![Module_Stock_Analysis_2018](Resources/Module_Stock_Analysis_2018.png)
   
    
   Refactored Timings:

   ![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)
   

    As reflected , refactored code runs faster by making use of array data structures 

## Summary

   ### Advanatages and Disadvatages of Refactored code
       Main difference in the refactored code was using arrays for Totalvolume, Startingprice and Endingprice. Arrays can use multiple data of similar type at a time in continuous memory allocation pattern. If we have similar set of data then array creation is useful instead of taking multiple variables because it reduces length & complexity of program so the program execute much faster.

   ### Advantages and Disadvantages of Orginal(Versus refactored code)
       Advantage of origibal versus refactored is that it has better memory utilization whereas arrays have inefficient memory utilization. Also arrays have slow insertion/deletion time.