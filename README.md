# stock_analysis
Performing analysis on stock data to uncover daily volume and yearly returns.
stock_analysis 

#**VBA_Challenge**

#**Overview of Project**

Steve,our client, wants to analyze green energy stocks and compare the total daily volume and yearly return for each stock.
 To do so I will Create a VBA macro that can automate these analyses on a click of a button.
 >Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year.

1. Create a worksheet to hold the data. Adding a header and assiging cell  row values.
2. Calculate the total daily volume using loops, conditionals and code pattern.
3. Calculate the yearly return of stocks by determining the first closing price and the last closing price.
4. Format the output sheet to make it easier to visualize.
5. Repurpose the VBA macros to analyze multiple stocks.

##**Purpose**

Using the green_stocks dataset we can refactor a Microsoft Excel VBA code to collect certain stoc information for the year 2017 and 2018 and determine which stocks had a positive yearly return and how active each stock was traded.

#**Analysis and Challenges**

**Analysis of Daily Volume and Yearly Return of Stocks**

**1. Create a worksheet to hold the data analysis for "DQ" stocks.**
    -Calculate the daily volume in 2018 using loops for "DQ"stocks.
    -Calculate the yearly return of " DQ" stocks by determining the first closing price and last closing price.

**2. Create a new worksheet to hold data of "All Stocks Analysis".**
    - Reuse the code from DQAnalysis and change the text to work on AllSticksAnalysis.
    - Copy the code from DQAnalysis and make the following changes:

        >Activate "All Stocks Analysis" instead of "DQ Analysis."
        
        >Change the A1 value to "All Stocks (2018)."
        
        >Change the first column header to "Ticker."
        
**3. >Our new macro should do the following:**
 
    '1)Format the output sheet on the "All Stocks Analysis" worksheet.
            
            Worksheets("All Stocks Analysis").Activate
            
            Range("A1").Value = "All Stocks (2018)"
        
        'Create a header row
            Cells(3, 1).Value = "Ticker"
            Cells(3, 2).Value = "Total Daily Volume"
            Cells(3, 3).Value = "Return"
   
        
    '2)Initialize an array of all tickers.
        Dim tickers(11) As String
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
        
    '3)Prepare for the analysis of tickers.
        
        '3a)Initialize variables for the starting price and ending price.
            
            Dim startingPrice As Single
            Dim endingPrice As Single
        
        '3b)Activate the data worksheet.
            
            Worksheets("2018").Activate
        
        '3c)Find the number of rows to loop over.
             
             RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4)Loop through the tickers.
       
       For i = 0 to 11
       ticker = tickers(i)
       totalVolume = 0
       
     '5) loop through rows in the data
       
       Worksheets("2018").Activate
       For j = 2 to RowCount
           
        
        '5a)Find the total volume for the current ticker.
        
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
        
        '5b)Find the starting price for the current ticker.
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If
        
        '5c)Find the ending price for the current ticker.
        
         If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
    '6)Output the data for the current ticker.
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

End Sub

**4. Debugging, going through the code to make sure the code is working properly.

**5. Static Formatting:
     
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

**6. Conditional Formatting:
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

**7. Create a run button to automatically run analysis when pressed.

**8. Create a run the ClearWorksheet macro. In order to reset the analyzed data.

**9. Replace Hard-Coded Values to run the analysis for any year.
    
    -Add at the beginning of a new macro:
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    - Replace 
    Range("A1").Value = "All Stocks (2018)" for
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    -Replace first to get the row count, and inside the "For" loop.
    Worksheets("2018").Activate for 
    Sheets(yearValue).Activate

**10. Measure code performance. The amount of time it will take to run "All Stocks Analysis"
        - First we will have to determine the start and end time, then set each variable equal to the "Timer" function
    Sub AllStocksAnalysis()
        Dim startTime As Single
        Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer

        - Scroll to the end of the AllStocksAnalysis script and add end time.
        Next i

    endTime = Timer
    
        -Add msg box to show start and end time for each year
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

**11. Refactor code:


 '1a) Create a ticker Index

  tickerIndex = 0

   '1b) Create three output arrays

  Dim tickerVolumes(12) As Long

  Dim tickerStartingPrices(12) As Single

  Dim tickerEndingPrices(12) As Single

'2a) Create a for loop to initialize the tickerVolumes to zero.

' If the next row’s ticker doesn’t match, increase the tickerIndex.

    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0

Next i

'2b) Loop over all the rows in the spreadsheet.

    For i = 2 To RowCount

 '3a) Increase volume for current ticker
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i


Macros completed

###**Challenges**

1. First challenge was to write a macro in VBA get the numbers of rows with data:
    Solution was found in Stackoverflow
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
https://stackoverflow.com/questions/18088729/row-count-where-data-exists/41538965

2. Second challenge was replacing the Hard-Coded Values. Especially knowing where to activate the "worksheets(yearValue).Activate" to get the row count and initialize "FOR" loop. 
https://zoom.us/rec/play/3y8QSEzGpawN1aFaIdTTB07ORUoGVwlzCxrhpsY_vSxIxMMKwIEolwHZrynFP_TXJ1XLWYVjWcEC4rxd.4q3FJMwSSyZxC3BG?startTime=1637415343000&_x_zm_rtaid=-i-cEITxScOE-OFMLjGzqQ.1637430207851.2075cf95887cdc9463ec29789c0911c9&_x_zm_rhtaid=958

3. Third challenge was setting up a "Nested For Loop".
    https://zoom.us/rec/play/3y8QSEzGpawN1aFaIdTTB07ORUoGVwlzCxrhpsY_vSxIxMMKwIEolwHZrynFP_TXJ1XLWYVjWcEC4rxd.4q3FJMwSSyZxC3BG?startTime=1637415343000&_x_zm_rtaid=-i-cEITxScOE-OFMLjGzqQ.1637430207851.2075cf95887cdc9463ec29789c0911c9&_x_zm_rhtaid=958

#**Results**


By looking at the Microsoft Excel sheet " All Stock Analysis" we can conclude that in the year 2017 out of the 12 tickers analyzed only one "TERP" had a negative yearly return. While the other 11 tickers yearly return percent change varies between 8.9% to 199.4%. While during the year 2018, 10 out of the 12 tickers show a negative yearly returns, with only "ENPH" and "RUN" showing a positive percent change of 81% and 84% respectively. It is important to note that the by refactoring the VBA script total macro run time decreased by approximately 0.47 seconds.
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/93225405/142737218-fd962561-7310-4b84-b608-2fede3e1841e.png">

<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/93225405/142737248-439ef45c-b2ec-46a8-85fa-dd6d17ba971a.png">


#**Summary**

**Advantages and Disadvantages of Refactoring code in general:

Refactoring helps to organize and clean the macros so that VBA script can run faster. The advantages of refactoring a code is that it will run more efficiently, by taking fewer steps,using less memory, or to improve the logic of the code to make it easier to read.One disavantage is the added time it needs to refactor a code that is already working. Also a refactor code might need additional information or data.

**Advantages and Disadvantages of the original and refactored VBA script:

The biggest advantage that occurred by refactoring the VBA script was an decrease in the total macro run time.Originally for the year 2017 the code ran in 0.72 seconds after refactoring the code fro 2017 ran for 0.25 seconds. The same can be viewed for the year 2018.
