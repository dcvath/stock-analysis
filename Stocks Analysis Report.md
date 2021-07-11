# Green Energy Stocks Analysis Written Report

## Overview of Project
### Purpose
Analyze green energy stocks to help Steve identify what stocks his parents should invest their funds in.

### Background
Steve's partents are passionate about green energy and they have some money to invest in it. They have not done much research and originally decided to invest all of their money in DAQO New Energy Corp. Steve just graduated with his Finance Degree and he promised to look into DAQO stock for his parents, but he wants to diversify their funds, so he is analyzing some green energy stocks, in addition to DAQO stock.

## Results

### Comparision of Stock Performance Between 2017 and 2018

#### 2017
For 2017, all of the stocks had a positive return, ranging from RUN at 5.5% to DQ at 199.4%, except for TERP, which had a negative return of -7.2%.

<img width="233" alt="2017 Stocks Analysis" src="https://user-images.githubusercontent.com/85654649/125211770-bd9af700-e276-11eb-8299-f51ac0b460d6.png">


#### 2018
In contrast, for 2018, all of the stocks had a negative return, except for two: ENPH had a postive return of 81.9% and RUN had a positive return of 84%. 

<img width="235" alt="2018 Stocks Analysis" src="https://user-images.githubusercontent.com/85654649/125211782-ce4b6d00-e276-11eb-8734-753dea33685f.png">

#### Looking Across Both Years (2017 and 2018)
When looking at 2017 and 2018 as a whole, the only two stocks that had a positive return both years in a row were ENPH and RUN. For 2017, ENPH had a positive return of 129.5% and RUN had a positive return of 5.5%. For 2018, ENPH had a positive return 81.9% and RUN had a positive return of 84%.

### Execution Times of Original Script and the Refactored Script
As the screen shots below indicate, the refactored script on both 2017 and 2018 ran faster than the original scripts. the Using images and examples of your code as well as the execution times of the original script and the refactored script.

#### Original Script 2017
Ran at 0.20 seconds.
<img width="274" alt="2017 Original Script" src="https://user-images.githubusercontent.com/85654649/125212465-972b8a80-e27b-11eb-89df-38cdbe363ad1.png">

#### Refactored Script 2017
Ran at 0.18 seconds.
<img width="283" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85654649/125211421-38164780-e274-11eb-8ee9-48bcc8a1d833.png">

#### Original Script 2018
Ran at 0.20 seconds.
<img width="265" alt="2018 Original Script" src="https://user-images.githubusercontent.com/85654649/125212467-9dba0200-e27b-11eb-8726-ec970cbdeb77.png">

#### Refactored Script 2018
Ran at 0.18 seconds.
<img width="283" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85654649/125211425-3b113800-e274-11eb-9016-9297427b7bc6.png">

#### My Refactored AllStocksAnalysis Code
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


## Summary
### Advantages of Refactoring Code
Two key advantages of refactoring code are optimization and opportunities to improve the logic for future readers of your code. For example, optimizing your code to run faster when you are working with a large data set could be a huge asset as it will save time and puts less strain on your computer's resources. Another example would be improving the logic, so that if you are out of the office one day and something needs to be changed, your coworker can easily pick up where you left off.

### Disadvantages of Refactoring Code
Disadvantages of refactoring code are that is takes time to go back through the code and creates a new risk of errors because anytime you edit your code, you could make an error that breaks the entire script.

### Pros and cons apply to refactoring the original VBA script?

