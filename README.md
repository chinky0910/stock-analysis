Stock Analysis With Excel VBA

Overview of Project

Purpose

The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code.

The Data

The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

Results

Analysis

Before refactoring the code, I began with code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.

'1a) Create a ticker Index
tickerIndex = 0

'1b) Create three output arrays
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim tickerVolumes(12) As Long

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
 For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

''2b) Loop over all the rows in the spreadsheet.
For j = 2 To RowCount

    '3a) Increase volume for current ticker
      If Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
     If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j,       1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

        '3d Increase the tickerIndex.
       tickerIndex = tickerIndex + 1
        End If

    Next j

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For k = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + k, 1).Value = tickers(i)
    Cells(4 + k, 2).Value = tickerVolumes(i)
    Cells(4 + k, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
   Next k

Pros and Cons of Refactoring Code

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

The Advantages of Refactoring Stock Analysis

The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run. Attached below are the screenshots that indicatePros and Cons of Refactoring Code
Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

The run time for our new analysis.
![image](https://user-images.githubusercontent.com/95595378/149650669-77338b3a-eb24-4be0-8aab-c520c1f2170e.png)
![image](https://user-images.githubusercontent.com/95595378/149650674-b4d9d4ab-2745-40ee-b7be-fc9819052606.png)

