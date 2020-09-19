# VBA Stock Analysis

## Overview of Project 

The purpose of this analysis was to assist a client in obtaining and comparing the *Total Daily Volume* and *Yearly Return* for several different companies within a dataset of green energy companies.  The client was particularly interested in the data for DQOS ("DQ") but wanted to compare performance for each of the other comapanies in the dataset as well.  Using a VBA Macro, the client is able to enter the year they wish to analyze and can then compare the results to determine which stocks performed better than others and ultimately, which stocks they'd like to invest in.  By using VBA rather than a manual process to make the calculations, the analysis is guaranteed to be executed the same way each time it is performed (ensuring consistency and accuracy), it is completed much quicker than if it were done manually and it eliminates the risk of manual errors.

In addition to obtaining and comparing the outputs for the dataset provided by the client, the client also wanted to be able to expand thier dataset to the entire stock market and run the same analysis.  To make this possible, the VBA script needed to be refactored to make it more efficient by taking fewer steps, using less memory and improving the logic of the code.

## Results 

The results of the analysis showed that most of the stocks in the dataset performed better in 2017 than they did in 2018 as can be seen in the tables below.  Two exceptions to this trend were the stocks with tickers "RUN" and "TERP" which both had higher returns in 2018.  For DQOS, which the client had particular interest in, the data shows that the stock had a total daily volume of 35.8 million in 2017 and 107.9 million in 2018.  Despite the increase in total daily volume, the return for DQOS was lower in 2018 (-62.6%) than it was in 2017 (199.4%).

<img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/All%20Stocks%202017.png" height=300 width=300> &emsp; &emsp; &emsp; &emsp; <img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/All%20Stocks%202018.png" height = 300 width = 300>

After writing the initial VBA script for analyzing the dataset provided, the code was refactored to improve the logic flow, reduce the number of steps required and make the code more efiicient.  The basic logical flow of the VBA script to analyze the dataset provided is as follows:

First, knowing that the run time needed to be measured to compare to the refactored script, timer variables were initialized.  However, before starting the timer, an input box was needed to allow the user to be able to input the year of the data they would like to analyze.  Once the user enters the year and presses Enter, the timer is started and the output worksheet for the analysis is formatted with a title and headers.  See script below.

![Formatting Output Sheet](https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Formatting%20Output%20Sheet.png)

Next, the ticker and output arrays/variables were initialized and the RowCount was defined to determine the last row number in the worksheet that contains data.  One difference in this section between the original code and the refactored code is that refactored code used arrays for the output values rather than single values.  To loop through the array values and access the correct index accross the different arrays, a tickerIndex was defined and initialized to zero as shown below.

&emsp; &emsp; &emsp; &emsp; _***Original VBA Script***_ &emsp; &emsp; &emsp; &emsp;  &emsp; &emsp; &emsp; &emsp;  &emsp; &emsp; &emsp; &emsp; &emsp; &emsp; &emsp; &emsp;***Refactored VBA Script***
<img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Initial%20Initializing%20Arrays.png" width=525>
<img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Initializing%20Arrays.png" width=375>

After the variables were initialized, loops were used to iterate through the data and perform the analysis.  In the original script, a For loop was used to iterate through each of the tickers.  A nested loop then went row by row through each line of data for each ticker.  If the ticker in the first column of the data matched the ticker culled out in the loop, calculations were performed to find the *Total Daily Volume* and *Return*.  In the refactored data, a For loop was used to interate through each tickerIndex.  However, instead of going through the entire set of data for each tickerIndex, the nested inner loop was set up to run the calculations for the specified ticker until that ticker no longer matched.  The ticker was then increased by one, moving the calculations to the next value in the array and the calculations wer repeated for that next ticker.  By using the array approach, the script only has to run through each data row one time, rather than 12 times as was required by the initial script.  The two scripts are shown below.

***Initial VBA Script***

![Initial Loop Analysis](https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Initial%20Loop%20Analysis.png)

***Refactored VBA Script***

![Loop Analysis](https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Loop%20Analysis.png)

Another difference between the two codes shown above is that the initial script included the outputs within the loop and the refactored code didn't.  In the original code, because the analysis was done throughout the entire data set for each ticker, the output could be included in the loop for each ticker.  For the refactored code, as soon as the analysis reached the final row for the ticker, the tickerIndex was increased by one and the analysis began for the next ticker.  This meant that if the output was placed inside the initial For loop, it would try to give an output for tickerIndex=12 which does not exist.  Therefore, the output had it's own separate loop where the tickerIndex only ranged from 1 to 11.  See the output loop below.

![Output Loop](https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Output%20loop.png)

Finally, the last steps of the script involved formatting the output values to make them easier to read, using conditional formatting to fill the cell red or green depepnding on whether the return was positive or negative and then stopping the timer when the analysis was complete.

![Formatting Results](https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Formatting%20Results.png)

The message box that is added at the end of the code tells the user how long it took to run the script.  We can see from the readouts that the initial script had run times of about 0.7 seconds while the refactored script had run times of about 0.1 seconds.  This shows that refactoring the data resulted in the script being able to run 7 times faster than the original. 

***Initial VBA Script***

<img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Original%20VBA%20Code%202017.png"> &emsp; &emsp; &emsp; &emsp; <img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/Original%20VBA%20Code%202018.png">

***Refactored VBA Script***

<img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/VBA%20Challenge%202017.png"> &emsp; &emsp; &emsp; &emsp; <img src="https://github.com/whitneylosinski/stock-analysis/blob/master/Resources/VBA%20Challenge%202018.png">

## Summary

The original VBA script that was created to perform the analysis was perfectly capable of correctly analyzing the dataset that was given.  However, by refactoring the code, it was able to be restructured to run faster and more efficiently without changing the external outcome.  This allows the final script to be able to be used on a larger dataset such as data for the entire stock market if the client chooses to do so.

Using the refactored VBA script, the yearly stock data can easily be compared to see that the majority of the stocks in this particular dataset had better returns in the year 2017 than 2018.  For the DQ stock, which the client has particular interest in, the data shows that it had a positive return of almost 200% in 2017 but a negative return of nearly 63% in 2018.  Because this stock followed the same trends of the yearly data as most of the other stocks, additional data would need to be analyzed to determine if this would be a good stock to invest in.

### *Advantages and Disadvantages of Refactoring Code*

Refactoring code has many advantages and also some disadvantages.  By using the process of refactoring, code can often be simplified and better understood by both the writer of the script and also by the end user/future users of the script.  Often times, the code can be reduced to fewer steps through improved logic which also leads to using less memory and faster analysis speeds.  A refactored, simplified code is easier to enhance and maintain if used again in the future.

Some disadvantages to refactoring are the amount of time it takes and the risk of creating errors when simplifying or improving the logic of the code.  Refactoring can be extremely time consuming and very frustrating if the logic is misunderstood or simplified incorrectly.  Because of the amount of time refactoring requires, it can be very costly and could lead to missed deadlines.  There is also a risk of introducing errors into the code which can be an extremely frustrating, time consuming and costly mistake.  

### *Pros and Cons of Refactoring the Original VBA Script*

In this particular project, refactoring the original VBA script led to faster run times and allows for the client to be able to run their analysis on the larger dataset of the entire stock market.  This was accomplished by using arrays for the analysis and output data rather than single values.  Instead of running the outer loop and iterating through the entire data set for each ticker like in the original script, the refactored code only ran the outer loop one time, while the inner loop completed the entire analyis in the arrays by iterating through the entire data set only one time.  The final code was also updated with more descriptive comments, creating an easier to understand process flow which will be advantageous for future users.  The refactoring was frustrating and did take several hours to complete due to some bugs that were created in the process, but it was worth the extra work in order to provide the client with a quick, easy-to-use analysis.

