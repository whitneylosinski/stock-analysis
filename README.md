# VBA Stock Analysis

## Overview of Project 

The purpose of this analysis was to assist a client in obtaining and comparing the *Total Daily Volume* and *Yearly Return* for several different companies within a dataset of green energy companies.  The client was particularly interested in the data for DQOS ("DQ") but wanted to compare performance for each of the other comapanies in the dataset as well.  Using a VBA Macro, the client is able to enter the year they wish to analyze and can then compare the results to determine which stocks performed better than others and ultimately, which stocks they'd like to invest in.  By using VBA rather than a manual process to make the calculations, the analysis is guaranteed to be executed the same way each time it is performed (ensuring consistency and accuracy), it is completed much quicker than if it were done manually and it eliminates the risk of manual errors.

In addition to obtaining and comparing the outputs for the dataset provided by the client, the client also wanted to be able to expand thier dataset to the entire stock market and run the same analysis.  To make this possible, the VBA script needed to be refactored to make it more efficient by taking fewer steps, using less memory and improving the logic of the code.

## Results 

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.


## Summary

The original VBA script that was created to perform the analysis was perfectly capable of correctly analyzing the dataset that was given.  However, by refactoring the code, it was able to be restructured to run faster and more efficiently without changing the external outcome.  This allows the final script to be able to be used on a larger dataset such as data for the entire stock market if the client chooses to do so.

Using the refactored VBA script, the yearly stock data can easily be compared to see that the majority of the stocks in this particular dataset had better returns in the year 2017 than 2018, with the exception of the RUN and TERP stocks, which performed better in 2018.  For the DQ stock, which the client has particular interest in, he data shows that it had a positive return of almost 200% in 2017 but a negative return of nearly 63% in 2018.  Because this stock followed the trends of the yearly data for most of the other stocks, additional data would need to be analyzed to determine if this would be a good stock to invest in.

### *Advantages and Disadvantages of Refactoring Code*

Refactoring code has many advantages and also some disadvantages.  By using the process of refactoring, code can often be simplified and better understood by both the writer of the script and also by the end user/future users of the script.  Often times, the code can be reduced to fewer steps through improved logic which also leads to using less memory and faster analysis speeds.  A refactored, simplified code is easier to enhance and maintain if used again in the future.

Some disadvantages to refactoring are the amount of time it takes and the risk of creating errors when simplifying or improving the logic of the code.  Refactoring can be extremely time consuming and very frustrating if the logic is misunderstood or simplified incorrectly.  Because of the amount of time refactoring requires, it can be very costly and could lead to missed deadlines.  There is also a risk of introducing errors into the code which can be an extremely frustrating, time consuming and costly mistake.  

### *Pros and Cons of Refactoring the Original VBA Script?*

In this particular project, refactoring the original VBA script led to faster run times and allows for the client to be able to run their analysis on the larger dataset of the entire stock market.  This was accomplished by using arrays for the analysis and output data rather than single values.  Instead of running the outer loop and iterating through the entire data set for each ticker like in the original script, the refactored code only ran the outer loop one time, while the inner loop completed the entire analyis in the arrays by iterating through the entire data set only one time.  The final code was also updated with more descriptive comments, creating an easier to understand process flow which will be advantageous for future users.  The refactoring was frustrating and did take several hours to complete due to some bugs that were created in the process, but it was worth the extra work in order to provide the client with a quick, easy-to-use analysis.

