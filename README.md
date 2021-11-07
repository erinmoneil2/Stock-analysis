# Stock-analysis VBA Challange using Excel 

## Overview 

### Purpose 
The purpose of the All Stocks Analysis is to efficiently analyze a data set for twelve stocks over the course of two years to measure which stocks are doing better than others and which ones are worth investing in. In addition to the data analysis, the goal is to refactor the code that sorts through only 12 stocks and make it applicable to run efficiently through thousands of stocks with code that uses less memory, fewer steps, and improves the processing time. 

## Results 

### All Stocks Analysis 
After producing the Total Daily Volume and the return for all twelve stocks, it is easier to see which stocks performed better between the years 2017 and 2018. We know that Steve's parents were pretty set on investing in DQ because of their history with Dairy Queen, but looking at the return for DQ stock, it would not be in their best interest to invest in DQ stocks. If I was Steve, I may recommend my parents investing in other stocks such as ENPH or RUN. 

All Stocks Return 2017
![Screen Shot 2021-11-06 at 8 41 17 PM](https://user-images.githubusercontent.com/92831268/140657573-aa464b33-63a0-4c7b-a78c-19bbf3869e53.png)

All Stocks Return 2018
![Screen Shot 2021-11-06 at 8 41 00 PM](https://user-images.githubusercontent.com/92831268/140657576-95719fac-e28e-4581-a8dd-63f110f72438.png)


### The Code 
In the old code, the for loop has two variables, both i and j. The for i loop loops through the tickers 0 - 11, and the for j loop, loops through the assigned formula, or what we want the loop to do. In this case, get the total volume for the current ticker, get the starting price for the current ticker, and get the ending price for the current ticker. With this first code, what VBA is doing is sifting through all of the tickers one at a time, starting at 1, runs through the for j loop, and then starts back at 1, then to 2, runs through the loop, back to 3, and so on for each assigned ticker. In our VBA challenge, we have eliminated the need to go back to the beginning by creating a tickerIndex. Instead of looping through the entire document and starting back at 0 to find each index, the loop that is created starts with the first index, conducts the same analysis of finding the volume, starting and ending price, and then as soon as it hits the next tickerIndex, it starts the loop over with the next ticker. By doing this, the VBA code runs down the excel sheet one time, simulaneously conducting the daily volume and return analysis without having to start back in the first row of the excel sheet. 

Below is the point in the Green Stocks Analysis where we have both the For i loop, looping through the tickers, and then the For j loop, that once we hit the specified ticker, it loops through the For j loop, and then back to the For i loop to the next ticker. 
![Green Stocks Analysis_For i, For j](https://user-images.githubusercontent.com/92831268/140657282-0a2dde92-35cc-4479-8aec-f66a76de001d.png)

Below is the refactored code that loops through all the tickers down through the excel spreadsheet. 
![VBA Challange_For i](https://user-images.githubusercontent.com/92831268/140657405-76228ed2-638e-4e7c-86af-5c1c3e756145.png)

This refactored code reduced the amount of loops and times the system had to sort through the data. It also reduced the run times from roughly .08 seconds, down to less than .02 seconds. 

2017 Stock Analysis Run Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92831268/140657529-7cc7d5f0-a50b-42fc-b500-212e86f3fb56.png)

2018 Stock Analysis Run Time 
![VBA_Challence_2018](https://user-images.githubusercontent.com/92831268/140657528-d9f0a4e3-76e8-42b2-a082-0467a69a5b61.png)

## Summary 

### Refactorying Code 
While refactoring code does not necessarily change the functionality of the code, it can help to organize the code better, make the code run quicker, and look more readable. This can help to make changes later on if the code is easier to read, more organized and requires less changes and less opportunities to mess up the code. For example, if you assign an iterator to a set, then you only have to change the iterator instead of every number in the case that the entire data set moves cells. You would only have to change the locations in the code where that iterator exists. A disadvantage of refactoring code is that you can potentially add bugs within the refactored code that will not be decected by the debugging tests. This can result in an error in the code, but you being unable to find where the code went wrong. 

### Refactorying VBA Script
One of the obvious advantages to refactoring the original script was the reduced run time to perform the script. The run tmie decreased roughy .6 seconds fromt the original to the refactored script. Other advantage to me personally, was the the refactored script seemed to make more sense. It could have been due to the mulitple repetitions and trying to figure out what worked and what did not, but the final product seemed to be more efficiant and make more logical sense than the original script that we wrote during the module. I think the disadvantage was having to retest the refactored code over and over again. There were many times where pieces of the code worked in the original but not in the refactored script and having to retest the script several times. Refactoring code also brought up other questions of functionality in the case that the data set was mot neatly organized as it was. While I know that we are able to easily sort out the data to have all of the tickers organized in groups, but how would the functionality of the script work if it was not. This script runs to find the first and last ticker within an organized list, but if the tickers were organized in, for example, ascending order, the script would not work the same. 
