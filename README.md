# Stock Analysis with VBA

## Overview of Project
The purpose of this project is to analyze performance of several stocks in the green energy space to analyze their total daily volume and yearly return.  A VBA macro and buttons were added to simplify the analysis process and make it easily repeatable.

To accomplish this a macro that makes use of arrays, loops, and conditional logic was used to aggregate total daily trade volume data and its change in valuation.  Formatting was applied to make the results more readable. In the end, the code was refactored to improve efficiency and usability by minimizing the number of loops required to process the data and also by adding buttons to allow the user to quickly and easily repeat their analysis for any year.
  
[Green Stock Analysis](https://github.com/curt0230/stock-analysis/blob/main/Submission/VBA_Challenge.xlsm)

## Results
During this analysis, the focus was on the performance of stocks of approximately a dozen green energy companies for the years 2017 and 2018.  The Total Daily Volume (number of trades per day), was summed by Ticker, and then the yearly return was represented as a percent of the starting and ending price for the selected term.  This represents how much the stockholder has historically gained or lost, with green values representing gains and the red representing losses.

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

The original code set used nested looping logic in order to complete the analysis.  In the original construct below, the application evaluates each row, comparing its ticker value to the prior row to determine if it represents the start of a company's data set and if it is the opening price is stored.  Similarly, the ticker is compared to the next row to determine if it represents the end of the company's data set and in that case the closing value is stored.  Also, the day's trading value is added to a running sum.  This process is repeated for each and every company in the index of the outer loop.  Sorting the data by Ticker in memory was not necessary since it is already stored this way in the 2017 and 2018 worksheets.

![VBA_Challenge_Original_Loops](/Resources/VBA_Challenge_Original_Loops.png)

While accurate, this is heavy processing work.  For just twelve subject companies with trading data already aggregated to the year level, analysis took nearly one second to complete:

![VBA_Challenge_2017_Before_Refactoring.png](/Resources/VBA_Challenge_2017_Before_Refactoring.png) ![VBA_Challenge_2018_Before_Refactoring.png](/Resources/VBA_Challenge_2018_Before_Refactoring.png)

While the original logic achieved accurate results it was clear performance could be a problem on larger data sets.  To address this issue the code was refactored in order to make it more readable as well as more performant.  This construct requires less memory and CPU time for processing, speeding up the application and therefore creating a better user experience.  The savings is realized because the loop is exited when the end of each company's data set is reached rather than continuing to process rows.

![VBA_Challenge_Refactored_Loops](/Resources/VBA_Challenge_Refactored_Loops.png)

As you can see, this modification yielded approximately a 85% reduction in execution time: 

![VBA_Challenge_2017_After_Refactoring.png](/Resources/VBA_Challenge_2017_After_Refactoring.png) ![VBA_Challenge_2018_After_Refactoring.png](/Resources/VBA_Challenge_2018_After_Refactoring.png)

## Summary
Implementing application code with an eye on performance, readability, and workflow of application code should always be a primary goal.  Refactoring may be necessary over time as applications are modified, especially if system performance issues are identified either via monitoring or by users reporting slow downs.  While time consuming and sometimes tedious, refactoring often pays off in the end.  In this case we achieved a significant improvement in performance and also made the code easier to follow if ever the logic needs to be revisited, for example to handle unsorted data.