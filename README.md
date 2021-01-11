# An Analysis of the green stocks performance in 2017 and 2018
## Overview of Project
### Purpose
The purpose of this project is using VBA code to analyse and compare the performance of 12 stocks in 2017 and 2018. The result will help to build the fund.

## Results
### Stock performance in 2017 and 2018
In 2017, 11 out of 12 stocks had positive return, and DQ had the most return, 199.4%.

![2017_Result](Resources/2017_Result.png)

In 2018, only 2 out of 12 stocks had positive return, and ENPH had the most return, 84%. 

![2018_Result](Resources/2018_Result.png)

As shown above, the return of stocks in 2018 was systematically worse than that for 2017. And ENPH had the most stable performance, it had realtively good return in both 2017 and 2018.

### Execution times of VBA scripts
The execution time for the original script was aroung 0.6s, and the execution time for the refactored BVA script was 0.125 secound as showned below.

![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)


## Summary
### Advantages and disadvantages of refactoring code in general 
Advantages:
1. increase readability of the code
2. increase the maintainbility
3. increase the extensibility
4. increase the efficiency in long term

Disadvantages:
1. Error or mistake made during refactoring may influence the whole structure


### Advantages and disadvantages of the original and refactored VBA script
- the original script didn't use array, less variable were initialized.
- the refactored script was more organized and even if there is more years of data, it can still process without chaning anything in the script.
