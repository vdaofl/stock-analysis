# **stock-analysis with VBA**

## **Overview of Project**
**_Background:_** Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, we will edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that we did in this module. Then, we’ll determine whether refactoring the code successfully made the VBA script run faster. Finally, we’ll present a written analysis that explains our findings.

### **Results**
**_Results Before:_** Before refactoring the code the resulting times to execute were 0.169688 seconds for 2017 data and .05703125 seconds for 2018 data

![2017 before refactoring](https://utoronto.bootcampcontent.com/smacpherson/stock-analysis/-/raw/master/Resources/2017_before_refactoring.png)
![2018 before refactoring](https://utoronto.bootcampcontent.com/smacpherson/stock-analysis/-/raw/master/Resources/2018_before_refactoring.png)


**_Results After:_** After refactoring the code the resulting times to execute were 0.109375 seconds for 2017 data and 0.1054688 seconds for 2018 data

![2017 after refactoring](https://utoronto.bootcampcontent.com/smacpherson/stock-analysis/-/raw/master/Resources/VBA_Challenge_2017.png)
![2018 after refactoring](https://utoronto.bootcampcontent.com/smacpherson/stock-analysis/-/raw/master/Resources/VBA_Challenge_2018.png)

**_Conclusion:_** Refactoring the code successfully made the VBA script run faster.

## Summary

### What are the advantages or disadvantages of refactoring code?   

#### Some advantages:
- Can make the code more efficient—by taking fewer steps
- Can use less memory
- Can improve the logic of the code to make it easier for future users to read

#### Some disadvantages:
- Can be risky when the application is big by introducing bugs
- Can be risky when existing code doesn't have proper test cases to find introduced bugs
- Can be risky when developers do not understand the original code
- Can takes extra time and budget that might not be available

### How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original VBA script resulted in a more efficient program that resulted in faster run times. In this exercise since the data set was not too large the performance increase was minimal and when wieghed against the extra time if took to refactor the original code not of much value. Hours of time were spent to save a fraction of a second in exectution time. If the data set were much larger and executed many times only then might the refactoring have been worth the time to do.  
