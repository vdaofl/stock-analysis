# **stock-analysis with VBA**

## Overview of Project: Explain the purpose of this analysis.
**_Background:_** Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, we will edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that we did in this module. Then, we’ll determine whether refactoring the code successfully made the VBA script run faster. Finally, we’ll present a written analysis that explains our findings.

### **Results Before and After**
**_Results Before:_** Before refactoring: 0.169688 seconds for 2017 data and .05703125 seconds for 2018 data

![2017 before refactoring](https://i.imgur.com/4UFxydv.png)
![2018 before refactoring](https://i.imgur.com/WkkzH9U.png)


**Results After:** After refactoring: 0.109375 seconds for 2017 data and 0.1054688 seconds for 2018 data

![2017 after refactoring](https://i.imgur.com/BnLcbR0.png)
![2018 after refactoring](https://i.imgur.com/irtzbNO.png)

**Conclusion:** After refactoring the code, the script runs faster.

## Summary

### What are the advantages or disadvantages of refactoring code?   

#### Advantages:
- The code becomes more efficient. Less steps.
- Uses less memory.
- Easy to read for future users.

####  Disadvantages:
- Bugs can be introduced when application becomes too big.
- Risky when existing code doesn't have proper test cases to find introduced bugs
- Risks involved if developers don' understand the original code
- This takes extra time, and sometimes that it not available.

### How do these pros and cons apply to refactoring the original VBA script?

After refactoring the code, the script ran only marginally faster. However, to make this happen, it took hours. In a business sense, the benefits do not outweigh the costs (time and money). The only scenario where refactoring the code may be advantageous would be if the data set were much larger.
