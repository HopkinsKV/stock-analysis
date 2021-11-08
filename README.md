# Refactoring Stock Analysis

## Overview of Project

### Background
Analysis requested for Steve's parents to better understand the investment options available. Steve's parents are interested in DQ stock. They believe that a high frequency of trade will reflect an accurate value of the stock. Steve is interested in providing data for other stock options.

Code was previously written to process this information, but needs revision to allow for appropriate scaling.

### Purpose
- The analysis includes a daily volume and yearly return for each stock.
- The macro needs to be available to reuse through multiple years of data.
- Additionally, the macro will include conditional formatting to easily identify stock performance.
- The previously written code will be refactored to improve performance.

## Results
Refactoring this code allowed for a roughly 80% decrease in run time.


Resulting table for 2017:

![image](https://user-images.githubusercontent.com/91762315/140813145-b562b756-886f-4489-a79b-854477441d63.png)


Resulting table for 2018:

![image](https://user-images.githubusercontent.com/91762315/140813087-825ae0ce-e5e8-4df3-837a-0597cd83e4ae.png)

### Original Script
Run time:

![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/91762315/140182676-88700cfd-9f38-451f-ad21-d6f6696c8ca6.png)
![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/91762315/140182679-ea3fe67b-2f7d-4d5d-88b8-e2331babdf26.png)

Loop through rows:

![image](https://user-images.githubusercontent.com/91762315/140814599-a1c3c457-4caa-4783-9504-54723639b4b0.png)


### Refactored Script
Run time:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91762315/140182674-59ad8893-64cb-4e6e-8d54-7078254e79ad.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/91762315/140182678-122c8e1b-23cc-4704-93f8-22b60f8f990f.png)

Loop through rows:
![image](https://user-images.githubusercontent.com/91762315/140814454-88afa7e3-ae39-4fdf-bfbc-039aca557347.png)






Source Data: [VBA_Challenge.xlsx](https://github.com/HopkinsKV/stock-analysis/blob/main/VBA_Challenge.xlsm)

Note: the above link contains both the original and refactored code.


## Summary

### Pros
Refactoring code can increase efficiency. It should streamline the process, allowing future viewers of the code to more easily read it.

Refactoring this code provided a drastic decrease in run time. This will allow for larger sets of data to be compiled without slowing down the process.

### Cons
The disadvantage of refactoring code is the increased debugging process.

This script in particular still requires manually defining each item in an array, which is not scalable. If new companies need review, they must be added to the array by hand. The references to this array will also need to have the for loops updated to match the new array amount.

Lastly, the refactored code may use more memory while it increases efficiency.
