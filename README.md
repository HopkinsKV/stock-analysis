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
Include original:
Image with 2017 data
image with 2018

Include refactored:
Image with 2017 data
image with 2018

Refactoring this code allowed for a roughly 80% decrease in run time.

Source Data: VBA Challenge
Note: the above link contains both the original and refactored code.


## Summary

### Pros
Refactoring code can increase efficiency. It should streamline the process, allowing future viewers of the code to more easily read it.

Refactoring this code provided a drastic decrease in run time. This will allow for larger sets of data to be compiled without slowing down the process.

### Cons
The disadvantage of refactoring code is the increased debugging process.

This script in particular still requires manually defining each item in an array, which is not scalable. If new companies need review, they must be added to the array by hand. The references to this array will also need to have the for loops updated to match the new array amount.

Lastly, the refactored code may use more memory while it increases efficiency.
