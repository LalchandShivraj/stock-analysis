# Refactor VBA Code and Measure the Performance of the Greenstock Analysis

## Overview: Write code to go thru the data in one pass instead of looking up for tickers one at a time.

1. Create arrays to store the volumes and starting and ending prices.
2. Create a for loop to go thru the rows of data.
3. If the for loop accumulate volume and starting and ending price for each ticker.
4. Output the results by using a for loop to loop thru the arrays.
5. Create buttons, to run the code.

## Resources
- Data Source: Excel
- Software: Excel (Office 365)
- Output: Excel and screen prints to show performance.

## Results

The analysis of the code Refraction show that:

- The performance of running the code against the 2017 data improves from about .67 to .096 seconds. 
 ![VBA_Challenge_2017](https://user-images.githubusercontent.com/78666055/111034937-ebb21c80-83e5-11eb-9761-8732efb7f5ff.png)

- The performance of running the code against the 2018 data improves from about .75 to .097 seconds.
![VBA_Challenge_2018](https://user-images.githubusercontent.com/78666055/111034961-04bacd80-83e6-11eb-80b0-d1daa8c65bfd.png)


## Summary
- By using the arrays to store results while processing the data in only one pass, proves to improve the performance.
- Further, by refactoring, it gives one a chance to "clean-up" the coding, adding comments/documentation as needed. This should make it easier to maintain.
- Refactoring can be time consuming and there in no guarantee of signifcant enough performance improvement to justify.
