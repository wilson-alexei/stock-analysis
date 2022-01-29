# stock-analysis
Refactoring VBA Code to create a dataset filled with stock information 

## Overview of Project
For this project, we are helping Steve research and expand a dataset that possibly includes the whole stock market. We are given an Excel filled with different stock tickers containing information such as their daily Open and Close price, daily High and Low price, volumes traded, adjusted closing price, and the trading date.  
Now, we need to refactor the code to efficiently loop through all the data and give us the proper result where we can see which stock is profitable or not between the year of 2017 and 2018. We are also attempting to refactor VBA codes to run through the dataset faster.  

##Results 
As we can see from the results, the 2017 stocks performed significantly better than the 2018 stocks. Almost all the 2017 stocks have a positive return except 'TERP' while the 2018 stocks are dominated with negative returns except for two stocks which are 'ENPH' and 'RUN'. With high percentage of returns in 2017, we can speculate that the 2017 market was highly inflated and the bubble eventually popped in the 2018 market as the prices are readjusted back to pre-inflation.  

The purpose of refactoring code is to make the code is more efficient by improving the code processes and simplifying the code for future users. Based on the trials from both codes, original and refactored, the numbers show that the execution times from refactored code are significantly faster than the original code. I uploaded 4 screenshots with 2 from the original code and 2 from the refactored code from both years. For the year of 2017, the Original_Code_2017 ran in 0.30859 seconds while the refactored code on VBA_Challenge_2017.png ran in 0.08984 seconds which is significantly faster. For the year of 2018, the Original_Code_2018 rans in 0.30469 seconds while the refactored code on VBA_Challenge_2018.png ran in 0.08984 seconds which is also significantly faster. I also noticed that the run time for the refactored code is the exact same. 

## Summary 

- What are the advantages and disadvantages of refactoring code? 
The advantages of refactoring code include reducing the complexity for future users to have an easier understanding and potentially saving time and money. Refactored code allows future users such as a new employee to learn faster as the code is slightly simplified. It is also better to show your other analysts/developers in your team so that they can comprehend better. With more simplicity, refactored code has a lower risk of errors moving forward and its functionality can be implemented for other similar uses which will save time and money. 
The disadvantage of refactoring the code is that it could be time consuming since we are not sure how much time it may take, and it could potentially take more time if we are faced with problems due to the level of code complexity. 

- How do these pros and cons apply to refactoring the original VBA script?
I want to start answering this question from the cons. The cons mentioned above are solely based on my experience while working through Deliverable 1. With the provided instructions, I underestimated the complexity of refactoring the code and ended up working on it longer than I expected. I was faced with many errors and bugs that forced me to go through the code from top to bottom again which takes a lot of time and it was very stress inducing. There was even a time where I was confused and clueless on my thought process. However, the pros outweigh the cons. It leads to a better functionality and understanding of the code since there are several uses for it such as sharing it with our teammates or future users and potentially achieve a more efficient running time. From this challenge alone, we can see the positive impact it had on a big enough Excel data. The running time is significantly faster than the original code which will be beneficial for a bigger project and data if we want to extend the functionality of the code.

**Email: wilsonalexei123@yahoo.com**
**LinkedIn: https://www.linkedin.com/in/wilson-alexei/**




