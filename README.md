# stock-analysis
VBA_Challenge
  ## Overview of Project: Explain the purpose of this analysis.
  Performing financial analysis using VBA Programming language on 2018 and 2017 Green Energy stocks.In this project, I refactored the previous subqueries that performed           analysis on different stock tickers using nested loops and arrays. By introducing TickerIndex which contains four different arrays (Ticker, TotalVolume, TickerStartingPrices,   TickerEndingPrices) we were able to reduce the amount of computations required for this subquery which resulted in lower time for computation.
  ##Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the           refactored script.
  Please refer to attached snippets of (4) runtime images in the resources folder. As indicated in the pictures, the initial runtime for 2017 and 2018 annual year were           respectiveley 1.113281 and 1.097656 seconds. In Module 2, I created a refactiring subquery using tickerindex to reduce the computation time. As shown in pictures, the runtime   for 2017 and 2018 after refactoring, is respectively .1796975 and .1875 seconds which is a considerable faster runtime.
  ##Summary: In a summary statement, address the following questions.
    ###What are the advantages or disadvantages of refactoring code?
    I start with the advantages. I believe, working to refactor a code that you have already developed and run, enables you to dive dipper on the code and better understand the    workflow and the areas that can be improved. For this educational challenge, refactoring this code, helped me better understand how nested loops and arrays actually work in    the program. The refactored code is easier to understand and fresher.
   Disadvantage of the refactoring process is being time consuming and very hard to predict. you do not know for the fact how much time it requires to refactor a code and          there are possibility of running into more errors.
    

    ###How do these pros and cons apply to refactoring the original VBA script?
    By revisiting the code and subqueries to see if there has been a variable without dedication of data type "DIM". By using Dim and allocating data types to variables, we         will have a more efficient program and lessram usage. Also by looking at nested loops again, they will be able to brainstorm for more efficient loops and functions. For the     Stock Analysis challenge, I my self believe that it could have been more efficient to use date factot for calculating the startingPrice and endingPrice by finding the price     in January 1st and December 30th.
