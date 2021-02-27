# stock-analysis

Overview:

  The purpose of this analysis is to evaluate how 12 different green energy stocks performed in 2017 and 2018 in order to help my friend Steve give his parents advice on which       stock to invest in.
  
Results:

 Attached is a screenshot of the final analysis for the 2017 stocks [https://github.com/azarowj/stock-analysis/blob/main/Resources/Screen%20Shot%202021-02-26%20at%208.16.35%20PM.png] and here is the analysis for the 2018 stocks [https://github.com/azarowj/stock-analysis/blob/main/Resources/Screen%20Shot%202021-02-26%20at%208.17.04%20PM.png]. It is clear that the majority of these stocks had positive returns in 2017, however this was not the case for 2018. Only two stocks had positive returns in both 2017 and 2018. These stocks were ENPH and RUN. It would then make sense for our friend Steve to suggest to his parents to invest in either of these two stocks
 
 The execution times of the original script and the refactored script were extreme. Attached is a screenshot showing the popup when the original code was run for 2017 [https://github.com/azarowj/stock-analysis/blob/main/Resources/Screen%20Shot%202021-02-26%20at%208.06.04%20PM.png] which can be compared to how quickly it ran with the refactored code for 2017 [https://github.com/azarowj/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png]. This is also evident with the speed for 2018 as seen next [https://github.com/azarowj/stock-analysis/blob/main/Resources/Screen%20Shot%202021-02-26%20at%208.06.17%20PM.png] [https://github.com/azarowj/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png]
 
 Below is an example of code from the original analysis
 
 ```
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
 ```
 
 Below is an example of code from the refactored analysis
 
 ```
  Cells(k + 4, 1) = tickers(k)
  Cells(k + 4, 2) = tickerVolumes(k)
  Cells(k + 4, 3) = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
 ```
 
 
Summary:
  1. One major advantage of refactoring the code is the speed that the program runs. One disadvantage is that the refactored code involves more code and therefore takes longer to      write. Combining the code of the stock analysis and the formatting also means that I only need to run one Macro instead of two in order to get the outcome that we are looking      for.
  2. When refactroring the original code, we have to create a new for loop with a new variable for the formatting within the all_stock_analysis code. 
