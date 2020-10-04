# stocks-analysis
##  Overview of Project
Steve asked me to help him analysis a stock called Daqo from thousands of datas for his parents to see the transaction volumes and return with Excel VBA. After this analysis we found out that  Daqo might not be the best option for Steve's parents to invest in, so we analyzed multiple stocks to find some better choices for them. 
### Purpose and background  
The Purpose of this project is to refactor the code to loop through all the data one time in order to collect the same information. We’ll determine whether refactoring the code successfully made the VBA script run faster. 
##  Results
In order to make the code runs more efficient, I created a new variable (tickerIndex) and set this variable tickerIndex to access the correct index across the four different arrays, the tickers array and the three output arrays tickerVolumes, tickerStartingPrices, and tickerEndingPrices. By doing it in this way, the analysis can complete faster than before. Below please find the Refactor code, Original code and the different run times:

#### Here are the run times with original code
![2017_original time](https://user-images.githubusercontent.com/71739110/95020578-99363c80-069e-11eb-8838-7844248b8a29.png)
![2018_original time](https://user-images.githubusercontent.com/71739110/95020580-99ced300-069e-11eb-8095-8b22ae529259.png)

#### Here are the run times with refactored code
![2017_updated time](https://user-images.githubusercontent.com/71739110/95020579-99ced300-069e-11eb-9674-77bb2b0a1c46.png)
![2018_updated time](https://user-images.githubusercontent.com/71739110/95020581-99ced300-069e-11eb-8f40-8fc4956f5f98.png)

#### Refactor Code

#### Original Code

##  Summary

### There is a detailed statement on the advantages and disadvantages of refactoring code in general 

### There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script 
