# VBA_Challenge_Stock_Analysis

## OverView of the Project

### Refractor challenge to loop through stock market data set.
We delivered Steve a workbook where at the click of a button he can analyze an entire dataset to compare stocks between two years, 2017 and 2018. We are now wondering if we can expand the dataset to include the entire stock market over the last few years. Would this script work for thousands of stocks? And if it does work would it take a long time to execute?

### Result:

#### Examples of Code:

In the following code we set up the variable "Ticker/Tickers" to correspond with the different stocks we would be running through.  Then we set up the RowCount, tickerIndex, tickerVolume, tickerStartPrice, and tickerEndingPrice.

<img width="647" alt="Tickers and set up" src="https://user-images.githubusercontent.com/111904266/196484435-d62de04b-bd2b-4d2c-880d-a26c798da599.png">


We then ran this code through the dataset finding the Volume, Start Price, and End Price for each ticker.

<img width="835" alt="Script to run through data" src="https://user-images.githubusercontent.com/111904266/196484800-23d1e631-af8c-47b5-8e45-2bcb118e04d3.png">

the following code is the data filling out a table on a worksheet when the user selects the year.

<img width="734" alt="Output the data " src="https://user-images.githubusercontent.com/111904266/196485015-66e7340f-43b5-4bd8-aa04-8578a04a9d89.png">


#### Execution times of original script

![Green_Stocks_2017](https://user-images.githubusercontent.com/111904266/196477943-de4546ad-45e8-44f9-8f7e-092e2c6d416b.png)

![Green_Stocks_2018](https://user-images.githubusercontent.com/111904266/196478046-a30b3aa5-85e6-4f3f-8bfd-6fd57df92209.png)


#### Execution times of refactored script 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111904266/196478103-c1f291e4-cc22-4788-8999-63c79e037f1b.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/111904266/196478127-b3b6fe71-27ba-4a52-8b4b-70b1e4fedae2.png)


### Summary:

#### Would this scrip work for thousands of stocks?



#### Did the Refractoring successfully run the script faster?


#### What are the advantages of refactoring code?


#### What are the disadvantages of refactoring code?

#### How do these pros and cons apply to refactoring the original VBA script?


