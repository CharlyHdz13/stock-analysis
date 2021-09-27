# Analysis on renewable energy companies' stocks

## Overview of the Project

Today exist various renewable energy companies, which are attractive for investrors to invest in. Mainly because of their huge growth in the recent years. Therefore this project consists on analysing the stocks of these companies and determine which ones are give the best bang for your buck. Due to the inmense amount of data the data set has, it is easier to automate this proces via a VBA code.

## Results

Two codes with the same funcionality were made to analysis the find the starting price, ending price and volume of the stocks for twelve different companies. 

The second code is a refactored of the first and both can be found in the VBA_Challenge.xlsm file.

### Stock Performance

The following image shows the stocks performance in 2017:

![Stocks_2017](https://user-images.githubusercontent.com/89402038/134991351-a100a13e-234f-4faa-bc91-0f64a2434414.png)

The following image shows the stocks performance in 2018:

![Stocks_2018](https://user-images.githubusercontent.com/89402038/134991389-9b7716df-9cde-4ed5-8e01-0dd5bc676736.png)

From the analysis I can say that the stocks of most of the comapnies had a great 2017, with the exception of TERP. Nevertheless, in 2018 most of the companies had a huge drop in their return, except for ENPH and RUN. ENPH had a little dip in their return percentage, but remained positive. On the other hand, RUN had a huge boost in their 2018 return percentage. 

I would recommend bring more data into the data set to check mor years in order to have a better understanding of where are the stocks of this comapnies being headed to.

### Code Performance

As mentioned previously the second code is a refactored version of the first code. The first code is divided into two subroutines: the analysis and formatting of the cells. And the performance was the following:

![VBA_Challenge_2017_Code1](https://user-images.githubusercontent.com/89402038/134992295-ae81beb1-4425-4a04-b75f-7efee9505525.png)
![VBA_Challenge_2018_Code1](https://user-images.githubusercontent.com/89402038/134992300-1a75afdf-11e0-4d33-a160-d6b5e1563c90.png)

The second code has the analysis and formatting into one only subroutine. And the performance was the following:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/89402038/134992458-8d4c0d72-c120-44fc-bb0f-832ddafc7b3e.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/89402038/134992464-3b59cee3-eed1-4b68-bc15-2700059c6be1.png)

## Summary
- What are the advantages or disadvantages of refactoring code?
  
  Some of the advantages of refactoring the code is that you might be able to reduce the amount of resources your computer requires to do the task. One of the disadvantages is that it may make your code more complex and difficult to understand for another person, if it is not well commented. On this project the refactoring of the code made the running times longer but instead we have the analyising part and formatting part all in a subroutine.
  
- How do these pros and cons apply to refactoring the original VBA script?
  
