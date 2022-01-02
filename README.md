# VBA Challenge

# Overview
The purpose of this project is to help Steve, a recent finance graduate, help his parents with their stocks. They invested all of their funds into DAQO New Energy Corp without doing any research on the company. He wants us to help analyze a handful of green energy stocks in additional to DAQO stocks to help diversify their funds. By using VBA, a programming language for Excel, I will be helping him create macros to help automate tasks and reviewing the data that Steve has provided us. This will alsp help him analyze new data in the future if he needs it. 

## Results
I created a code to help Steve figure out the return of different green energy stocks to see whether or not some would be worth looking into for his parents for them to diversify their funds. The code tells Steve what tickers had a positive or negative return for that year. After running the code, Steve will be able to compare data from 2017 vs 2018 easier than looking at each indiviual line from his dataset. By doing this, Steve can see that ENPH and RUN has positive returns for both 2017 and 2018. This indicates that these are two possible companies that Steve can advise his parents to invest into. See the two images below for visual examples of this data.

![This is an image](https://github.com/TracyKien/stock-analysis/blob/main/Resources/Return%20Status%20for%202017.PNG?raw=true)

![This is an image](https://github.com/TracyKien/stock-analysis/blob/main/Resources/Return%20Status%20for%202018.PNG?raw=true)


## Refactoring the Original Code
Now Steve can take this information that I provided him and use it on future analyses. However, if he wanted to run the code on a much larger dataset than what he provided us to analyze, it could potentially take longer to excute. I had to refactor the original code to make it more efficient and run faster for Steve. 

If you look at the following images, you can see that the orignal code took 1.234375 seconds to exceute theough all of the data for 2017. By refactoring the code, the new runtime was only 0.2109375 seconds, which is significantly faster.

![This is an image](https://github.com/TracyKien/stock-analysis/blob/main/Resources/VBA_Challenge_2017%20Original%20Time.PNG?raw=true)
![This is an image](https://github.com/TracyKien/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

The same proved true for the 2018 data.

![This is an image](https://github.com/TracyKien/stock-analysis/blob/main/Resources/VBA_Challenge_2018%20Original%20Time.PNG?raw=true)
![This is an image](https://github.com/TracyKien/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG?raw=true)



## Summary
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

