# VBA-Challenge
### Objective
Develop VBA script that analyzes multiple years of stock market data, and creates a summary table of the stocks' performance for each year.

### Description of the Data: Multi_year_stock_data
The stock data is contained in a workbook called "Multi_year_stock_data", organized by year in worksheets entitled "2014", "2015", and "2016". Each worksheet contains the opening, closing, highest, and lowest price of each stock, as well as its volume, for almost every day of the year. Going down the rows within each sheet, the data is organized first alphabetically based on each stock's ticker name, such that all the rows for each stock are grouped together; then, the data for each individual stock is ordered chronologically, such that the first row for a stock contains the values at the start of the year and the last row for the stock contains the values at the end of the year. 

### What the VBA Script should do
For each year's respective worksheet, the VBA script should loop through all of the stocks and create a summary table containing the following infomation for each unique stock:
- The stock's ticker name 
- The stock's yearly price change, calculated based on its opening price on the first day of the year and its closing price on the last day of the year
- The stock's percent change, calculated by dividing its yearly price change by its opening price on the first day of the year (positive values should be highlighted in green, and negative values in red)
- The stock's total volume, calculated as the sum of the stock's daily volumes

As a bonus, a second summary table should be created listing the ticker name of the stocks with the greatest percent increase, greatest percent decrease, and greatest yearly volume, as well as their respective corresponding values.


