# VBA-challenge
This VBA challenge will loop through several thousand lines of stock data, aggregate the stock values, and output their annual price and percentage change, total volume, and the stocks with the greatest increase and decrease of price change. 
To output aggregated stock data, a for loop was created to run through columns and combine matching stock values. Once combined codnitional statements were created to do the following:
* Combine stock volumes, and output their total yearly volume
* Compare the stock's opening and closing prices to find their yearly change
* compare the yearly change to the opening price to output the stock's annual percentage change
* Highlight by color a positive or negative change in cost.

A secondary for loop was then created to loop through the newly created values, while using conditional statements, and output the following:
  * Read through the Percent change coloumn and compare the greatest percent increase
  * Read through the Percent change coloumn and compare the greatest percent decrease
  * Read through the Total Stock Volumn coloumn and compare values to find the greatest stock volume

Lastly, a 'For Each' function was used to enable the module to run on all sheets within the Excel workbook
