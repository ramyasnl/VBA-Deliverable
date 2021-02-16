# VBA-Deliverable
#Overview of Project:#<br/>
We are analyzing the stocks to chose the best for Steve's parents .<br/>
Steve wants to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Steve's parents are interested in him DQ's stock, so we start with DQ.<br/>
We want to know how DQ performed in 2018. One way to measure this is to calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year.<br/>
We are moving out of DQ since it did not give good returns to all stocks analysis to find the best .And we are refactoring the code to make it work more efficient .<br/>
#Challenges <br/>
i.**Statement which helped us to navigate through the spreadsheet<br/>**
lastRow = Cells(Rows.Count, "A").End(xlUp).Row<br/>
Cells(Rows.Count, "A") goes to the bottom cell in column A, which may extend past the last row of data in the sheet.<br/>
.End(xlUp) is the same as pressing END and then the up arrow in Excel, which will go to the last cell with data in column A. We use this to move from the bottom of the sheet to the last row of data.<br/>
.Row returns the row number.<br/>
 ii.**Setting up the output Spread Sheet<br/>**
  Calculate the yearly return of  stocks<br/>
 1.	Make the title in cell A1 "All Stocks (2018)."<br/>
 2.	Add three columns with the following headers:<br/>
	  Ticker<br/>
	  Total Daily Volume<br/>
	   Return <br/>
 3.Create a"tickerIndex" Variable<br/>
 4.Create three output Arrays "tickerVolumes"- Long ,"tickerStartingPrice","tickerEndingPrice".-Single <br/>
 5.Create a for loop to initialize TickerVolumes<br/>
 6.Create a for loop to loopover all the rows ,i= 2 to RowCount<br/>
 7.Calculate the tickerVolumes<br/>
 8.Write if -then to check the current row is the first row with the selected "tickerIndex""tickerStarting price",<br/>
   similarly another if- then to check the last row for tickerEndingPrice ,<br/>
 9. Create a for loop to loop through arrays to output "ticker","Total Daily Volume"and "Return"<br/>
 
 #  Result Analysis<br/>
https://github.com/ramyasnl/VBA-Deliverable/blob/main/2021-02-14%20(2).png<br/>
https://github.com/ramyasnl/VBA-Deliverable/blob/main/2021-02-14%20(1).png<br/>
##In the year 2017<br/>
The ticker TERP has the least return which is -7.2% while DQ has the maximum return 199.4% <br/>
##In the year 2018<br/>
The ticker RUN has the maximum return of -39.7% while DQ has the minimum of -63%<br/>
##Refactoring<br/>
Advantages<br/>Refactoring improves the design of software, makes software easier to understand, helps us find bugs and also helps in executing the program faster.<br/>
Also it can changes the way a developer thinks about the implementation when not refactoring. <br/>
Disadvantages<br/>While refactoring we should document the changes we have done to the original program, and should be commented  to avoid confusion.






@ramyasnl
