# VBA-challenge

Module 2 Challenge – Excel VBA Scripting: The VBA of Wall Street

Objective:
This project will apply the skills learnt in Module 2 – VBA Scripting for Excel. The skills will establish the foundation of programming logic formulation. This project will be submitted to GitHub/Gitlab. This will help developed the skills for project deployment and versioning.

Project Details:
This project will analyze the stock market data using VBA script. The stocks data were saved in multiple worksheets with one sheet for each year. Each sheet has multiple rows containing the following stocks daily trading data:
1.	ticker symbol – stocks ticker symbol sorted in ascending order
2.	trading date – trading date sorted in ascending order
3.	open price 
4.	highest price 
5.	lowest price 
6.	closing price 
7.	total volume

The script will loop though the list of stocks data and will create a summary table on the same worksheet as the raw data. The summary table will contain the following data for each year:
1.	Ticker 		– unique list of ticker symbols
2.	Yearly Change – Price yearly change from the open price of the first trading date to the close price of the last trading date for each unique ticker symbols. The value must be in number format with 2 decimal places.
3.	Percent Change – Percent price yearly change from the open price of the first trading date to the close price of the last trading date for each unique ticker symbols. The value must be in percent format with 2 decimal places i.e., “00.00%”.
4.	Total Stock Volume – Total stock volume for each unique ticker symbols

The script will create an additional summary table on the same worksheet as the raw data. The summary table will contain the following data for each year:
1.	Greatest % Increase – The ticker symbol with the greatest % increase value for the year 
2.	Greatest % Decrease – The ticker symbol with the greatest % decrease value for the year
3.	Greatest Total Volume – The ticker symbol with the greatest total volume for the year

The script will apply the conditional formatting to the data in “Yearly Change” and “Percent Change” column in the first summary table. It will highlight the positive change in green and the negative change in red.

The script will run on all sheets containing the stocks yearly data. It will create the two summary tables mentioned above.

Project Submission:
The project will be submitted by uploading the following files to GitHub/Gitlab repository called “VBA-challenge”:
1.	Screenshots of the results in .png format
2.	VBA script file in .vbs format
3.	README file in .md format
