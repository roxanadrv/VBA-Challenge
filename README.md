Overview
This repository contains an Excel VBA script designed to analyze stock data across multiple spreadsheets within a single workbook. The script calculates the yearly change, percent change, and total stock volume for a list of stocks. It also identifies the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

Contents
- Multiple_year_stock_data_RoxanaDarvari.vbs and Multiple_year_stock_data_RoxanaDarvari.bas: The VBA script that performs the analysis.
- alphabetical_testing_RoxanaDarvari.xlsm: An Excel workbook with multiple tabs, each representing a different set of stock data for analysis.
- Screenshots: containing screenshots of each tab in the workbooks after the script has been run.

How to Use
•	Open the desired Excel workbook.
•	Go to Developer menu and open the Visual Basic.
•	Import the Multiple_year_stock_data_RoxanaDarvari.vbs file into the VBA editor.
•	Run the AnalyzeStockSheet subroutine.
•	A message box will appear upon completion of the script, indicating that the analysis is complete.
•	Review the results in each worksheet which will include:
      o	The "Ticker" column showing the stock symbol.
      o	The "Yearly Change" column showing the difference between the opening price at the beginning of the year and the closing price at the end.
      o	The "Percent Change" column showing the percentage change from the opening price at the start of the year to the closing price at the end.
      o	The "Total Stock Volume" column showing the sum of traded volume throughout the year.
      o	The cells corresponding to the greatest increase and decrease percentages, as well as the greatest total volume, will be highlighted.

Features
- Multi-Sheet Analysis: The script is capable of iterating through each sheet within a workbook and performing the analysis on each one.
- Conditional Formatting: The script applies conditional formatting to the "Yearly Change" column, coloring positive changes green and negative changes red.
- Dynamic Calculation: The script dynamically calculates the greatest values and outputs them
