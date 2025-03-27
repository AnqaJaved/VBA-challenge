# VBA Stock Analysis Project


## Author: Anqa Javed

This repository contains a **VBA (Visual Basic for Applications)** script designed to analyze stock data in Excel. The script processes multiple sheets of stock data and performs calculations such as **Quarterly Change**, **Percentage Change**, and identifies the **Greatest % Increase**, **Greatest % Decrease**, and **Greatest Total Volume**.

### Files and Structure:

- **Resources**: Contains all the data, screenshots, and scripts used for the project.
- **alphabetical_testing.xlsx**: The file used to store the stock data and results. It contains multiple sheets with stock information, which is processed by the VBA script to generate the necessary calculations and results.
- **Multiple_year_stock_data.xlsm**: The primary stock data file used in the project. It contains multi-year stock data across several years and is used as the main data source to run the analysis and obtain results using the `VBA_Stock_Analysis.bas` script.
- **VBA_Stock_Analysis.bas**: The exported VBA script containing all the functionality for data analysis.
- **Screenshots**: Includes screenshots of the results after running the VBA code.

### Features:
- **Conditional Formatting**: Applied to the **Quarterly Change** and **Percent Change** columns, highlighting positive changes in green and negative changes in red.
- **Summary Table**: Displays **Ticker**, **Quarterly Change**, **Percent Change**, and **Total Stock Volume** for each stock.
- **Bonus Table**: Displays the stock with the **Greatest % Increase**, **Greatest % Decrease**, and **Greatest Total Volume**.

### How to Use:
1. Open the **Multiple_year_stock_data.xlsm** file that contains the stock data for multiple years.
2. Open the **alphabetical_testing.xlsm** file that contains the results of running the VBA script.
3. Open the **VBA Developer Console** in Excel.
4. Run the `StockAnalysis` macro to calculate and analyze the stock data.
5. View the results in the summary and bonus tables within the Excel workbook.


