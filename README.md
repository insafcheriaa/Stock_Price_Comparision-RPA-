# Stock Price Trend Comparison using UiPath

## Overview
This project is an RPA solution developed using **UiPath Studio** to automate the collection, analysis, and reporting of stock price trends for two companies: **Exxon RPA Corp.** and **WEX Academy Inc.** The solution extracts stock prices from a webpage at regular intervals, generates a graphical representation of the trends, and sends the final report via email at the end of the trading day.

## Features
- Automated data extraction of stock prices every 30 minutes during trading hours (9 AM - 4 PM).
- Real-time recording of stock prices in an Excel file with date and timestamp.
- Generation of a dual-line graph to visualize price trends over time.
- Emailing the final report (Excel file) to a predefined recipient at the end of the trading session.

## Prerequisites
### System Requirements:
- **UiPath Studio**: Academic Alliance Version 2021.10 or later.
- **Microsoft Excel**: For data storage and graph generation.
- **Microsoft Outlook**: For sending emails.
- **Web Browser**: Microsoft Edge (for extracting stock prices).

### Required Packages:
Ensure the following UiPath Studio packages are installed:
- UiPath.Excel.Activities (v2.11.4)
- UiPath.System.Activities (v22.4.1)
- UiPath.UiAutomation.Activities (v21.10.5)
- UiPath.Mail.Activities (v1.12.3)

### Initial Files:
1. **Data.xlsx**: 
   - Columns: `System Date`, `System TimeStamp`, `Exxon RPA Corp`, `WEX Academy Inc`.
   - Predefined headers and an empty graph template.

2. **Config.xlsx**:
   - Contains email configuration details:
     - Receiver’s Email Address
     - Mail Subject (e.g., "Stock Price Report")
     - Mail Body Message

## Process Workflow
1. **Start**: Verify if the current time is within trading hours.
2. **Data Extraction**:
   - Open the RPA Stock Market webpage.
   - Retrieve stock prices for both companies.
   - Store the data in `Data.xlsx` with a timestamp.
3. **Graph Creation**:
   - After all intervals are processed, create a dual-line graph in the same Excel file.
4. **Report Delivery**:
   - Email the final Excel file to the recipient defined in `Config.xlsx`.
5. **End**: Stop the automation.

## Webpage Used
Stock prices are retrieved from [RPA Stock Market](https://www.rpachallenge.com/assets/rpaStockMarket/index.html).

## How to Run
1. Ensure all prerequisites are met and initial files are configured.
2. Open the project in UiPath Studio.
3. Execute the process, ensuring the system remains active during the trading hours (9 AM - 4 PM).

## Results
At the end of the trading day:
- The generated Excel file will contain the stock prices and their visual trends.
- The file will be emailed to the designated recipient.

## License
This project is developed as part of an RPA specialization and is free to use for educational and non-commercial purposes.
