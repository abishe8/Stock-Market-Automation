# Zerodha Excel Trading System

# DISCLAIMER: ANY LOSSES INCURRED VIA THE EXCEL TRADING SYSTEM ARE AT THE USER'S OWN RISK 

The **Zerodha Excel Trading System** is a Python-based application that integrates with the Zerodha KiteConnect API to automate stock trading and portfolio management using an Excel workbook (`Trading_System.xlsx`). It allows users to monitor market data, place buy/sell orders, and generate daily trade summaries, all managed through a user-friendly Excel interface. The application supports both live trading and paper trading modes, making it suitable for testing strategies without risking real funds.

## Features

- **Excel-Based Interface**: Manage trades, view portfolio data, and track order history directly in Excel.
- **Real-Time Market Data**: Fetch and display Last Traded Price (LTP), Open, High, Low, and Previous Close prices for selected stocks.
- **Order Placement**: Place market, limit, and cover orders (with stop-loss) for stocks on NSE and BSE exchanges.
- **Risk Management**: Optional "open=high" filter to block risky buy orders when a stock's open price equals its high price.
- **Portfolio Dashboard**: View funds available, margin used, positions, holdings, and detailed trade metrics (e.g., profit/loss, day buy/sell quantities).
- **Order History**: Log all executed and rejected orders with timestamps, quantities, and prices.
- **Daily Trade Summary**: Generate an end-of-day trade summary in a separate Excel file (`DD-MM-YYYY-trade_summary.xlsx`) with completed orders.
- **Paper Trading Mode**: Simulate trades without connecting to the Zerodha API for testing purposes.
- **Error Handling and Logging**: Robust logging to a file (`trading_system.log`) and console for debugging and monitoring.

## Prerequisites

Before using the application, ensure you have the following:

1. **Python 3.8+** installed on your system.
2. **Zerodha KiteConnect API Credentials**:
   - Obtain an API Key and API Secret from Zerodha's developer portal.
   - Generate an Access Token by logging into the KiteConnect API (instructions provided during setup).
3. **Excel**: Microsoft Excel installed (Windows or macOS) for the `xlwings` library to interact with the workbook.
4. **Python Libraries**:
   - Install required libraries by running:

     ```bash
     pip install kiteconnect xlwings pandas pytz
     ```
5. **Operating System**: The application is designed for Windows (due to commented `winsound` usage in earlier versions) but can run on macOS or Linux with minor modifications (e.g., remove or replace `winsound` beeps if re-enabled).

## Setup Instructions

1. **Clone or Download the Code**:

   - Download the `algotrade.py` file or clone the repository to your local machine. Make sure the python code and the excel file are in the same directory

2. **Install Dependencies**:

   - Open a terminal in the project directory and run:

     ```bash
     pip install kiteconnect xlwings pandas pytz
     ```

3. **Configure the Excel Workbook**:

   - Run `algotrade.py` once to generate the `Trading_System.xlsx` file in the same directory:

     ```bash
     python algotrade.py
     ```
   - The script creates an Excel workbook with four sheets:
     - **Configuration**: Stores API credentials and settings.
     - **Dashboard**: Displays portfolio and trade metrics.
     - **Order Book**: Input orders and view real-time market data.
     - **Order History**: Logs all executed and rejected orders.

4. **Set Up API Credentials**:

   - Open `Trading_System.xlsx` and navigate to the **Configuration** sheet.
   - Fill in the following fields in the `Value` column (starting at cell `B2`):
     - **User ID**: Your Zerodha user ID.
     - **API Key**: Your KiteConnect API key.
     - **API Secret**: Your KiteConnect API secret.
     - **Access Token**: Leave blank initially; the script will prompt you to generate one.
     - **Polling Interval**: Time (in seconds) between updates (default: `20`). It is advisable not to set the polling interval under 15 seconds
     - **Paper Trading**: Set to `TRUE` for simulated trading or `FALSE` for live trading (default: `FALSE`).
   - Save the workbook.

5. **Generate Access Token**:

   - Run the script:

     ```bash
     python algotrade.py
     ```
   - If no valid Access Token is found in the Configuration sheet, the script will display a login URL (e.g., `https://kite.trade/connect/login?api_key=<your_api_key>`).
   - Visit the URL in a browser, log in with your Zerodha credentials, and copy the `request_token` from the redirect URL.
   - Paste the `request_token` into the terminal prompt.
   - The script will generate an Access Token and save it to the Configuration sheet (`B5`).

6. **Verify Setup**:

   - Ensure the script connects to the KiteConnect API (check the console for a message like `Connected to KiteConnect as <user_name>`).
   - Confirm that `Trading_System.xlsx` opens automatically and is populated with the template sheets.

## How to Use

### 1. Monitor Market Data

- Open the **Order Book** sheet in `Trading_System.xlsx`.
- In cells `A2:A8`, enter stock tickers in the format `STOCKCODE-EXCHANGE` (e.g., `GREENPOWER-NSE`, `GENNEX-BSE`, `IDEA-NSE`).
- When the market is open (Monday–Friday, 9:15 AM to 3:30 PM IST), the script updates cells `B2:F8` with:
  - **LTP**: Last Traded Price
  - **Open Price**: Day's opening price
  - **High Price**: Day's high price
  - **Low Price**: Day's low price
  - **Previous Close Price**: Previous day's closing price
- If the market is closed or the ticker is invalid, these cells will show `N/A` or `Invalid Ticker`.

### 2. Place Orders

- Navigate to the **Order Book** sheet, starting at row 10 (below headers).
- Fill in the following columns for each order:
  - **Stock Code**: The stock symbol (e.g., `IDEA`, `GREENPOWER`).
  - **Exchange**: `NSE` or `BSE`.
  - **Quantity**: Number of shares (e.g., `1`).
  - **Transaction Type**: `buy` or `sell` (case-insensitive).
  - **Order Type**: `market`, `limit`, or `cover` (case-insensitive).
  - **Product Type**: `MIS` (intraday) or `CNC` (delivery).
  - **risk buy filter (open=high)**: Set to `TRUE` to block buy orders if the stock's open price equals its high price; otherwise, leave blank or set to `FALSE`.
  - **Buy Price**: For `limit` or `cover` orders, specify the buy price; leave blank for `market` orders.
  - **Sell Price**: For limit sell orders, specify the sell price (optional).
  - **Stop Loss Price**: For `cover` orders, specify the stop-loss trigger price (optional).
- The script processes orders every 20 seconds (configurable via Polling Interval).
- The **Order Status** column (`K`) will update with:
  - `Processing...` (yellow): Order is being processed.
  - `Bought | Order ID: <id> | Status: <status>` (green): Buy order executed.
  - `Sold | Order ID: <id> | Status: <status>` (green): Sell order executed.
  - `Error: <message>` (red): Order failed (e.g., insufficient funds, invalid stock code).
- For paper trading, orders are simulated, and the status will show `Bought (Paper Trading)` or `Sold (Paper Trading)`.

### 3. View Portfolio and Trade Metrics

- Open the **Dashboard** sheet to monitor:
  - **Funds Available**: Available cash balance (cell `B1`).
  - **Margin Used**: Margin utilized (cell `B2`).
  - **Positions**: Number of open positions (cell `B3`).
  - **Holdings**: Number of stocks in your portfolio (cell `B4`).
  - **Trade Table** (starting at `A7`): Displays per-stock metrics, including:
    - Stock Code, Market, Product Type, Quantity, Avg Buy/Sell Price, LTP, Profit/Loss, Currently Holding, Day Buy/Sell Quantity, Day Buy/Sell Price.
    - Profit/Loss cells are color-coded: green for gains, red for losses.
- The dashboard updates every 20 seconds during market hours.

### 4. Track Order History

- The **Order History** sheet logs all executed (`COMPLETE`) and rejected (`REJECTED`) orders.
- Columns include:
  - **Timestamp**: Date and time of the order (IST).
  - **Stock Code**: Stock symbol.
  - **Action**: `BUY` or `SELL` (color-coded: green for BUY, red for SELL).
  - **Quantity**: Number of shares.
  - **Price**: Average execution price or limit price.
  - **Status**: `COMPLETE` or `REJECTED`.
  - **Order ID**: Unique order identifier from Zerodha.
- The sheet is cleared and updated with the latest orders every 20 seconds.

### 5. Generate Daily Trade Summary

- After market hours (post 3:30 PM IST, Monday–Friday), the script generates a daily trade summary file named `DD-MM-YYYY-trade_summary.xlsx` (e.g., `24-07-2025-trade_summary.xlsx`).
- The file is created only once per day and only if it doesn’t already exist.
- The summary includes all completed orders for the day, with columns:
  - **Timestamp**: Order execution time (IST).
  - **Stock Code**: Stock symbol.
  - **Quantity**: Number of shares.
  - **Action**: `BUY` or `SELL` (color-coded: green for BUY, red for SELL).
  - **Product Type**: `MIS` or `CNC`.
  - **Average Price**: Execution price.
- The summary is generated automatically when the market is closed and the script is running.

### 6. Stop the Application

- Press `Ctrl+C` in the terminal to stop the script.
- The script will save the Excel workbook (`Trading_System.xlsx`) before exiting.
- Check the `trading_system.log` file for detailed logs of all actions, errors, and API interactions.

## Example Usage

1. **Configure the Workbook**:

   - Open `Trading_System.xlsx` and update the Configuration sheet:

     ```
     Field            | Value
     -----------------|--------------------------------
     User ID          | <Zerodha ID>
     API Key          | <API Key>
     API Secret       | <API secret>
     Access Token     | <leave blank initially>
     Polling Interval | 20
     Paper Trading    | FALSE
     ```

2. **Run the Script**:

   ```bash
   python algotrade.py
   ```

   - Follow the prompt to generate an Access Token if needed.

3. **Monitor Stocks**:

   - In the Order Book sheet, enter tickers in `A2:A8`:

     ```
     A2: GREENPOWER-NSE
     A3: GENNEX-BSE
     A4: IDEA-NSE
     ```
   - Check cells `B2:F8` for real-time market data during market hours.

4. **Place an Order**:

   - In the Order Book sheet, row 11:

     ```
     Stock Code: IDEA
     Exchange: NSE
     Quantity: 1
     Transaction Type: buy
     Order Type: market
     Product Type: MIS
     risk buy filter (open=high): FALSE
     Buy Price: <blank>
     Sell Price: <blank>
     Stop Loss Price: <blank>
     ```
   - Wait 20 seconds for the script to process the order. Check the `Order Status` column for updates.

5. **Check Trade Summary**:

   - After 3:30 PM IST, check the project directory for `DD-MM-YYYY-trade_summary.xlsx`.
   - Open the file to view all completed orders for the day, e.g.:

     ```
     Timestamp                 | Stock Code | Quantity | Action | Product Type | Average Price
     --------------------------|------------|----------|--------|--------------|---------------
     2025-07-24 11:37:00       | GREENPOWER | 1        | SELL   | MIS          | 14.44
     2025-07-24 11:37:01       | GREENPOWER | 1        | BUY    | MIS          | 14.45
     2025-07-24 14:36:00       | IDEA       | 1        | BUY    | MIS          | 7.4
     ```

## Troubleshooting

- **API Connection Issues**:
  - Ensure your API Key, API Secret, and Access Token are correct in the Configuration sheet.
  - If the Access Token is invalid, delete it from `B5` and rerun the script to generate a new one.
- **Excel File Not Saving**:
  - Check if `Trading_System.xlsx` is open in another application; close it before running the script.
  - Verify write permissions in the project directory.
- **No Trade Summary Generated**:
  - Ensure the script runs after 3:30 PM IST when the market is closed.
  - Check `trading_system.log` for errors related to `generate_trade_summary`.
- **Invalid Stock Code**:
  - Use correct stock symbols and exchanges (e.g., `IDEA-NSE`, not `IDEA`).
  - Verify the stock is listed on NSE or BSE.
- **Paper Trading Mode**:
  - Set `Paper Trading` to `TRUE` in the Configuration sheet to test without real trades.
  - Simulated orders will appear in the Order History sheet with `Status: Paper Trading`.

## Notes

- **Market Hours**: The application processes orders and updates market data only during market hours (9:15 AM to 3:30 PM IST, Monday–Friday).
- **Trade Summary Timing**: The trade summary is generated after 3:30 PM IST, capturing all completed orders for the day.
- **Logging**: All actions and errors are logged to `trading_system.log` in the project directory.
- **API Rate Limits**: Be mindful of Zerodha's API rate limits to avoid throttling (consult KiteConnect documentation).
- **Excel Compatibility**: Ensure Microsoft Excel is installed, as `xlwings` requires it to manipulate the workbook.

## Support

For issues or feature requests, please:

- Check the `trading_system.log` file for detailed error messages.
- Refer to the Zerodha KiteConnect API documentation for API-related queries.
- Contact Zerodha support for API credential issues.
- For application-specific bugs, consider raising an issue on the project repository (if available) or consulting a Python developer.

Happy trading!
