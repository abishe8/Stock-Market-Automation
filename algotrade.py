import os
import time
import datetime
import pandas as pd
from kiteconnect import KiteConnect
import xlwings as xw
import sys
import signal
import logging
#import winsound
import pytz
import re

class ZerodhaExcelTradingSystem:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path
        self.kite = None
        self.paper_trading = False
        self.polling_interval = 20  # seconds
        self.running = True
        self.app = None
        self.wb = None
        self.trade_summary_generated = False  # Flag to track if trade summary was generated today
        
        # Configure logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('trading_system.log'),
                logging.StreamHandler()
            ]
        )
        
        # Register signal handler for graceful shutdown
        signal.signal(signal.SIGINT, self.signal_handler)
        
        # Initialize Excel and Kite Connect
        self.init_excel()
        self.load_config()
        self.init_kiteconnect()

    def signal_handler(self, sig, frame):
        """Handle Ctrl+C to save Excel before exiting"""
        logging.info("Saving Excel and shutting down...")
        self.save_excel()
        self.running = False
        sys.exit(0)

    def init_excel(self):
        """Initialize Excel connection and create/load workbook"""
        try:
            self.app = xw.App(visible=True, add_book=False)
            if os.path.exists(self.excel_file_path):
                self.wb = self.app.books.open(self.excel_file_path)
                logging.info(f"Opened Excel file: {self.excel_file_path}")
            else:
                self.wb = self.app.books.add()
                self.create_template_sheets()
                self.wb.save(self.excel_file_path)
                logging.info(f"Created new Excel file: {self.excel_file_path}")
        except Exception as e:
            logging.error(f"Error initializing Excel: {e}")
            raise

    def create_template_sheets(self):
        """Create template sheets with specified structure"""
        try:
            # Configuration Sheet
            config_sheet = self.wb.sheets.add('Configuration')
            config_data = [
                ['Field', 'Value'],
                ['User ID', ''],
                ['API Key', ''],
                ['API Secret', ''],
                ['Access Token', ''],
                ['Polling Interval', 20],
                ['Paper Trading', False]
            ]
            config_sheet.range('A1').value = config_data

            # Dashboard Sheet
            dashboard_sheet = self.wb.sheets.add('Dashboard')
            dashboard_labels = ['Funds Available', 'Margin Used', 'Positions', 'Holdings', '', '', '', '', '']
            dashboard_sheet.range('A1:A9').value = [[label] for label in dashboard_labels]
            dashboard_sheet.range('A1:A4').color = (200, 200, 200)
            dashboard_headers = ['Stock Code', 'Market', 'Product Type', 'Quantity', 'Avg Buy Price', 'Avg Sell Price', 'LTP', 'Profit/Loss', 'Currently Holding', 'Day Buy Quantity', 'Day Sell Quantity', 'Day Buy Price', 'Day Sell Price']
            dashboard_sheet.range('A7:M7').value = dashboard_headers
            dashboard_sheet.range('A7:M7').color = (200, 200, 200)

            # Order Book Sheet
            order_book_sheet = self.wb.sheets.add('Order Book')
            order_book_columns = ['Stock Code', 'Exchange', 'Quantity', 'Transaction Type', 'Order Type', 'Product Type', 'risk buy filter (open=high)', 'Buy Price', 'Sell Price', 'Stop Loss Price', 'Order Status']
            order_book_sheet.range('A10').value = order_book_columns
            order_book_sheet.range('A10:K10').color = (200, 200, 200)
            # Add Ticker Data section with new headers
            order_book_sheet.range('A1').value = 'Ticker Data'
            order_book_sheet.range('B1').value = 'LTP'
            order_book_sheet.range('C1').value = 'Open Price'
            order_book_sheet.range('D1').value = 'High Price'
            order_book_sheet.range('E1').value = 'Low Price'
            order_book_sheet.range('F1').value = 'Previous Close Price'
            order_book_sheet.range('A1:F1').color = (200, 200, 200)
            order_book_sheet.range('A2:A8').value = [['']] * 7  # Empty cells for user input
            order_book_sheet.range('B2:F8').value = [['']] * 7 * 5  # Empty cells for price data display

            # Order History Sheet
            history_sheet = self.wb.sheets.add('Order History')
            history_headers = ['Timestamp', 'Stock Code', 'Action', 'Quantity', 'Price', 'Status', 'Order ID']
            history_sheet.range('A1').value = history_headers
            history_sheet.range('A1:G1').color = (200, 200, 200)

            self.wb.save()
            logging.info("Created template sheets")
        except Exception as e:
            logging.error(f"Error creating template sheets: {e}")
            raise

    def load_config(self):
        """Load configuration from Configuration sheet"""
        try:
            config_sheet = self.wb.sheets['Configuration']
            config_data = config_sheet.range('A1').expand().value
            config_dict = {row[0]: row[1] for row in config_data[1:] if row[0]}
            
            self.api_key = str(config_dict.get('API Key', '')).strip()
            self.api_secret = str(config_dict.get('API Secret', '')).strip()
            self.access_token = str(config_dict.get('Access Token', '')).strip()
            self.user_id = str(config_dict.get('User ID', '')).strip()
            self.polling_interval = int(config_dict.get('Polling Interval', 60))
            self.paper_trading = bool(config_dict.get('Paper Trading', False))
            
            if not self.api_key or not self.api_secret:
                raise ValueError("API Key or API Secret missing in Configuration sheet")
            logging.info("Configuration loaded successfully")
        except Exception as e:
            logging.error(f"Error loading configuration: {e}")
            raise

    def init_kiteconnect(self):
        """Initialize KiteConnect API"""
        try:
            self.kite = KiteConnect(api_key=self.api_key)
            if self.access_token:
                try:
                    self.kite.set_access_token(self.access_token)
                    profile = self.kite.profile()
                    logging.info(f"Connected to KiteConnect as {profile['user_name']}")
                    return
                except Exception as e:
                    logging.warning(f"Invalid access token: {e}. Generating new token...")
            self.generate_access_token()
            self.kite.set_access_token(self.access_token)
            profile = self.kite.profile()
            logging.info(f"Connected to KiteConnect as {profile['user_name']}")
        except Exception as e:
            logging.error(f"Error initializing KiteConnect: {e}")
            raise

    def generate_access_token(self):
        """Generate new access token"""
        try:
            login_url = f"https://kite.trade/connect/login?api_key={self.api_key}"
            print(f"Please visit: {login_url}")
            print("After login, copy the 'request_token' from the redirect URL and paste it below.")
            request_token = input("Enter request token: ").strip()
            if not re.match(r'^[a-zA-Z0-9]{20,40}$', request_token):
                raise ValueError("Invalid request token format")
            data = self.kite.generate_session(request_token, api_secret=self.api_secret)
            self.access_token = data["access_token"]
            self.wb.sheets['Configuration'].range('B5').value = self.access_token
            self.wb.save()
            logging.info("Generated and saved new access token")
        except Exception as e:
            logging.error(f"Error generating access token: {e}")
            raise

    def is_market_open(self):
        """Check if market is open (9:15 AM to 3:30 PM IST, Monday-Friday)"""
        try:
            now = datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
            weekday = now.weekday()
            current_time = now.time()
            market_open = datetime.time(9, 15)
            market_close = datetime.time(15, 30)
            is_open = weekday < 5 and market_open <= current_time <= market_close
            logging.debug(f"Market open: {is_open}")
            return is_open
        except Exception as e:
            logging.error(f"Error checking market hours: {e}")
            return False

    def fetch_portfolio_data(self):
        """Fetch holdings, margins, positions, and orders"""
        try:
            holdings = self.kite.holdings() if not self.paper_trading else []
            margins = self.kite.margins() if not self.paper_trading else {'equity': {'available': {'cash': 0, 'live_balance': 0}, 'utilised': {'debits': 0}}}
            positions = self.kite.positions()['net'] if not self.paper_trading else []
            orders = self.kite.orders() if not self.paper_trading else []
            return holdings, margins, positions, orders
        except Exception as e:
            logging.error(f"Error fetching portfolio data: {e}")
            return [], {'equity': {'available': {'cash': 0, 'live_balance': 0}, 'utilised': {'debits': 0}}}, [], []

    def update_dashboard(self):
        """Update Dashboard sheet with consolidated trade data including day-specific metrics"""
        try:
            dashboard_sheet = self.wb.sheets['Dashboard']
            holdings, margins, positions, orders = self.fetch_portfolio_data()
            
            # Update summary data in B1:B4
            summary_data = [
                [margins.get('equity', {}).get('available', {}).get('live_balance', 0)],
                [margins.get('equity', {}).get('utilised', {}).get('debits', 0)],
                [len(positions)],
                [len(holdings)]
            ]
            dashboard_sheet.range('B1:B4').value = summary_data

            # Fetch LTP for relevant symbols
            symbols = []
            for h in holdings:
                symbols.append(f"{h['exchange']}:{h['tradingsymbol']}")
            for p in positions:
                symbols.append(f"{p['exchange']}:{p['tradingsymbol']}")
            ltp_data = {}
            if symbols and self.is_market_open():
                try:
                    ltp_data = self.kite.ltp(list(set(symbols)))
                except Exception as e:
                    logging.warning(f"Error fetching LTP: {e}")

            # Aggregate trade data with product type in key
            stock_data = {}
            for h in holdings:
                key = (h['tradingsymbol'], h['exchange'], h.get('product', 'CNC'))
                stock_data[key] = {
                    'quantity': h['quantity'],
                    'buy_price_total': h['average_price'] * h['quantity'],
                    'buy_count': h['quantity'],
                    'sell_price_total': 0,
                    'sell_count': 0,
                    'product': h.get('product', 'CNC'),
                    'exchange': h['exchange'],
                    'last_price': h['average_price'],  # Fallback for after-hours
                    'day_buy_quantity': 0,
                    'day_sell_quantity': 0,
                    'day_buy_price': 0,
                    'day_sell_price': 0
                }
            
            for p in positions:
                key = (p['tradingsymbol'], p['exchange'], p.get('product', 'MIS'))
                if key not in stock_data:
                    stock_data[key] = {
                        'quantity': 0,
                        'buy_price_total': 0,
                        'buy_count': 0,
                        'sell_price_total': 0,
                        'sell_count': 0,
                        'product': p.get('product', 'MIS'),
                        'exchange': p['exchange'],
                        'last_price': p['buy_price'] or p['sell_price'],
                        'day_buy_quantity': p.get('day_buy_quantity', 0),
                        'day_sell_quantity': p.get('day_sell_quantity', 0),
                        'day_buy_price': p.get('day_buy_price', 0),
                        'day_sell_price': p.get('day_sell_price', 0)
                    }
                stock_data[key]['quantity'] += p['quantity']
                if p['quantity'] > 0:
                    stock_data[key]['buy_price_total'] += p['buy_price'] * p['quantity']
                    stock_data[key]['buy_count'] += p['quantity']
                else:
                    stock_data[key]['sell_price_total'] += p['sell_price'] * abs(p['quantity'])
                    stock_data[key]['sell_count'] += abs(p['quantity'])
                stock_data[key]['day_buy_quantity'] = p.get('day_buy_quantity', 0)
                stock_data[key]['day_sell_quantity'] = p.get('day_sell_quantity', 0)
                if p.get('day_buy_quantity', 0) > 0:
                    stock_data[key]['day_buy_price'] = p.get('day_buy_price', 0)
                if p.get('day_sell_quantity', 0) > 0:
                    stock_data[key]['day_sell_price'] = p.get('day_sell_price', 0)

            # Update dashboard with net and day-specific quantities
            dashboard_sheet.range('A8:M1000').clear_contents()
            row = 8
            for (stock_code, exchange, product), data in stock_data.items():
                symbol = f"{exchange}:{stock_code}"
                ltp = ltp_data[symbol]['last_price'] if symbol in ltp_data else data['last_price']
                avg_buy_price = data['buy_price_total'] / data['buy_count'] if data['buy_count'] > 0 else ''
                avg_sell_price = data['sell_price_total'] / data['sell_count'] if data['sell_count'] > 0 else ''
                net_quantity = data['quantity']
                profit_loss = (ltp - avg_buy_price) * net_quantity if net_quantity > 0 and avg_buy_price else 0
                currently_holding = net_quantity > 0
                
                dashboard_sheet.range(f'A{row}:M{row}').value = [
                    stock_code, exchange, product, abs(net_quantity) if net_quantity > 0 else 0,
                    avg_buy_price, avg_sell_price, ltp, profit_loss, currently_holding,
                    data['day_buy_quantity'], data['day_sell_quantity'], data['day_buy_price'], data['day_sell_price']
                ]
                if isinstance(profit_loss, (int, float)) and profit_loss != 0:
                    dashboard_sheet.range(f'H{row}').color = (0, 255, 0) if profit_loss > 0 else (255, 0, 0)
                row += 1
            
            dashboard_sheet.autofit()
            self.wb.save()
            logging.info("Dashboard updated")
        except Exception as e:
            logging.error(f"Error updating dashboard: {str(e)}")

    def update_order_history(self):
        """Update Order History sheet with executed orders"""
        try:
            history_sheet = self.wb.sheets['Order History']
            orders = self.kite.orders() if not self.paper_trading else []
            
            # Clear existing data (keep headers)
            history_sheet.range('A2:G' + str(history_sheet.cells.last_cell.row)).clear_contents()
            
            last_row = 1  # Start after the header row

            for order in orders:
                if order['status'] in ['COMPLETE', 'REJECTED']:
                    order_id = order['order_id']
                    status = order['status']
                    quantity = order['quantity']
                    price = order['average_price'] or order['price']
                    order_timestamp = order['order_timestamp']
                    if isinstance(order_timestamp, datetime.datetime):
                        timestamp = order_timestamp.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        try:
                            timestamp = datetime.datetime.fromtimestamp(order_timestamp).strftime('%Y-%m-%d %H:%M:%S')
                        except Exception as e:
                            logging.warning(f"Invalid timestamp for order {order_id}: {e}")
                            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    
                    stock_code = order['tradingsymbol']
                    action = order['transaction_type']
                    
                    last_row += 1
                    history_sheet.range(f'A{last_row}').value = [
                        timestamp, stock_code, action, quantity, price, status, order_id
                    ]
                    history_sheet.range(f'C{last_row}').color = (0, 255, 0) if action == 'BUY' else (255, 0, 0)
            
            history_sheet.autofit()
            self.wb.save()
            logging.info("Order history updated")
        except Exception as e:
            logging.error(f"Error updating order history: {str(e)}")

    def generate_trade_summary(self):
        """Generate end-of-day trade summary in a separate Excel file"""
        try:
            now = datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
            current_date = now.date()
            file_name = current_date.strftime("%d-%m-%Y-trade_summary.xlsx")
            file_path = os.path.join(os.path.dirname(self.excel_file_path), file_name)
            
            # Check if file already exists
            if os.path.exists(file_path):
                logging.info(f"Trade summary file {file_name} already exists, skipping generation")
                return
            
            # Fetch all orders from Zerodha API
            orders = self.kite.orders() if not self.paper_trading else []
            
            # Filter executed orders for today
            today_trades = []
            for order in orders:
                if order['status'] != 'COMPLETE':
                    continue
                try:
                    # Parse timestamp and check if it's for today
                    order_timestamp = order['order_timestamp']
                    if isinstance(order_timestamp, datetime.datetime):
                        order_date = order_timestamp.date()
                    else:
                        order_timestamp = datetime.datetime.fromtimestamp(order_timestamp)
                        order_date = order_timestamp.date()
                    
                    if order_date != current_date:
                        continue
                    
                    # Get product type from order or positions
                    product_type = order.get('product', 'MIS')  # Default to MIS for intraday
                    positions = self.kite.positions()['net'] if not self.paper_trading else []
                    for pos in positions:
                        if pos['tradingsymbol'] == order['tradingsymbol'] and pos['quantity'] != 0:
                            product_type = pos.get('product', 'MIS')
                            break
                    
                    # Format timestamp in IST
                    timestamp = order_timestamp.astimezone(pytz.timezone('Asia/Kolkata')).strftime('%Y-%m-%d %H:%M:%S')
                    
                    today_trades.append([
                        timestamp,                     # Timestamp
                        order['tradingsymbol'],       # Stock Code
                        order['quantity'],             # Quantity
                        order['transaction_type'],     # Action
                        product_type,                  # Product Type
                        order['average_price'] or 0   # Average Price
                    ])
                except (ValueError, TypeError) as e:
                    logging.warning(f"Skipping invalid order entry {order.get('order_id', 'N/A')}: {e}")
                    continue
            
            # Create new Excel file
            trade_summary_app = xw.App(visible=False, add_book=False)
            trade_summary_wb = trade_summary_app.books.add()
            trade_summary_sheet = trade_summary_wb.sheets[0]
            
            # Write headers
            headers = ['Timestamp', 'Stock Code', 'Quantity', 'Action', 'Product Type', 'Average Price']
            trade_summary_sheet.range('A1').value = headers
            trade_summary_sheet.range('A1:F1').color = (200, 200, 200)
            
            # Write trade data
            if today_trades:
                trade_summary_sheet.range('A2').value = today_trades
                for row, trade in enumerate(today_trades, start=2):
                    trade_summary_sheet.range(f'D{row}').color = (0, 255, 0) if trade[3] == 'BUY' else (255, 0, 0)
            
            trade_summary_sheet.autofit()
            
            # Save and close
            for _ in range(3):
                try:
                    trade_summary_wb.save(file_path)
                    break
                except Exception as e:
                    logging.warning(f"Error saving trade summary: {e}")
                    time.sleep(1)
            else:
                logging.error(f"Failed to save trade summary after retries: {file_path}")
            
            trade_summary_wb.close()
            trade_summary_app.quit()
            logging.info(f"Generated trade summary: {file_name}")
            self.trade_summary_generated = True
        except Exception as e:
            logging.error(f"Error generating trade summary: {str(e)}")

    def process_order_book(self):
        """Process orders from Order Book sheet and update ticker price data"""
        if not self.is_market_open() and not self.paper_trading:
            logging.info("Market closed, skipping order processing and price data update")
            return
        
        try:
            order_sheet = self.wb.sheets['Order Book']
            
            # Update price data (LTP, Open, High, Low, Previous Close) in A2:F8
            ticker_data = order_sheet.range('A2:A8').value
            symbols = []
            for ticker in ticker_data:
                if ticker and isinstance(ticker, str) and '-' in ticker:
                    try:
                        stock_code, exchange = ticker.split('-')
                        if exchange.upper() in ['NSE', 'BSE']:
                            symbols.append(f"{exchange.upper()}:{stock_code}")
                    except ValueError:
                        continue
            price_data = {}
            if symbols and self.is_market_open() and not self.paper_trading:
                try:
                    price_data = self.kite.quote(list(set(symbols)))
                except Exception as e:
                    logging.warning(f"Error fetching price data for tickers: {e}")
            
            # Update price data in B2:F8
            for i, ticker in enumerate(ticker_data, start=2):
                if ticker and isinstance(ticker, str) and '-' in ticker:
                    try:
                        stock_code, exchange = ticker.split('-')
                        symbol = f"{exchange.upper()}:{stock_code}"
                        if symbol in price_data:
                            quote = price_data[symbol]
                            order_sheet.range(f'B{i}').value = quote.get('last_price', 'N/A')
                            order_sheet.range(f'C{i}').value = quote.get('ohlc', {}).get('open', 'N/A')
                            order_sheet.range(f'D{i}').value = quote.get('ohlc', {}).get('high', 'N/A')
                            order_sheet.range(f'E{i}').value = quote.get('ohlc', {}).get('low', 'N/A')
                            order_sheet.range(f'F{i}').value = quote.get('ohlc', {}).get('close', 'N/A')
                        else:
                            order_sheet.range(f'B{i}:F{i}').value = ['N/A'] * 5
                    except ValueError:
                        order_sheet.range(f'B{i}:F{i}').value = ['Invalid Ticker'] * 5
                else:
                    order_sheet.range(f'B{i}:F{i}').value = [''] * 5
            
            # Process orders starting from row 10
            order_data = order_sheet.range('A10').expand().value
            headers = order_data[0]
            orders = order_data[1:] if len(order_data) > 1 else []
            
            for i, order in enumerate(orders, start=11):  # Start at row 11 (after header at A10)
                order_dict = dict(zip(headers, order))
                if not order_dict.get('Stock Code') or not order_dict.get('Exchange'):
                    continue
                
                status = str(order_dict.get('Order Status', '')).strip()
                if 'Bought' in status or 'Sold' in status:
                    continue
                
                transaction_type = str(order_dict.get('Transaction Type', '')).lower()
                if transaction_type == 'buy':
                    self.process_buy_order(order_dict, i)
                elif transaction_type == 'sell':
                    self.process_sell_order(order_dict, i)
            
            order_sheet.autofit()
            self.wb.save()
            logging.info("Order book processed and ticker price data updated")
        except Exception as e:
            logging.error(f"Error processing order book: {e}")

    def process_buy_order(self, order, row_num):
        """Process a buy order"""
        try:
            stock_code = order['Stock Code']
            exchange = order['Exchange'].upper()
            if exchange not in ['NSE', 'BSE']:
                raise ValueError("Invalid exchange")
            symbol = f"{exchange}:{stock_code}"
            order_sheet = self.wb.sheets['Order Book']
            status_col = order_sheet.range('A10').end('right').column
            
            order_sheet.range(f'K{row_num}').value = "Processing..."
            order_sheet.range(f'K{row_num}').color = (255, 255, 0)
            self.wb.save()
            logging.info(f"Processing BUY order for {stock_code} ({exchange})...")
            #winsound.Beep(1000, 200)

            ltp_data = {}
            if self.is_market_open():
                try:
                    ltp_data = self.kite.quote(symbol)
                    if symbol not in ltp_data:
                        raise ValueError("Invalid stock code")
                except Exception as e:
                    status = f"Error: {str(e)[:100]}"
                    order_sheet.range(f'K{row_num}').value = status
                    order_sheet.range(f'K{row_num}').color = (255, 0, 0)
                    self.wb.save()
                    logging.error(status)
                    return

            risk_filter = str(order.get('risk buy filter (open=high)', '')).strip().lower() == 'true'
            if risk_filter and ltp_data and ltp_data[symbol].get('ohlc', {}).get('open') == ltp_data[symbol].get('ohlc', {}).get('high'):
                status = "Blocked: Open price equals High price"
                order_sheet.range(f'K{row_num}').value = status
                order_sheet.range(f'K{row_num}').color = (255, 0, 0)
                self.wb.save()
                logging.info(status)
                #winsound.Beep(500, 500)
                return

            if not self.paper_trading:
                try:
                    order_type = str(order.get('Order Type', '')).strip().lower()
                    buy_price = order.get('Buy Price', None)
                    
                    # Validate order type and price
                    if order_type == 'market':
                        buy_price = 0  # Market orders don't require a price
                    elif order_type in ['limit', 'cover'] and (buy_price is None or buy_price <= 0):
                        raise ValueError("Buy Price must be provided and greater than 0 for LIMIT or cover orders")
                    
                    # Check funds for non-market orders
                    if order_type != 'market':
                        margins = self.kite.margins()
                        available_cash = margins.get('equity', {}).get('available', {}).get('live_balance', 0)
                        required_margin = order['Quantity'] * buy_price
                        if available_cash < required_margin:
                            status = "Insufficient funds"
                            order_sheet.range(f'K{row_num}').value = status
                            order_sheet.range(f'K{row_num}').color = (255, 0, 0)
                            self.wb.save()
                            logging.warning(status)
                            return
                    
                    order_id = self.kite.place_order(
                        variety="regular" if order_type != 'cover' else "co",
                        tradingsymbol=stock_code,
                        exchange=exchange,
                        transaction_type='BUY',
                        quantity=int(order['Quantity']),
                        order_type='MARKET' if order_type == 'market' else 'LIMIT',
                        product=order['Product Type'],
                        price=buy_price if order_type != 'market' else 0,
                        validity='DAY'
                    )
                    
                    order_details = self.kite.order_history(order_id)[0]
                    status = f"Bought | Order ID: {order_id} | Status: {order_details['status']}"
                    order_sheet.range(f'K{row_num}').value = status
                    order_sheet.range(f'K{row_num}').color = (0, 255, 0)
                    self.wb.save()
                    logging.info(status)
                    #winsound.Beep(1500, 300)

                    # Check for stop loss if Order Type is 'cover'
                    stop_loss_price = order.get('Stop Loss Price', None)
                    if order_type == 'cover' and stop_loss_price is not None and stop_loss_price > 0:
                        sl_order_id = self.kite.place_order(
                            variety="co",
                            tradingsymbol=stock_code,
                            exchange=exchange,
                            transaction_type='SELL',
                            quantity=int(order['Quantity']),
                            order_type='SL',
                            product=order['Product Type'],
                            price=0,
                            trigger_price=stop_loss_price,
                            validity='DAY'
                        )
                        sl_order_details = self.kite.order_history(sl_order_id)[0]
                        logging.info(f"Placed Stop Loss order for {stock_code} at {stop_loss_price} | Order ID: {sl_order_id} | Status: {sl_order_details['status']}")

                    # Check for sell price and place sell order if specified
                    sell_price = order.get('Sell Price', None)
                    if sell_price and sell_price > 0:
                        sell_order_id = self.kite.place_order(
                            variety="regular",
                            tradingsymbol=stock_code,
                            exchange=exchange,
                            transaction_type='SELL',
                            quantity=int(order['Quantity']),
                            order_type='LIMIT',
                            product=order['Product Type'],
                            price=sell_price,
                            validity='DAY'
                        )
                        sell_order_details = self.kite.order_history(sell_order_id)[0]
                        logging.info(f"Placed SELL order for {stock_code} at {sell_price} | Order ID: {sell_order_id} | Status: {sell_order_details['status']}")
                    
                    ltp = ltp_data[symbol]['last_price'] if symbol in ltp_data else (buy_price if buy_price else 0)
                    profit_loss = (ltp - buy_price if buy_price else 0) * order['Quantity']
                    
                    history_sheet = self.wb.sheets['Order History']
                    last_row = history_sheet.range('A' + str(history_sheet.cells.last_cell.row)).end('up').row
                    if history_sheet.range(f'A{last_row}').value is None:
                        last_row -= 1
                    history_sheet.range(f'A{last_row + 1}').value = [
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        stock_code, 'BUY', order['Quantity'], order_details['average_price'] or buy_price or ltp,
                        order_details['status'], order_id
                    ]
                    history_sheet.range(f'C{last_row + 1}').color = (0, 255, 0)
                    history_sheet.autofit()
                    self.wb.save()
                    
                    self.update_dashboard()
                except Exception as e:
                    status = f"Error: {str(e)[:100]}"
                    order_sheet.range(f'K{row_num}').value = status
                    order_sheet.range(f'K{row_num}').color = (255, 0, 0)
                    self.wb.save()
                    logging.error(status)
                    #winsound.Beep(500, 500)
            else:
                status = "Bought (Paper Trading)"
                order_sheet.range(f'K{row_num}').value = status
                order_sheet.range(f'K{row_num}').color = (0, 255, 0)
                self.wb.save()
                logging.info(status)
                #winsound.Beep(1500, 300)
                
                history_sheet = self.wb.sheets['Order History']
                last_row = history_sheet.range('A' + str(history_sheet.cells.last_cell.row)).end('up').row
                if history_sheet.range(f'A{last_row}').value is None:
                    last_row -= 1
                history_sheet.range(f'A{last_row + 1}').value = [
                    datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    stock_code, 'BUY', order['Quantity'], buy_price or ltp_data[symbol]['last_price'] if symbol in ltp_data else 0,
                    'Paper Trading', 'N/A'
                ]
                history_sheet.range(f'C{last_row + 1}').color = (0, 255, 0)
                history_sheet.autofit()
                self.wb.save()
                
                self.update_dashboard()
        except Exception as e:
            order_sheet = self.wb.sheets['Order Book']
            order_sheet.range(f'K{row_num}').value = f"Error: {str(e)[:100]}"
            order_sheet.range(f'K{row_num}').color = (255, 0, 0)
            self.wb.save()
            logging.error(f"Error processing buy order: {e}")
            #winsound.Beep(500, 500)

    def process_sell_order(self, order, row_num):
        """Process a sell order"""
        try:
            stock_code = order['Stock Code']
            exchange = order['Exchange'].upper()
            if exchange not in ['NSE', 'BSE']:
                raise ValueError("Invalid exchange")
            symbol = f"{exchange}:{stock_code}"
            order_sheet = self.wb.sheets['Order Book']
            status_col = order_sheet.range('A10').end('right').column
            
            order_sheet.range(f'K{row_num}').value = "Processing..."
            order_sheet.range(f'K{row_num}').color = (255, 255, 0)
            self.wb.save()
            logging.info(f"Processing SELL order for {stock_code} ({exchange})...")
            #winsound.Beep(1000, 200)

            ltp_data = {}
            if self.is_market_open():
                try:
                    ltp_data = self.kite.quote(symbol)
                    if symbol not in ltp_data:
                        raise ValueError("Invalid stock code")
                except Exception as e:
                    status = f"Error: {str(e)[:100]}"
                    order_sheet.range(f'K{row_num}').value = status
                    order_sheet.range(f'K{row_num}').color = (255, 0, 0)
                    self.wb.save()
                    logging.error(status)
                    return

            if not self.paper_trading:
                try:
                    order_type = str(order.get('Order Type', '')).strip().lower()
                    sell_price = order.get('Sell Price', None)
                    
                    # Validate order type and price
                    if order_type == 'market':
                        sell_price = 0  # Market orders don't require a price
                    elif order_type in ['limit', 'cover'] and (sell_price is None or sell_price <= 0):
                        raise ValueError("Sell Price must be provided and greater than 0 for LIMIT or cover orders")
                    
                    order_id = self.kite.place_order(
                        variety="regular",
                        tradingsymbol=stock_code,
                        exchange=exchange,
                        transaction_type='SELL',
                        quantity=int(order['Quantity']),
                        order_type='MARKET' if order_type == 'market' else 'LIMIT',
                        product=order['Product Type'],
                        price=sell_price if order_type != 'market' else 0,
                        validity='DAY'
                    )
                    
                    order_details = self.kite.order_history(order_id)[0]
                    status = f"Sold | Order ID: {order_id} | Status: {order_details['status']}"
                    order_sheet.range(f'K{row_num}').value = status
                    order_sheet.range(f'K{row_num}').color = (0, 255, 0)
                    self.wb.save()
                    logging.info(status)
                    #winsound.Beep(1500, 300)
                    
                    ltp = ltp_data[symbol]['last_price'] if symbol in ltp_data else (sell_price if sell_price else 0)
                    profit_loss = (sell_price if sell_price else ltp) * order['Quantity']
                    
                    history_sheet = self.wb.sheets['Order History']
                    last_row = history_sheet.range('A' + str(history_sheet.cells.last_cell.row)).end('up').row
                    if history_sheet.range(f'A{last_row}').value is None:
                        last_row -= 1
                    history_sheet.range(f'A{last_row + 1}').value = [
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        stock_code, 'SELL', order['Quantity'], order_details['average_price'] or sell_price or ltp,
                        order_details['status'], order_id
                    ]
                    history_sheet.range(f'C{last_row + 1}').color = (255, 0, 0)
                    history_sheet.autofit()
                    self.wb.save()
                    
                    self.update_dashboard()
                except Exception as e:
                    status = f"Error: {str(e)[:100]}"
                    order_sheet.range(f'K{row_num}').value = status
                    order_sheet.range(f'K{row_num}').color = (255, 0, 0)
                    self.wb.save()
                    logging.error(status)
                    #winsound.Beep(500, 500)
            else:
                status = "Sold (Paper Trading)"
                order_sheet.range(f'K{row_num}').value = status
                order_sheet.range(f'K{row_num}').color = (0, 255, 0)
                self.wb.save()
                logging.info(status)
                #winsound.Beep(1500, 300)
                
                history_sheet = self.wb.sheets['Order History']
                last_row = history_sheet.range('A' + str(history_sheet.cells.last_cell.row)).end('up').row
                if history_sheet.range(f'A{last_row}').value is None:
                    last_row -= 1
                history_sheet.range(f'A{last_row + 1}').value = [
                    datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    stock_code, 'SELL', order['Quantity'], sell_price or ltp_data[symbol]['last_price'] if symbol in ltp_data else 0,
                    'Paper Trading', 'N/A'
                ]
                history_sheet.range(f'C{last_row + 1}').color = (255, 0, 0)
                history_sheet.autofit()
                self.wb.save()
                
                self.update_dashboard()
        except Exception as e:
            order_sheet = self.wb.sheets['Order Book']
            order_sheet.range(f'K{row_num}').value = f"Error: {str(e)[:100]}"
            order_sheet.range(f'K{row_num}').color = (255, 0, 0)
            self.wb.save()
            logging.error(f"Error processing sell order: {e}")
            #winsound.Beep(500, 500)

    def save_excel(self):
        """Save Excel file with retry"""
        try:
            for _ in range(3):
                try:
                    self.wb.save()
                    break
                except Exception as e:
                    logging.warning(f"Error saving Excel: {e}")
                    time.sleep(1)
            else:
                logging.error("Failed to save Excel after retries")
        except Exception as e:
            logging.error(f"Error saving Excel: {e}")

    def run(self):
        """Main execution loop"""
        logging.info("Starting trading system...")
        last_date = None
        while self.running:
            try:
                now = datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
                current_date = now.date()
                
                # Reset trade_summary_generated flag at midnight (new day)
                if last_date != current_date:
                    self.trade_summary_generated = False
                    last_date = current_date
                
                # Check if market is closed and trade summary hasn't been generated
                if not self.is_market_open() and now.time() > datetime.time(15, 30) and not self.trade_summary_generated:
                    self.generate_trade_summary()
                
                self.update_dashboard()
                self.update_order_history()
                self.process_order_book()
                self.save_excel()
                logging.info(f"Waiting for {self.polling_interval} seconds...")
                time.sleep(self.polling_interval)
            except Exception as e:
                logging.error(f"Error in main loop: {str(e)}")
                time.sleep(self.polling_interval)
        logging.info("Trading system stopped")

if __name__ == "__main__":
    excel_file = "Trading_System.xlsx"
    trading_system = ZerodhaExcelTradingSystem(excel_file)
    trading_system.run()
