import pandas as pd
import yfinance as yf
from datetime import datetime
import openpyxl

# Function to fetch stock data
def get_stock_data(ticker, start_date, end_date):

    try:
        data = yf.download(ticker, start=start_date, end=end_date)
        return data['Close']
    except Exception as e:
        print(f"Could not fetch data for {ticker}. Error: {e}")
        return None

def stock_allocation(csv_file, start_date, end_date, total_investment):
    # Read the stock details from the CSV
    try:
        stocks = pd.read_csv(csv_file)
    except FileNotFoundError:
        print("Error: CSV file not found. Please check the file path.")
        return
    except Exception as e:
        print(f"Error reading the CSV file: {e}")
        return

    results = pd.DataFrame()

    for _, row in stocks.iterrows():
        ticker = row['Ticker']
        weightage = row['Weightage']

        # All the ticker symbols should end with .NS for NSE stocks
        if not ticker.endswith('.NS'):
            ticker = f"{ticker}.NS"

        # Get the stock data
        stock_data = get_stock_data(ticker, start_date, end_date)
        if stock_data is None:
            continue

        # Allocate investment based on weightage and calculate shares
        allocation = total_investment * weightage
        shares = allocation / stock_data

        # Prepare data for the current stock
        shares_df = pd.DataFrame({ticker: shares})
        shares_df.index = shares_df.index.strftime('%Y-%m-%d')
        shares_df = shares_df.transpose()
        shares_df['Ticker'] = row['Ticker']
        shares_df['Weightage'] = weightage
        results = pd.concat([results, shares_df])


    # Reset the index and reorder columns
    results.reset_index(drop=True, inplace=True)
    results = results[['Ticker', 'Weightage'] + list(results.columns.difference(['Ticker', 'Weightage']))]

     # Save to Excel
    output_file = 'stock_allocation_results.xlsx'
    try:
        results.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Results saved to {output_file}")
    except Exception as e:
        print(f"Error saving results to Excel: {e}")

if __name__ == "__main__":
    print("********** Stocks Allocationyzer **********")

    # Get user input
    csv_file = "Stocks.csv" 
    start_date = input("Enter the start date (YYYY-MM-DD): ")
    end_date = input("Enter the last date (YYYY-MM-DD): ")

    try:
        total_investment = float(input("Enter the total investment amount: "))
    except ValueError:
        print("Error: Total investment must be a number.")
        exit(1)

    # Validate the dates
    try:
        start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
        end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")
        current_date = datetime.now()

        if start_date_obj > current_date or end_date_obj > current_date:
            print("Error: Dates must be in the past.")
            exit(1)
        if end_date_obj < start_date_obj:
            print("Error: End date cannot be greater than the start date.")
            exit(1)
        if total_investment <= 0:
            print("Error: Investment amount must be greater than zero.")
            exit(1)
    except ValueError:
        print("Error: Invalid date format. Please use YYYY-MM-DD.")
        exit(1)

    # Main function for stock allocation
    stock_allocation(csv_file, start_date, end_date, total_investment)
