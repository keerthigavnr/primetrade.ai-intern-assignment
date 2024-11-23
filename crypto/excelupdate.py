import requests
import pandas as pd
import time

# Define the exchange rate (1 USD = 82.5 INR)
exchange_rate = 82.5

# Function to fetch live cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.json()
        # Convert data to DataFrame
        df = pd.DataFrame(data)
        # Select relevant fields
        df = df[['symbol', 'name', 'current_price', 'market_cap', 'market_cap_rank', 
                 'total_volume', 'high_24h', 'low_24h', 'price_change_24h', 
                 'price_change_percentage_24h', 'market_cap_change_24h', 
                 'market_cap_change_percentage_24h']]
        return df
    else:
        print("Error fetching data:", response.status_code)
        return pd.DataFrame()

# Function to format the DataFrame with USD and INR
def format_dataframe_with_symbols(df, rate):
    # Duplicate USD columns and add INR columns next to them with symbols
    df["current_price(USD)"] = df["current_price"].apply(lambda x: f"${x:,.2f}")
    df["market_capitalization(USD)"] = df["market_cap"].apply(lambda x: f"${x:,.0f}")
    df["24_hour_trading_volume(USD)"] = df["total_volume"].apply(lambda x: f"${x:,.0f}")
    df["high_24h(USD)"] = df["high_24h"].apply(lambda x: f"${x:,.2f}")
    df["low_24h(USD)"] = df["low_24h"].apply(lambda x: f"${x:,.2f}")
    df["price_change_24h(USD)"] = df["price_change_24h"].apply(lambda x: f"${x:,.2f}")
    
    # INR Columns with formatting
    df["current_price(INR)"] = df["current_price"].apply(lambda x: f"₹{x * rate:,.2f}")
    df["market_capitalization(INR)"] = df["market_cap"].apply(lambda x: f"₹{x * rate:,.0f}")
    df["24_hour_trading_volume(INR)"] = df["total_volume"].apply(lambda x: f"₹{x * rate:,.0f}")
    df["high_24h(INR)"] = df["high_24h"].apply(lambda x: f"₹{x * rate:,.2f}")
    df["low_24h(INR)"] = df["low_24h"].apply(lambda x: f"₹{x * rate:,.2f}")
    df["price_change_24h(INR)"] = df["price_change_24h"].apply(lambda x: f"₹{x * rate:,.2f}")

    # Format percentage columns
    df["price_change_percentage_24h"] = df["price_change_percentage_24h"].apply(lambda x: f"{x:.2f}%")
    df["market_cap_change_percentage_24h"] = df["market_cap_change_percentage_24h"].apply(lambda x: f"{x:.2f}%")

    # Remove original columns with exponential notation
    df.drop(["current_price", "market_cap", "total_volume", "high_24h", "low_24h", "price_change_24h"], axis=1, inplace=True)

    return df

# Function to save the formatted data to Excel
def update_excel():
    df = fetch_crypto_data()
    if not df.empty:
        # Format the data
        formatted_df = format_dataframe_with_symbols(df, exchange_rate)
        # Save to Excel file
        formatted_df.to_excel("crypto_data.xlsx", index=False, engine="openpyxl")
        print("Excel updated at:", pd.Timestamp.now())

# Main loop for live updates
while True:
    update_excel()
    time.sleep(10)  # Wait for 5 minutes (300 seconds)
