import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import os
import warnings

# Suppress specific Visual Code warnings
warnings.filterwarnings("ignore", category=FutureWarning, module="yfinance")
warnings.filterwarnings("ignore", category=RuntimeWarning, module="numpy")

def compute_beta(stock_returns, market_returns):
    if len(stock_returns) < 2 or len(market_returns) < 2:
        return np.nan
    covariance_matrix = np.cov(stock_returns, market_returns)
    market_variance = covariance_matrix[1, 1]
    if market_variance == 0:
        return np.nan
    beta = covariance_matrix[0, 1] / market_variance
    return beta

def compute_sharpe_ratio(returns, risk_free_rate):
    if len(returns) < 2:
        return np.nan
    excess_returns = returns - risk_free_rate
    std_excess_returns = np.std(excess_returns)
    if std_excess_returns == 0:
        return np.nan
    sharpe_ratio = np.mean(excess_returns) / std_excess_returns
    return sharpe_ratio

def download_data(ticker, interval='1d'):
    return yf.download(ticker, period='max', interval=interval)

def get_annual_metrics(tickers, market_ticker='^GSPC'):
    risk_free_rate = 0.02 / 252  # Assuming a constant annual risk-free rate of 2%, adjusted for daily returns. Adjust this value if needed.
    metrics = []
    success_tickers = []
    failed_tickers = []

    market_data = download_data(market_ticker)
    
    for ticker in tickers:
        try:
            stock_data = download_data(ticker)
            for year in stock_data.index.year.unique():
                try:
                    stock_year_data = stock_data[stock_data.index.year == year]
                    market_year_data = market_data[market_data.index.year == year]

                    if stock_year_data.empty or market_year_data.empty:
                        continue

                    # Align the data by dates
                    stock_year_data = stock_year_data.reindex(market_year_data.index).dropna()
                    market_year_data = market_year_data.reindex(stock_year_data.index).dropna()

                    stock_returns = stock_year_data['Adj Close'].pct_change().dropna()
                    market_returns = market_year_data['Adj Close'].pct_change().dropna()

                    beta = compute_beta(stock_returns, market_returns)
                    sharpe_ratio = compute_sharpe_ratio(stock_returns, risk_free_rate)

                    metrics.append({
                        'Ticker': ticker,
                        'Year': year,
                        'Metric': 'Beta',
                        'Value': beta
                    })
                    metrics.append({
                        'Ticker': ticker,
                        'Year': year,
                        'Metric': 'Sharpe Ratio',
                        'Value': sharpe_ratio
                    })
                except Exception as e:
                    print(f"Failed to process data for {ticker} in year {year}: {e}")
                    continue
            success_tickers.append(ticker)
        except Exception as e:
            print(f"Failed to download data for {ticker}: {e}")
            failed_tickers.append(ticker)
    
    return pd.DataFrame(metrics), success_tickers, failed_tickers

# Companies tickers go here
tickers = ['', '', '', '', '', '', '', '', '', '']
# Empty tickers list: ['', '', '', '', '', '', '', '', '', '']

# Get annual metrics
annual_metrics, success_tickers, failed_tickers = get_annual_metrics(tickers)

# Pivot the data frame to the desired format
pivot_table = annual_metrics.pivot_table(index=['Ticker', 'Metric'], columns='Year', values='Value').reset_index()

# Specify the folder path
folder_path = r'C:\Users\wagne\OneDrive\IAE Aix-en-Provence\Master Thesis\Thesis'

# Specify the file path
file_path = os.path.join(folder_path, 'annual_metrics.xlsx')

# Save to Excel
pivot_table.to_excel(file_path, index=False)

# Print summary of successes and failures
print("\nSummary of Data Processing:")
print(f"Successfully processed tickers: {success_tickers}")
print(f"Failed to process tickers: {failed_tickers}")
