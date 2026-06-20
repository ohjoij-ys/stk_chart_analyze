
import pandas as pd
import sys
import os

# Add src to path
sys.path.append(os.path.join(os.getcwd(), 'src'))

from data_loader import load_company_data
from config import DATA_DIR

def verify_weekly_pe_ps():
    # Load data for a sample company (assuming one exists in data dir)
    # We'll list directories in data/financial_data to find a company
    companies = [d for d in os.listdir(os.path.join(DATA_DIR, 'financial_data')) if os.path.isdir(os.path.join(DATA_DIR, 'financial_data', d))]
    if not companies:
        print("No companies found.")
        return
    
    company = companies[0]
    print(f"Verifying data for company: {company}")
    
    data = load_company_data(company, DATA_DIR)
    
    if 'mixed_weekly' not in data:
        print("Error: 'mixed_weekly' not found in data.")
        return
        
    df = data['mixed_weekly']
    print(f"Mixed Weekly Shape: {df.shape}")
    print(f"Columns: {df.columns.tolist()}")
    
    if 'pe_ttm' not in df.columns or 'ps_ttm' not in df.columns:
        print("Error: pe_ttm or ps_ttm missing from mixed_weekly")
        return
        
    # Check granularity
    print("\n--- Inspecting first 10 rows of PE/PS ---")
    print(df[['market_cap', 'profit', 'pe_ttm', 'revenue', 'ps_ttm']].head(10))
    
    # Check if dates are weekly (approx 7 days diff)
    dates = df.index.to_series()
    diffs = dates.diff().dt.days.dropna()
    print(f"\nDate diffs (unique): {diffs.unique()}")
    
    # Check if profit changes less frequently than market cap
    # We expect profit to be constant for chunks of time (quarters)
    # while market cap changes every week.
    
    print("\n--- Checking Variation ---")
    profit_changes = df['profit'].nunique()
    mc_changes = df['market_cap'].nunique()
    rows = len(df)
    
    print(f"Total Rows: {rows}")
    print(f"Unique Profit Values: {profit_changes}")
    print(f"Unique Market Cap Values: {mc_changes}")
    
    if profit_changes < rows / 4 and mc_changes > rows * 0.8:
        print("SUCCESS: Profit appears to be quarterly (stable) while Market Cap is weekly (volatile).")
    else:
        print("WARNING: Variation patterns might be unexpected.")

if __name__ == "__main__":
    verify_weekly_pe_ps()
