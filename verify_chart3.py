import sys
import os
import pandas as pd
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from data_loader import load_company_data

def verify():
    company = "比亚迪"
    print(f"Loading data for {company}...")
    data = load_company_data(company)
    
    if 'financials' in data:
        df = data['financials']
        print("\nFinancials Data Found!")
        print(f"Shape: {df.shape}")
        print(f"Columns: {df.columns}")
        
        if 'market_cap' in df.columns:
            print("\nSUCCESS: 'market_cap' column exists in financials.")
            print("First 5 rows of market_cap:")
            print(df['market_cap'].head())
            print("\nLast 5 rows of market_cap:")
            print(df['market_cap'].tail())
            
            # Check for NaNs in market_cap
            nan_count = df['market_cap'].isna().sum()
            print(f"\nNaN count in market_cap: {nan_count} / {len(df)}")
        else:
            print("\nFAILURE: 'market_cap' column MISSING from financials.")
            
    else:
        print("No financials data found.")

if __name__ == "__main__":
    verify()
