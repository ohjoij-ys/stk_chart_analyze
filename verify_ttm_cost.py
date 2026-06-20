
import sys
import os
import pandas as pd
import numpy as np

# Mocking the load_excel_data and calculate_ttm functions for isolation

def load_benefit_raw():
    f_path = r'e:/Work/Stock_chart/数据导出/比亚迪/002594_benefit_report.xls'
    try:
        df = pd.read_excel(f_path, header=None, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
    except TypeError:
        df = pd.read_excel(f_path, header=None, engine='xlrd')
        
    date_row_idx = 1
    dates = df.iloc[date_row_idx, 1:].values
    data_start_idx = date_row_idx + 1
    indicators = df.iloc[data_start_idx:, 0].values
    values = df.iloc[data_start_idx:, 1:].values.T
    
    new_df = pd.DataFrame(values, index=dates, columns=indicators)
    new_df.index = pd.to_datetime(new_df.index, errors='coerce')
    new_df = new_df.sort_index()
    new_df = new_df[new_df.index.notnull()]
    
    # Clean and convert
    def clean_value(x):
        if isinstance(x, str):
            x = x.replace('%', '').replace(',', '')
        return x
    new_df = new_df.map(clean_value)
    new_df = new_df.apply(pd.to_numeric, errors='coerce')
    
    return new_df

def calculate_ttm_debug(df, col_name):
    print(f"\n--- Tracing TTM for {col_name} ---")
    
    # 1. Show raw cumulative data for last 2 years
    raw_series = df[col_name].tail(8)
    print("\n[Raw Cumulative Data]")
    print(raw_series)
    
    # 2. Calculate Single Quarter (SQ)
    sq_values = []
    sq_indices = []
    
    # We need to iterate through the whole DF to handle previous year logic correctly
    # But for display, we focus on the tail
    
    dates = df.index
    values = df[col_name].values
    
    print("\n[Single Quarter Calculation]")
    for i in range(len(df)):
        curr_date = dates[i]
        curr_val = values[i]
        sq_val = None
        
        note = ""
        
        if pd.isna(curr_val):
            sq_val = None
            note = "NaN"
        elif curr_date.month == 3:
            sq_val = curr_val
            note = "Q1 (Raw)"
        else:
            if i > 0:
                prev_date = dates[i-1]
                if prev_date.year == curr_date.year:
                    prev_val = values[i-1]
                    if pd.notna(prev_val):
                        sq_val = curr_val - prev_val
                        note = f"Cumulative - PrevQ ({curr_val:.0f} - {prev_val:.0f})"
                    else:
                        note = "PrevQ is NaN"
                else:
                    note = "No PrevQ in same year"
            else:
                note = "Start of data"
        
        sq_values.append(sq_val)
        sq_indices.append(curr_date)
        
        # Print only last 8 quarters
        if i >= len(df) - 8:
            print(f"{curr_date.date()}: SQ={sq_val:,.0f} ({note})")
            
    df_sq = pd.Series(sq_values, index=sq_indices)
    
    # 3. Calculate TTM (Rolling sum of 4 SQ)
    df_ttm = df_sq.rolling(window=4, min_periods=4).sum()
    
    print("\n[TTM Calculation (Sum of last 4 SQs)]")
    print(df_ttm.tail(8))
    
    return df_ttm

def verify_cost():
    print("Loading Benefit Report...")
    df = load_benefit_raw()
    
    target_col = '二、营业总成本(元)'
    if target_col in df.columns:
        ttm_series = calculate_ttm_debug(df, target_col)
    else:
        print(f"Column '{target_col}' not found.")

if __name__ == "__main__":
    verify_cost()
