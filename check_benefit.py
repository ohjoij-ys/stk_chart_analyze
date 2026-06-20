
import pandas as pd
import os
import sys

# Suppress warnings
import warnings
warnings.filterwarnings('ignore')

def load_benefit_columns():
    f_path = r'e:/Work/Stock_chart/数据导出/比亚迪/002594_benefit_report.xls'
    try:
        # Mimic load_excel_data from data_loader.py
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
        print("Loaded Benefit Report Columns:")
        print(new_df.columns.tolist())
        
        target_col = '其中：营业成本(元)'
        if target_col in new_df.columns:
            print(f"\nFound '{target_col}'. Head values:")
            print(new_df[target_col].head())
        else:
            print(f"\n'{target_col}' NOT found.")
            # Search for similar
            for col in new_df.columns:
                if '营业成本' in str(col):
                    print(f"Did you mean: {col} ?")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    load_benefit_columns()
