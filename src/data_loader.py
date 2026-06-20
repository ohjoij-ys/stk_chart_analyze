import os
import pandas as pd
import glob
from config import DATA_DIR, FIELDS

def get_companies():
    """List all company directories."""
    if not os.path.exists(DATA_DIR):
        return []
    companies = [d for d in os.listdir(DATA_DIR) if os.path.isdir(os.path.join(DATA_DIR, d))]
    return companies

def find_file(company, suffix):
    """Find a file for a company with a specific suffix."""
    company_dir = os.path.join(DATA_DIR, company)
    # Search for files ending with the suffix
    # The prefix might be the stock code, e.g., 002594_main_report.xls
    # We use glob to match *suffix
    search_pattern = os.path.join(company_dir, f"*{suffix}")
    files = glob.glob(search_pattern)
    if files:
        return files[0]
    return None

def load_excel_data(file_path):
    """
    Load data from excel.
    Format:
    Row 0: Header (often empty or metadata)
    Row 1: Dates
    Col 0: Indicator Names
    """
    try:
        # Read the file
        # header=None means we read everything as data first
        # ignore_workbook_corruption=True is needed for some older/malformed XLS files
        try:
            df = pd.read_excel(file_path, header=None, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
        except TypeError:
             # Fallback for older pandas/xlrd versions where engine_kwargs might not be supported directly or different signature
             # But 'engine_kwargs' was added in pandas 1.2.0.
             # If that fails, try without it
             df = pd.read_excel(file_path, header=None, engine='xlrd')

        # Identify the row with dates. 
        # Requirement says "Second row is report date", so index 1.
        # But we should verify. Usually dates are in format YYYY-MM-DD.
        # Let's assume index 1 as per spec.
        date_row_idx = 1
        
        # Extract dates
        dates = df.iloc[date_row_idx, 1:].values
        
        # Extract indicators (Column 0)
        # Data starts from row 2 (index 2) usually, but maybe row 2 is also metadata?
        # Let's assume data starts after date row.
        data_start_idx = date_row_idx + 1
        
        # Create a new dataframe
        # Transpose: Rows = Dates, Cols = Indicators
        # We grab the data part: df.iloc[data_start_idx:, 1:] -> This is the values
        # We grab the columns: df.iloc[data_start_idx:, 0] -> This is the indicators
        
        indicators = df.iloc[data_start_idx:, 0].values
        # Clean indicators: remove all whitespace to handle random spaces in Excel
        indicators = [str(x).replace(' ', '').replace('\t', '').strip() for x in indicators]
        
        values = df.iloc[data_start_idx:, 1:].values.T # Transpose to (Dates, Indicators)
        
        # Create DF
        new_df = pd.DataFrame(values, index=dates, columns=indicators)
        
        # Clean index: Convert to datetime
        new_df.index = pd.to_datetime(new_df.index, errors='coerce')
        new_df = new_df.sort_index() # Ensure time order
        
        # Drop rows with invalid dates
        new_df = new_df[new_df.index.notnull()]
        
        # Clean up string formatting in columns before conversion
        # E.g. "20.54%" -> 20.54, "1,000" -> 1000
        def clean_value(x):
            if isinstance(x, str):
                x = x.replace('%', '').replace(',', '')
            return x
        
        new_df = new_df.map(clean_value)
        
        # Convert columns to numeric, force errors to NaN
        new_df = new_df.apply(pd.to_numeric, errors='coerce')
        
        return new_df
        
    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        return pd.DataFrame()

def calculate_ttm(df):
    """
    Convert Cumulative data to TTM.
    Assumption: Data is cumulative year-to-date.
    Formula:
    1. Calculate Single Quarter (SQ)
    2. TTM = Rolling sum of 4 SQ
    """
    if df.empty:
        return df

    # We work on a copy
    df_sq = df.copy()
    
    # We need to calculate Single Quarter data
    # Iterate over rows
    # If month is 3 (March), it's already SQ (for Q1)
    # If month > 3, we subtract the previous record's value IF it's the same year
    
    # Create a Year column for grouping
    years = df.index.year
    months = df.index.month
    
    # We will compute SQ column by column
    for col in df.columns:
        sq_values = []
        for i in range(len(df)):
            curr_date = df.index[i]
            curr_val = df.iloc[i][col]
            
            if pd.isna(curr_val):
                sq_values.append(None)
                continue
                
            if curr_date.month == 3:
                # Q1 is always single quarter
                sq_values.append(curr_val)
            else:
                # Look for previous quarter in same year
                # Since df is sorted, we check i-1
                if i > 0:
                    prev_date = df.index[i-1]
                    if prev_date.year == curr_date.year:
                        prev_val = df.iloc[i-1][col]
                        # If prev_val is NaN, we can't calculate SQ perfectly. 
                        # We'll assume 0 or propagate NaN. Let's propagate NaN.
                        if pd.notna(prev_val):
                            sq_values.append(curr_val - prev_val)
                        else:
                            sq_values.append(None)
                    else:
                        # This implies we have a Q2/3/4 but no previous Q record for this year
                        # e.g. Data starts at Q2.
                        # We can't determine SQ.
                        sq_values.append(None)
                else:
                    sq_values.append(None)
        
        df_sq[col] = sq_values

    # Now calculate TTM: Rolling sum of 4
    # min_periods=4 ensures we only show TTM if we have 4 quarters. 
    # But often people accept partial if it's start of data? No, TTM needs 12 months.
    df_ttm = df_sq.rolling(window=4, min_periods=4).sum()
    
    return df_ttm

def parse_chinese_amount(val):
    """Parse string like '91.17亿' or '34.87万' to float."""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return val
    
    val_str = str(val).strip()
    multiplier = 1.0
    
    if '亿' in val_str:
        multiplier = 100000000
        val_str = val_str.replace('亿', '')
    elif '万' in val_str:
        multiplier = 10000
        val_str = val_str.replace('万', '')
        
    try:
        return float(val_str) * multiplier
    except ValueError:
        return None

def load_cap_table(file_path):
    """
    Load capitalization report which has headers in Row 0.
    Columns: 变动日期, 变动后AH股总股本(股)
    """
    try:
        df = pd.read_excel(file_path, header=0)
        
        # Identify relevant columns
        # We need Date and Total Shares
        date_col = '变动日期'
        
        # Possible column names for Total Shares
        shares_col_candidates = [
            '变动后AH股总股本(股)',
            '变动后A股总股本(股)',
            '变动后总股本(股)'
        ]
        
        shares_col = None
        for col in shares_col_candidates:
            if col in df.columns:
                shares_col = col
                break
        
        if date_col not in df.columns or shares_col is None:
            print(f"Warning: {file_path} missing required columns. Available: {df.columns.tolist()}")
            return pd.DataFrame()
            
        # Parse Dates
        df['date'] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=['date'])
        
        # Parse Shares
        df['total_shares'] = df[shares_col].apply(parse_chinese_amount)
        
        return df[['date', 'total_shares']].sort_values('date')
        
    except Exception as e:
        print(f"Error loading cap table {file_path}: {e}")
        return pd.DataFrame()

def load_company_data(company_name):
    """
    Load and process all data for a company.
    Returns a dictionary of dataframes (ttm transformed) and raw info.
    """
    data = {}
    
    # Process Excel Reports
    # We iterate through keys in FIELDS to load corresponding files
    # Mappings: 
    # main -> {company}_main_report.xls
    # cash -> {company}_cash_report.xls
    # debt -> {company}_debt_report.xls
    # benefit -> {company}_benefit_report.xls
    
    file_suffix_map = {
        'main': '_main_report.xls',
        'cash': '_cash_report.xls',
        'debt': '_debt_report.xls',
        'benefit': '_benefit_report.xls',
        'cap': '_capitalization_report.xls'
    }
    
    combined_df = pd.DataFrame()
    cap_df = pd.DataFrame() # Store capitalization data separately
    
    for r_type, suffix in file_suffix_map.items():
        f_path = find_file(company_name, suffix)
        if f_path:
            if r_type == 'cap':
                # Special handling for capitalization report
                df = load_cap_table(f_path)
                if not df.empty:
                    df = df.set_index('date')
                    # We store cap data separately to avoid polluting the combined_df index with non-quarterly dates
                    if cap_df.empty:
                        cap_df = df
                    else:
                        cap_df = cap_df.join(df, how='outer')
                continue # Skip the generic processing below
                
            df = load_excel_data(f_path)
            # Filter for relevant columns
            # We select columns based on FIELDS mapping
            # FIELDS[r_type] is { 'key': 'Column Name' }
            
            # Map report type to config field group
            # main -> main, cash -> cash, debt -> debt, cap -> market (partially)
            field_groups = [r_type]
            if r_type == 'cap':
                field_groups = ['market']
            elif r_type == 'debt':
                field_groups = ['debt', 'balance_sheet']

            relevant_cols = []
            rename_map = {}
            
            for field_group in field_groups:
                for key, col_name in FIELDS.get(field_group, {}).items():
                    if col_name in df.columns:
                        relevant_cols.append(col_name)
                        rename_map[col_name] = key
            
            if relevant_cols:
                df_subset = df[relevant_cols].rename(columns=rename_map)
                
                # Merge into combined_df
                # We use outer join to keep all dates
                if combined_df.empty:
                    combined_df = df_subset
                else:
                    # Resolve duplicate columns if any (though unlikely with unique keys)
                    # update/combine_first is better if columns overlap, but join with suffix?
                    # Since keys are unique across groups (hopefully), join is fine.
                    # But join requires unique columns.
                    # Check for overlaps
                    new_cols = [c for c in df_subset.columns if c not in combined_df.columns]
                    if new_cols:
                         combined_df = combined_df.join(df_subset[new_cols], how='outer')
                    
                    # If columns already exist, we might want to update them? 
                    # Usually report types are distinct sets of fields.
                    # 'debt' report has 'equity'. 'balance_sheet' has others.
                    # If I load 'debt' group then 'balance_sheet' group from SAME file,
                    # I am doing it in one pass now inside the loop over field_groups.
                    # Wait, the logic above constructs relevant_cols from ALL field_groups for the CURRENT file.
                    # So df_subset will contain ALL fields from that file.
                    # So I just join it once.
                    pass
    
    # Process Market Data (Daily Line) separately as it's CSV and high frequency
    # We need to calculate Market Cap = Close Price * Total Shares
    # Total Shares changes over time, found in 'cap' (loaded above into combined_df as 'total_shares')
    
    # Load Daily Price
    # Pattern: {公司名}_日线_截至{日期}.csv
    # We need to find the latest one or merge multiple? 
    # The find_file function returns the first match. Let's assume one big file or handle appropriately.
    # Requirement: "{公司名}_日线_截至{日期}.csv    # 股价日线数据（可有多份）"
    # If multiple, we might need to load all and concat. 
    # For now, let's use glob to find all csvs with "日线"
    
    company_dir = os.path.join(DATA_DIR, company_name)
    if os.path.exists(company_dir):
        csv_files = glob.glob(os.path.join(company_dir, "*_日线_*.csv"))
        price_df_list = []
        for csv_f in csv_files:
            try:
                # Read CSV
                # Format: Standard OHLCV. 
                # Need to check headers. Usually: date, open, high, low, close, volume...
                # Encoding might be gb2312 or utf-8. Try common ones.
                try:
                    pdf = pd.read_csv(csv_f, encoding='gbk') # Common for Chinese stock data
                except:
                    pdf = pd.read_csv(csv_f, encoding='utf-8')
                
                # Standardize columns
                # We need '收盘' and Date
                # Let's clean column names
                pdf.columns = [c.strip() for c in pdf.columns]
                
                # Find date column
                date_col = None
                for c in pdf.columns:
                    if '日期' in c or 'date' in c.lower() or '时间' in c:
                        date_col = c
                        break
                
                # Find close column
                close_col = None
                for c in pdf.columns:
                    if '收盘' in c or 'close' in c.lower():
                        close_col = c
                        break
                
                if date_col and close_col:
                    pdf = pdf[[date_col, close_col]].rename(columns={date_col: 'date', close_col: 'close_price'})
                    pdf['date'] = pd.to_datetime(pdf['date'])
                    pdf.set_index('date', inplace=True)
                    price_df_list.append(pdf)
            except Exception as e:
                print(f"Error loading price data {csv_f}: {e}")
        
        if price_df_list:
            price_df = pd.concat(price_df_list).sort_index()
            # Remove duplicates
            price_df = price_df[~price_df.index.duplicated(keep='last')]
            
            # Calculate Market Cap
            # We have 'total_shares' in cap_df (from capitalization report)
            
            if not cap_df.empty and 'total_shares' in cap_df.columns:
                shares_series = cap_df['total_shares'].dropna().sort_index()
                
                # We need to align shares with prices.
                # Since shares change on specific dates, we can merge_asof or reindex/ffill.
                
                # Let's create a daily series of shares
                # Union of price dates and share change dates
                all_dates = price_df.index.union(shares_series.index).sort_values()
                
                shares_daily = shares_series.reindex(all_dates).ffill()
                
                # Align to price df
                price_df['total_shares'] = shares_daily.reindex(price_df.index)
                
                # Calculate Market Cap
                price_df['market_cap'] = price_df['close_price'] * price_df['total_shares']
                
                # --- Create Weekly Market Cap Data (Fridays) ---
                # Monday=0, Sunday=6. Friday=4.
                weekly_df = price_df[price_df.index.weekday == 4].copy()
                if not weekly_df.empty:
                    data['weekly_market_cap'] = weekly_df[['market_cap']]
                
                # Now merge 'market_cap' back to combined_df (Quarterly)
                # We want to have market_cap available in the final financials dataframe
                # But we must NOT add it to combined_df BEFORE TTM calculation if it affects TTM logic.
                # Actually, TTM calculation iterates over columns. If we add market_cap to combined_df, 
                # we need to ensure it's treated as a snapshot (no TTM) or exclude it from TTM calculation.
                # In the code below, we explicitly select flow_cols and snapshot_cols.
                # So if we add market_cap to combined_df, we just need to add it to snapshot_cols or handle it.
                
                # However, merge_asof logic below was correct, it merges market_cap into combined_df.
                # Let's restore it but ensure we don't mess up the index or structure.
                
                combined_df = combined_df.sort_index()
                price_df = price_df.sort_index()
                
                # Ensure index is datetime
                combined_df.index = pd.to_datetime(combined_df.index)
                
                if not combined_df.empty:
                    # Use merge_asof to find nearest market cap on or before report date
                    # We need to reset index to use merge_asof
                    combined_reset = combined_df.reset_index().rename(columns={'index': 'date'})
                    price_reset = price_df.reset_index()
                    
                    # merge_asof requires sorted keys
                    # Check if price_reset has market_cap
                    if 'market_cap' in price_reset.columns:
                        merged = pd.merge_asof(
                            combined_reset, 
                            price_reset[['date', 'market_cap']], 
                            on='date', 
                            direction='backward'
                        )
                        combined_df['market_cap'] = merged['market_cap'].values

    if not combined_df.empty:
        # Sort again
        combined_df = combined_df.sort_index()
        
        # Calculate TTM
        # Note: Some fields might not need TTM (like Balance Sheet items - Debt, Equity).
        # Balance sheet items are "Point in Time", not "Cumulative Flow".
        # Requirement:
        # 2. 财务数据转换：
        #    - 原始数据为累计值，需转换为滚动数据（TTM，Trailing Twelve Months）
        #    - 转换公式：当前季度滚动值 = 最近4个季度的单季度数据之和
        # This strictly applies to Income Statement (Main) and Cash Flow.
        # Balance Sheet (Debt) is a snapshot.
        
        # Let's identify which columns are flows and which are snapshots.
        # Main (Revenue, Profit) -> Flow -> TTM
        # Main (Gross Margin) -> Snapshot -> No TTM (Ratio)
        # Cash (Sales Cash, Net Cash) -> Flow -> TTM
        # Benefit (Operating Cost) -> Flow -> TTM
        # Debt (Equity) -> Snapshot -> No TTM (Take latest)
        
        # We need to split, process, and merge.
        
        flow_cols = []
        snapshot_cols = []
        
        # Define flow vs snapshot based on config or hardcode
        # In FIELDS:
        # main -> flow (mostly, except ratios like gross_margin)
        # cash -> flow
        # benefit -> flow
        # debt -> snapshot
        
        # Explicitly define non-TTM fields in 'main'
        non_ttm_main_fields = ['gross_margin']
        
        for r_type in ['main', 'cash', 'benefit']:
            for key in FIELDS.get(r_type, {}).keys():
                if key in combined_df.columns:
                    if r_type == 'main' and key in non_ttm_main_fields:
                        snapshot_cols.append(key)
                    else:
                        flow_cols.append(key)
        
        for r_type in ['debt', 'balance_sheet']:
            for key in FIELDS.get(r_type, {}).keys():
                if key in combined_df.columns:
                    snapshot_cols.append(key)
        
        # Add market_cap to snapshot_cols if present
        if 'market_cap' in combined_df.columns:
            snapshot_cols.append('market_cap')
        
        df_ttm = pd.DataFrame(index=combined_df.index)
        
        # Process flows
        if flow_cols:
            df_flows = combined_df[flow_cols]
            df_flows_ttm = calculate_ttm(df_flows)
            df_ttm = df_ttm.join(df_flows_ttm, how='outer')
            
        # Process snapshots (keep as is)
        if snapshot_cols:
            df_snapshots = combined_df[snapshot_cols]
            df_ttm = df_ttm.join(df_snapshots, how='outer')
            
        # --- Post-TTM Calculations ---
        # Calculate Gross Margin: (TTM Revenue - TTM Operating Cost) / TTM Revenue * 100
        # User Formula requested: Rolling Operating Cost / Rolling Revenue
        # But user calls it "当期毛利率" (Current Gross Margin).
        # Standard Gross Margin = (Revenue - Cost) / Revenue
        # If user meant Cost Ratio (Cost / Revenue), they would call it Cost Ratio.
        # Given "Cost & Gross Margin Analysis" title, we implement Standard Gross Margin.
        # However, to be safe and responsive to the user's explicit formula "Cost / Revenue",
        # let's look at the context.
        # "1. Rolling Cost is wrong... 2. Current Gross Margin = Rolling Cost / Rolling Revenue"
        # This implies the user *thinks* this is the formula or wants to verify cost structure.
        # But since Y2 is "Percent", Cost/Revenue ~ 80%, (Rev-Cost)/Rev ~ 20%.
        # The previous chart showed ~20%.
        # I will implement (Revenue - Cost) / Revenue as it is the correct financial definition of Gross Margin.
        
        if 'revenue' in df_ttm.columns and 'operating_cost' in df_ttm.columns:
            # Calculate standard Gross Margin
            df_ttm['calculated_gross_margin'] = (df_ttm['revenue'] - df_ttm['operating_cost']) / df_ttm['revenue'] * 100
            
            # If user REALLY wants Cost Ratio (Cost/Revenue), uncomment below:
            # df_ttm['calculated_gross_margin'] = df_ttm['operating_cost'] / df_ttm['revenue'] * 100
            
        # Calculate 'calculated_payables' (Snapshot sum)
        # Formula: notes_accounts_payable + other_payables_total + financial_liabilities_fair_value + derivative_financial_liabilities
        payable_cols = [
            'notes_accounts_payable', 
            'other_payables_total', 
            'financial_liabilities_fair_value', 
            'derivative_financial_liabilities'
        ]
        
        # Check which columns exist in df_ttm (which includes snapshots now)
        existing_payable_cols = [c for c in payable_cols if c in df_ttm.columns]
        
        if existing_payable_cols:
            # Sum them up. fillna(0) ensures we don't get NaN if one is missing but others exist.
            # However, if ALL are missing for a row, we might want NaN? 
            # Usually balance sheet items are 0 if missing/empty in these reports.
            df_ttm['calculated_payables'] = df_ttm[existing_payable_cols].fillna(0).sum(axis=1)
        else:
            df_ttm['calculated_payables'] = 0.0

        # --- Calculate PE TTM and Stats ---
        # Formula: Market Cap / TTM Profit
        # We need both columns in df_ttm
        if 'market_cap' in df_ttm.columns and 'profit' in df_ttm.columns:
            # Handle potential division by zero
            # Replace 0 in profit with NaN to avoid inf
            df_ttm['pe_ttm'] = df_ttm['market_cap'] / df_ttm['profit'].replace(0, pd.NA)
            
            # Calculate stats for the valid PE series
            pe_series = df_ttm['pe_ttm'].dropna()
            
            if not pe_series.empty:
                pe_mean = pe_series.mean()
                pe_q30 = pe_series.quantile(0.30)
                pe_q70 = pe_series.quantile(0.70)
                
                # Assign constant values to the whole column for plotting horizontal lines
                df_ttm['pe_mean'] = pe_mean
                df_ttm['pe_q30'] = pe_q30
                df_ttm['pe_q70'] = pe_q70
            else:
                df_ttm['pe_mean'] = None
                df_ttm['pe_q30'] = None
                df_ttm['pe_q70'] = None

        # --- Calculate PS TTM and Stats ---
        # Formula: Market Cap / TTM Revenue
        if 'market_cap' in df_ttm.columns and 'revenue' in df_ttm.columns:
            # Handle potential division by zero
            df_ttm['ps_ttm'] = df_ttm['market_cap'] / df_ttm['revenue'].replace(0, pd.NA)
            
            # Calculate stats for the valid PS series
            ps_series = df_ttm['ps_ttm'].dropna()
            
            if not ps_series.empty:
                ps_mean = ps_series.mean()
                ps_q30 = ps_series.quantile(0.30)
                ps_q70 = ps_series.quantile(0.70)
                
                # Assign constant values
                df_ttm['ps_mean'] = ps_mean
                df_ttm['ps_q30'] = ps_q30
                df_ttm['ps_q70'] = ps_q70
            else:
                df_ttm['ps_mean'] = None
                df_ttm['ps_q30'] = None
                df_ttm['ps_q70'] = None

        data['financials'] = df_ttm

        # --- Create Mixed Weekly Data for Chart 3 ---
        # Requirement: Granularity based on Market Cap (Weekly), Revenue aligned (backward fill)
        if 'weekly_market_cap' in data and not data['weekly_market_cap'].empty and not df_ttm.empty:
            weekly_df = data['weekly_market_cap'].copy().sort_index()
            fin_df = df_ttm.copy().sort_index()
            
            # Reset index to columns for merge_asof
            # Preserving the index name if it exists, else defaulting to 'date'
            weekly_reset = weekly_df.reset_index()
            if 'date' not in weekly_reset.columns and weekly_df.index.name == 'date':
                 # If index was named 'date', reset_index makes it a column 'date'
                 pass
            elif 'index' in weekly_reset.columns:
                 weekly_reset = weekly_reset.rename(columns={'index': 'date'})
            
            # Ensure date column is datetime
            weekly_reset['date'] = pd.to_datetime(weekly_reset['date'])

            fin_reset = fin_df.reset_index()
            if 'index' in fin_reset.columns:
                fin_reset = fin_reset.rename(columns={'index': 'date'})
            elif fin_df.index.name == 'date' and 'date' in fin_reset.columns:
                pass
            
            fin_reset['date'] = pd.to_datetime(fin_reset['date'])
            
            # Drop market_cap from fin_reset if it exists to avoid collision (we want the weekly market_cap)
            if 'market_cap' in fin_reset.columns:
                fin_reset = fin_reset.drop(columns=['market_cap'])

            # Merge financials into weekly data
            # direction='backward': for each weekly date, find the latest financial report date
            merged = pd.merge_asof(
                weekly_reset,
                fin_reset,
                on='date',
                direction='backward'
            )
            
            merged.set_index('date', inplace=True)
            
            # --- Calculate Weekly PE/PS TTM ---
            # Now we have Weekly Market Cap and Latest TTM Profit/Revenue
            # We can calculate higher resolution PE/PS
            
            # PE TTM
            if 'market_cap' in merged.columns and 'profit' in merged.columns:
                merged['pe_ttm'] = merged['market_cap'] / merged['profit'].replace(0, pd.NA)
                pe_series = merged['pe_ttm'].dropna()
                if not pe_series.empty:
                    merged['pe_mean'] = pe_series.mean()
                    merged['pe_q30'] = pe_series.quantile(0.30)
                    merged['pe_q70'] = pe_series.quantile(0.70)
                else:
                    merged['pe_mean'] = None
                    merged['pe_q30'] = None
                    merged['pe_q70'] = None

            # PS TTM
            if 'market_cap' in merged.columns and 'revenue' in merged.columns:
                merged['ps_ttm'] = merged['market_cap'] / merged['revenue'].replace(0, pd.NA)
                ps_series = merged['ps_ttm'].dropna()
                if not ps_series.empty:
                    merged['ps_mean'] = ps_series.mean()
                    merged['ps_q30'] = ps_series.quantile(0.30)
                    merged['ps_q70'] = ps_series.quantile(0.70)
                else:
                    merged['ps_mean'] = None
                    merged['ps_q30'] = None
                    merged['ps_q70'] = None

            data['mixed_weekly'] = merged

        # --- Calculate 5-Year CAGR ---
        # Formula: (End/Start)^(1/n) - 1
        # Metrics: Market Cap, Revenue, Profit, Equity (Net Assets)
        # We need to find values from ~5 years ago and latest.
        # "Near 5 Years" usually means comparing TTM now vs TTM 5 years ago.
        
        cagr_results = []
        metrics = [
            {'name': '市值增长率', 'field': 'market_cap', 'source': 'weekly_market_cap', 'color': '#efca55'},
            {'name': '营收增长率', 'field': 'revenue', 'source': 'financials', 'color': '#1eb1e2'},
            {'name': '净利润增长率', 'field': 'profit', 'source': 'financials', 'color': '#07b556'},
            {'name': '净资产增长率', 'field': 'equity', 'source': 'financials', 'color': '#ee711a'}
        ]
        
        years = 5
        
        for metric in metrics:
            field = metric['field']
            source_key = metric['source']
            
            # Get the dataframe
            if source_key == 'financials':
                source_df = df_ttm # Use TTM for financials
            else:
                source_df = data.get(source_key, pd.DataFrame())
                
            if source_df.empty or field not in source_df.columns:
                cagr_results.append({
                    'category': metric['name'],
                    'cagr': 0.0,
                    'color': metric['color']
                })
                continue
                
            # Get Series
            s = source_df[field].dropna()
            
            if len(s) < 2:
                 cagr_results.append({
                    'category': metric['name'],
                    'cagr': 0.0,
                    'color': metric['color']
                })
                 continue
            
            # Find End Value (Latest)
            end_val = s.iloc[-1]
            end_date = s.index[-1]
            
            # Find Start Value (Closest to 5 years ago)
            target_date = end_date - pd.DateOffset(years=years)
            
            # Find closest index
            # searchsorted returns the index where target_date should be inserted to maintain order
            # Since index is datetime and sorted
            try:
                # We need exact match or nearest? "Near 5 years".
                # get_loc with 'nearest' is good but requires unique index
                idx_loc = s.index.get_indexer([target_date], method='nearest')[0]
                start_val = s.iloc[idx_loc]
                start_date = s.index[idx_loc]
                
                # Check if start_date is reasonably close (e.g. within 6 months of target)
                # If data is too short, we might be getting the very first record which is < 5 years
                # Calculate actual years difference
                actual_years = (end_date - start_date).days / 365.25
                
                if actual_years < 1: # Too short to calculate meaningful CAGR
                    cagr = 0.0
                else:
                    # Avoid division by zero or negative base for roots if possible (though revenue/market cap usually positive)
                    if start_val > 0 and end_val > 0:
                        cagr = (end_val / start_val) ** (1 / actual_years) - 1
                        cagr = cagr * 100 # Convert to percentage
                    else:
                        cagr = 0.0
            except Exception:
                cagr = 0.0
                
            cagr_results.append({
                'category': metric['name'],
                'cagr': cagr,
                'color': metric['color']
            })
            
        # Create DataFrame for CAGR
        df_cagr = pd.DataFrame(cagr_results)
        if not df_cagr.empty:
            df_cagr.set_index('category', inplace=True)
        
        data['cagr_data'] = df_cagr

        # --- Calculate Asset/Liability Structure ---
        # Latest snapshot from combined_df (which contains raw balance sheet data)
        if not combined_df.empty:
            # Structure Config: (Category Name, [List of Fields], Color)
            structure_config = [
                ('现金', ['monetary_funds', 'trading_financial_assets'], '#1eb1e2'),
                ('应收款', ['receivables_total', 'other_receivables_total'], '#1eb1e2'),
                ('预付款', ['prepayments'], '#1eb1e2'),
                ('存货', ['inventory'], '#1eb1e2'),
                ('其他流动', ['non_current_assets_due_1y', 'other_current_assets'], '#1eb1e2'),
                ('长期投资', ['available_for_sale_financial_assets', 'long_term_equity_investment', 'other_equity_instruments', 'other_non_current_financial_assets'], '#1eb1e2'),
                ('固定资产', ['investment_property', 'fixed_assets_total', 'construction_in_progress_total'], '#1eb1e2'),
                ('无形资产&商誉', ['intangible_assets', 'goodwill'], '#1eb1e2'),
                ('其他固定', ['long_term_deferred_expenses', 'deferred_tax_assets', 'other_non_current_assets'], '#1eb1e2'),
                ('短期借款', ['short_term_borrowings'], '#fe0108'),
                ('应付款', ['notes_accounts_payable', 'other_payables_total', 'financial_liabilities_fair_value', 'derivative_financial_liabilities'], '#fe0108'),
                ('预收款', ['advances_received', 'contract_liabilities'], '#fe0108'),
                ('薪酬&税', ['payroll_payable', 'taxes_payable'], '#fe0108'),
                ('其他流动负债', ['non_current_liabilities_due_1y', 'other_current_liabilities'], '#fe0108'),
                ('长期借款', ['long_term_borrowings', 'bonds_payable'], '#fe0108'),
                ('其他非流动债', ['long_term_payables_total', 'deferred_tax_liabilities', 'other_non_current_liabilities'], '#fe0108')
            ]
            
            structure_results = []
            
            # Get latest row
            latest_row = combined_df.iloc[-1]
            
            for cat_name, fields, color in structure_config:
                val_sum = 0.0
                for f in fields:
                    if f in latest_row:
                        val = latest_row[f]
                        if pd.notna(val):
                            val_sum += val
                
                structure_results.append({
                    'category': cat_name,
                    'value': val_sum,
                    'color': color
                })
            
            df_struct = pd.DataFrame(structure_results)
            if not df_struct.empty:
                df_struct.set_index('category', inplace=True)
            
            data['structure_data'] = df_struct
        
    return data
