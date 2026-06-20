import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import sys
import os
import akshare as ak
import requests
import json
from io import BytesIO
import time
import pickle
import xlsxwriter
from pathlib import Path
from bs4 import BeautifulSoup

# Add src to path so we can import modules
sys.path.append(os.path.join(os.path.dirname(__file__)))

from data_loader import get_companies, load_company_data
from visualizer import create_chart
from dowload_data import  download_equity_change_table,download_stock_statements_disk,get_stock_list_disk_cache,download_stock_price_history
from config import CHARTS, COLORS
from config import DATA_DIR
# Page config
st.set_page_config(
    page_title="财务数据可视化分析工具",
    layout="centered",
    initial_sidebar_state="expanded"
)

# 缓存股票列表数据
@st.cache_data(ttl=3600)  # 缓存1小时

# 搜索股票
def search_stocks(search_term, stock_list_df):
    """根据搜索词过滤股票"""
    if search_term and not stock_list_df.empty:
        search_term = str(search_term).upper()
        # 多条件搜索
        mask = (
            stock_list_df['name'].astype(str).str.contains(search_term, case=False) | 
            stock_list_df['code'].astype(str).str.contains(search_term) |
            # 拼音首字母搜索（简单实现）
            stock_list_df['name'].astype(str).apply(lambda x: ''.join([i[0] for i in x.split() if i])).str.contains(search_term.upper())
        )
        return stock_list_df[mask]
    return stock_list_df.head(100)  # 默认返回前100只


def main():

    search_term = st.sidebar.text_input("输入公司名称、股票代码或拼音缩写:", 
                                       placeholder="例如: 贵州茅台 或 600519 或 GZMT")
    selected_stock_name = ''  
    # 加载股票列表
     # 搜索按钮放在输入框下方，用户输入后点击搜索
    with st.spinner("正在加载股票列表..."):
        stock_list_df = get_stock_list_disk_cache()
        if stock_list_df.empty:
            st.error("无法加载股票列表，请检查网络连接或稍后重试")
            return
        
     # 搜索股票
    search_results = search_stocks(search_term, stock_list_df)
    if not search_results.empty:
        st.sidebar.success(f"找到 {len(search_results)} 条相关记录")
        
        
        # 格式化显示
        if 'code' in search_results.columns and 'name' in search_results.columns:
            display_df = search_results[['code', 'name']].copy()
            display_df.columns = ['股票代码', '公司名称']
            
            # 添加序号
            display_df.insert(0, '序号', range(1, len(display_df) + 1))
            
            
            # 让用户选择具体的股票
            # 让用户从列表中选择
            selected_index = st.sidebar.selectbox(
                "已搜索到的股票:", 
                display_df.apply(lambda x: f"{x['序号']}. {x['公司名称']} ({x['股票代码']})", axis=1)
            )

            # 提取选择的股票代码
            if selected_index:
                # 提取股票代码部分
                select_name = selected_index.split('. ')[1].split(' (')[0]
                selected_stock_name = select_name
                selected_code = selected_index.split('(')[-1].rstrip(')')
            st.subheader("财务报表下载")
            if st.sidebar.button("获取并下载财务报表", type="primary"):
                #将选择的股票代码传递给数据加载函数，并下载数据
                download_stock_statements_disk(stock_name=select_name, stock_code=selected_code)
                download_equity_change_table(stock_name=select_name,stock_code=selected_code)
                download_stock_price_history(stock_name=select_name,stock_code=selected_code, period="daily")
         
    else:
        if search_term:
            st.warning(f"未找到包含 '{search_term}' 的股票")
        else:
            st.info("在左侧搜索框中输入公司名称、股票代码或拼音缩写开始搜索")
    

    # 1. Company Selection
    companies = get_companies()
    if not companies:
        st.error("未找到公司数据，请检查数据路径配置。")
        return
    # Sidebar controls
    with st.sidebar:
        st.header("控制面板")
        #自动模式下，读取搜索栏输入的股票代码，自动选择对应公司，并显示对应图表，手动模式下，用户可以从下拉框中选择公司和图表类型
        cur_mode = st.radio("选择模式", options=["自动选择", "手动选择"])
        if cur_mode == "自动选择":
            if selected_stock_name:
                st.success(f"自动选择了: {selected_stock_name}")       
                selected_company = selected_stock_name     
            else:
                st.info("请在左侧搜索并选择股票后，自动模式将显示对应图表")
        else:
            selected_company = st.selectbox(
                "选择公司",
                options=companies
            )
            
        chart_keys = list(CHARTS.keys())
        
        # Initialize session state for chart selection if not present
        if 'selected_chart_key' not in st.session_state:
            st.session_state.selected_chart_key = chart_keys[0]

        # Helper functions for navigation
        def prev_chart():
            current_idx = chart_keys.index(st.session_state.selected_chart_key)
            new_idx = (current_idx - 1) % len(chart_keys)
            st.session_state.selected_chart_key = chart_keys[new_idx]

        def next_chart():
            current_idx = chart_keys.index(st.session_state.selected_chart_key)
            new_idx = (current_idx + 1) % len(chart_keys)
            st.session_state.selected_chart_key = chart_keys[new_idx]
            
        chart_type_key = st.selectbox(
            "选择图表",
            options=chart_keys,
            format_func=lambda x: CHARTS[x]['name'],
            key='selected_chart_key'
        )
        
        # Navigation buttons
        col1, col2 = st.columns(2)
        with col1:
            st.button("⬅️ 上一个", on_click=prev_chart, use_container_width=True)
        with col2:
            st.button("下一个 ➡️", on_click=next_chart, use_container_width=True)
        
        # Load data for selected company
        # Use st.cache_data to improve performance
        # But for now, direct call is fine as data is local and smallish
        with st.spinner('加载数据中...'):
            company_data = load_company_data(selected_company)
        
        # Select data based on chart type
        chart_config = CHARTS[chart_type_key]
        data_source_key = chart_config.get('data_source', 'financials')
        
        if data_source_key == 'financials':
            if 'financials' not in company_data or company_data['financials'].empty:
                st.warning(f"未找到 {selected_company} 的财务数据")
                return
            df = company_data['financials']
        else:
            if data_source_key not in company_data or (isinstance(company_data[data_source_key], pd.DataFrame) and company_data[data_source_key].empty):
                # Specific warning messages for known keys could be kept, but generic is fine
                st.warning(f"未找到 {selected_company} 的数据 ({data_source_key})")
                return
            df = company_data[data_source_key]
        
        # Time Range Control
        # Only for time-series data (where index is datetime)
        is_time_series = isinstance(df.index, pd.DatetimeIndex)
        
        if is_time_series:
            min_date = df.index.min().date()
            max_date = df.index.max().date()
            
            # Default last 10 years
            default_start = max(min_date, max_date - timedelta(days=365*10))
            
            date_range = st.slider(
                "时间范围",
                min_value=min_date,
                max_value=max_date,
                value=(default_start, max_date),
                format="YYYY-MM-DD"
            )
            
            # Filter data based on date range
            mask = (df.index.date >= date_range[0]) & (df.index.date <= date_range[1])
            chart_df = df.loc[mask]
        else:
            # Non-time series data (e.g. Categorical) - use full dataframe
            chart_df = df

    # Apply Custom CSS (Immersive Dark Theme)
    st.markdown(f"""
        <style>
        /* Global Settings */
        .stApp {{
            background-color: {COLORS['background']};
            background-image: radial-gradient(circle at 10% 20%, {COLORS['background']} 0%, {COLORS['sidebar_bg']} 100%);
        }}
        
        /* Sidebar Styling */
        [data-testid="stSidebar"] {{
            background-color: {COLORS['sidebar_bg']};
            border-right: 1px solid rgba(255, 255, 255, 0.1);
        }}
        
        /* Headers */
        h1, h2, h3 {{
            font-family: 'Helvetica Neue', sans-serif;
            font-weight: 600;
            color: {COLORS['white']};
        }}
        
        /* Card Container Style */
        .metric-card {{
            background-color: {COLORS['card_bg']};
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            border: 1px solid rgba(255, 255, 255, 0.05);
        }}
        
        /* Custom Streamlit Elements */
        .stSelectbox > div > div {{
            background-color: {COLORS['card_bg']};
            border: 1px solid rgba(255, 255, 255, 0.1);
            color: {COLORS['white']};
        }}
        
        .stSlider > div > div > div > div {{
            background-color: {COLORS['blue']};
        }}
        
        /* Dataframe styling */
        .dataframe {{
            font-family: 'Consolas', monospace;
        }}
        </style>
        """, unsafe_allow_html=True)

    # Filter data based on slider (Moved inside is_time_series block above)
    filtered_df = chart_df.copy() # Ensure we are working on a copy

    # Dynamic recalculation for PE/PS stats based on selected time range
    # User Request: Change Mean to Median (50% quantile) within the selected time range
    if chart_type_key in ['pe_trend', 'ps_trend'] and is_time_series and not filtered_df.empty:
        prefix = 'pe' if chart_type_key == 'pe_trend' else 'ps'
        series_col = f'{prefix}_ttm'
        
        if series_col in filtered_df.columns:
            # Calculate stats on the visible data
            valid_series = filtered_df[series_col].dropna()
            if not valid_series.empty:
                # Calculate new stats based on visible range
                # "Mean" line becomes Median (50%)
                new_mean = valid_series.quantile(0.50) 
                new_q30 = valid_series.quantile(0.30)
                new_q70 = valid_series.quantile(0.70)
                
                # Update columns in the dataframe to be plotted
                filtered_df[f'{prefix}_mean'] = new_mean
                filtered_df[f'{prefix}_q30'] = new_q30
                filtered_df[f'{prefix}_q70'] = new_q70

    # 2. Main Chart Area
    chart_config = CHARTS[chart_type_key]
    
    # Create and display chart
    # Always use transparent background for custom theme
    is_transparent = True
    fig = create_chart(filtered_df, chart_config, transparent_bg=is_transparent)
    st.plotly_chart(fig, use_container_width=True)

    # 3. Detailed Data View (Optional but helpful)
    with st.expander("查看详细数据"):
        # Identify numeric columns for formatting
        numeric_cols = filtered_df.select_dtypes(include=['float64', 'int64']).columns
        # Apply formatting only to numeric columns
        st.dataframe(filtered_df.style.format({col: "{:.2f}" for col in numeric_cols}))

if __name__ == "__main__":
    main()
