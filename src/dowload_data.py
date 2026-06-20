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
from config import DATA_DIR
import streamlit as st
import pandas as pd

#股价历史数据（日线）
def download_stock_price_history(stock_name,stock_code, file_path=DATA_DIR, period="daily", start_date=None, end_date=None):
    """
    下载股票历史价格数据（日线）从AKShare获取
    """
    stock_dir = Path(file_path) / stock_name
    stock_dir.mkdir(parents=True, exist_ok=True)
    
    try:  
        with st.spinner(f"正在获取 {stock_name}({stock_code}) 的{period}历史数据..."):
            # 1. 设置默认日期范围
            if end_date is None:
                end_date = datetime.now().strftime("%Y%m%d")
            if start_date is None:
                # 默认获取最近十年的数据
                start_date = (datetime.now() - timedelta(days=365 * 10)).strftime("%Y%m%d")
            
            # 2. 尝试获取数据
            df = None
            try:
                # 优先尝试深圳市场代码
                df = ak.stock_zh_a_hist_tx(symbol=f'sz{stock_code}', 
                                           start_date=start_date, 
                                           end_date=end_date, 
                                           adjust="")
            except Exception as e_sz:
                st.info(f"尝试深圳代码(sz{stock_code})失败: {e_sz}")
                try:
                    # 尝试上海市场代码
                    df = ak.stock_zh_a_hist_tx(symbol=f'sh{stock_code}', 
                                               start_date=start_date, 
                                               end_date=end_date, 
                                               adjust="")
                except Exception as e_sh:
                    st.error(f"尝试上海代码(sh{stock_code})也失败: {e_sh}")
                    return None
            
            # 3. 检查数据是否为空
            if df is None or df.empty:
                st.warning(f"未获取到 {stock_code} 的历史数据")
                return None
            
            # 4. 查看并统一重命名列名为中文（核心修复）
            st.info(f"原始数据形状: {df.shape}, 原始列名: {list(df.columns)}")
            
            # 定义一个全面的列名映射字典（从各种可能的原始列名 -> 目标中文列名）
            # 这确保了无论AKShare返回什么列名，我们都能将其转换为统一的中文。
            column_mapping_to_chinese = {
                # 可能的原始英文列名
                'date': '日期',
                'open': '开盘',
                'high': '最高', 
                'low': '最低',
                'close': '收盘',
                'volume': '成交量',
                'amount': '成交额',
                'change': '涨跌额',
                'pct_chg': '涨跌幅',
                'amplitude': '振幅',
                'turnover': '换手率',
                # 可能的原始中文列名（确保它们也被正确识别）
                '日期': '日期',
                '开盘': '开盘',
                '最高': '最高',
                '最低': '最低',
                '收盘': '收盘',
                '成交量': '成交量',
                '成交额': '成交额',
                '涨跌额': '涨跌额',
                '涨跌幅': '涨跌幅',
                '振幅': '振幅',
                '换手率': '换手率',
            }
            
            # 创建实际的重命名字典：只映射DataFrame中实际存在的列
            rename_dict = {}
            for old_col in df.columns:
                if old_col in column_mapping_to_chinese:
                    rename_dict[old_col] = column_mapping_to_chinese[old_col]
                else:
                    # 如果遇到未知列名，保留原样并给出提示
                    st.warning(f"发现未映射的列名: '{old_col}'，将保留原名称。")
                    # 如果想强制保留，可以取消下面这行的注释
                    # rename_dict[old_col] = old_col 
            
            # 应用重命名
            if rename_dict:  # 仅在重命名字典不为空时执行
                df.rename(columns=rename_dict, inplace=True)
                st.success("列名重命名完成。")
            else:
                st.warning("未对任何列进行重命名，列名可能已符合要求或完全未知。")
            
            # 5. 确保“日期”列是datetime类型（如果存在）
            if '日期' in df.columns:
                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
            else:
                st.error("处理后的数据中未找到‘日期’列，请检查列名映射。")
                st.write("当前列名:", list(df.columns))
                return None
            
            # 6. 保存数据
            csv_path = stock_dir / f"{stock_name}_{stock_code}_日线_数据_{start_date}_至_{end_date}.csv"
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            
            st.success(f"✅ 成功获取 {stock_name}({stock_code}) 的{period}历史数据")
            st.write(f"- 数据范围: {df['日期'].min().date()} 到 {df['日期'].max().date()}")
            st.write(f"- 数据条数: {len(df)} 条")
            st.write(f"- 数据列: {list(df.columns)}")
            st.write(f"- CSV文件已保存到: {csv_path}")
            
            # 可选：在界面中预览数据
            with st.expander("点击预览前10行数据"):
                st.dataframe(df.head(10))
            
            return df
            
    except Exception as e:
        st.error(f"获取股票历史数据失败: {str(e)}")
        import traceback
        st.error(traceback.format_exc())  # 打印详细错误栈，便于调试
        return None




#股本变动
def download_equity_change_table(stock_name,stock_code,file_path=DATA_DIR):
    stock_dir = Path(file_path) / stock_name
    """
    获取AH股历次股本变动表格数据
    """
    # 1. 设置请求头（模拟浏览器）
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': f'https://basic.10jqka.com.cn/{stock_code}/',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    }
    
    # 2. 请求页面
    url = f"https://basic.10jqka.com.cn/{stock_code}/equity.html"
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.encoding = 'gbk'
        
        if response.status_code != 200:
            st.error(f"请求失败，状态码: {response.status_code}")
            return None
            
        # 3. 解析HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 4. 查找表格
        # 方法1：通过id查找表格容器
        table_container = soup.find('div', id='astockchange')
        table = None
        
        if table_container:
            # 在容器内查找表格
            table = table_container.find('table')
        
        # 方法2：如果方法1找不到，尝试通过class查找
        if table is None:
            table = soup.find('table', class_='m_table')
        
        # 方法3：如果还找不到，尝试通过标题查找
        if table is None:
            for h2 in soup.find_all(['h2', 'h3']):
                if 'AH股历次股本变动' in h2.text:
                    table = h2.find_next('table')
                    break
        
        if table is None:
            st.error("未找到表格")
            return None
        # 5. 提取表头
        headers_list = []
        thead = table.find('thead')
        
        if thead:
            for th in thead.find_all('th'):
                headers_list.append(th.text.strip())
        else:
            # 如果没有thead，尝试从第一个tr获取
            first_row = table.find('tr')
            if first_row:
                for th in first_row.find_all(['th', 'td']):
                    headers_list.append(th.text.strip())
        
        # 如果仍然没有表头，使用已知的表头
        if not headers_list:
            headers_list = ['变动日期', '变动原因', '变动后AH股总股本(股)', 
                           '变动后流通A股(股)', '变动后限售A股(股)']
        
        # 6. 提取数据行 - 关键修正
        data = []
        
        # 查找所有数据行，无论是否显示
        # 使用CSS选择器查找所有类包含J_pageritem的tr元素
        rows = table.select('tr.J_pageritem')
        
        # 如果上面没找到，尝试另一种方式
        if not rows:
            rows = table.find_all('tr')
        
        for row in rows:
            # 跳过表头行
            if row.find('th'):
                continue
                
            cells = row.find_all('td')
            if len(cells) >= 3:  # 至少3列
                row_data = [cell.text.strip() for cell in cells]
                data.append(row_data)
        
        # 7. 创建DataFrame
        if data:
            # 如果提取到的列数与表头数不匹配，调整表头
            if len(data[0]) != len(headers_list):
                headers_list = [f'列{i+1}' for i in range(len(data[0]))]
            
            df = pd.DataFrame(data, columns=headers_list)
            
            # 8. 保存到磁盘
            # 确保目录存在
            os.makedirs(stock_dir, exist_ok=True)
            
            # 保存为CSV和Excel
            csv_path = os.path.join(stock_dir, f"{stock_code}_capitalization_report.csv")
            excel_path = os.path.join(stock_dir, f"{stock_code}_capitalization_report.xls")
            
            # 先保存为CSV
            df.to_csv(csv_path, index=False, encoding='gbk')
            
            # 尝试保存为Excel
            try:
                df.to_excel(excel_path, index=False,engine='xlsxwriter')
                st.success(f"股本变动数据提取成功！\nCSV保存到: {csv_path}\nExcel保存到: {excel_path}")
            except Exception as e:
                st.warning(f"CSV保存成功，但Excel保存失败: {e}\n数据已保存到xls: {csv_path}")
            
            return df
        else:
            st.error("未提取到有效数据")
            return None
    except Exception as e:
        st.error(f"获取数据失败: {str(e)}")
        return None            

def download_stock_statements_disk(stock_name,stock_code,file_path=DATA_DIR):
    """
    下载财报数据到磁盘
    """

    """
    下载财报数据到磁盘
    返回: 成功下载的文件列表
    """
    # 1. 创建股票目录
    stock_dir = Path(file_path) / stock_name
    stock_dir.mkdir(parents=True, exist_ok=True)

    # 2. 定义要下载的报表类型和对应的文件名
    reports = {
        "cash": {
            "name": "现金流量表",
            "export": "cash",
            "filename": f"{stock_code}_cash_report.xls"
        },
        "debt": {
            "name": "资产负债表", 
            "export": "debt",
            "filename": f"{stock_code}_debt_report.xls"
        },
        "main": {
            "name": "主要指标表",
            "export": "main", 
            "filename": f"{stock_code}_main_report.xls"
        },
        "benefit": {
            "name": "利润表",
            "export": "benefit",
            "filename": f"{stock_code}_benefit_report.xls"
        }
    }
    
    # 3. 模拟浏览器的请求头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': f'https://stockpage.10jqka.com.cn/{stock_code}/finance/',  # 关键：设置正确的来源页面
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Cache-Control': 'max-age=0',
    }
    downloaded_files = []
    # 4. 使用Streamlit进度条
    progress_bar = st.progress(0)
    status_text = st.empty()
    for i, (report_type, report_info) in enumerate(reports.items()):
        try:
            # 更新进度
            progress = (i + 1) / len(reports)
            progress_bar.progress(progress)
            status_text.text(f"正在下载 {report_info['name']}...")
            
            # 构建完整的文件路径
            filepath = stock_dir / report_info['filename']
            
            # # 如果文件已存在，跳过下载
            # if filepath.exists():
            #     st.info(f"{report_info['name']} 已存在，跳过下载")
            #     downloaded_files.append(str(filepath))
            #     continue
            
            # 构建下载URL
            url = f"https://basic.10jqka.com.cn/api/stock/export.php"
            params = {
                "export": report_info['export'],
                "type": "report", 
                "code": stock_code
            }
            
            # 发送请求
            response = requests.get(
                url, 
                params=params, 
                headers=headers, 
                stream=True,
                timeout=30
            )
            # 检查响应状态
            if response.status_code == 200:
                # 保存文件
                with open(filepath, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                
                downloaded_files.append(str(filepath))
                st.success(f"{report_info['name']} 下载成功")
                
            elif response.status_code == 403:
                st.error(f"下载 {report_info['name']} 失败: 403 Forbidden (被网站拒绝)")
                st.info("可能需要添加更多请求头或使用Cookie")
            else:
                st.error(f"下载 {report_info['name']} 失败: HTTP {response.status_code}")
                
        except requests.exceptions.Timeout:
            st.error(f"下载 {report_info['name']} 超时")
        except Exception as e:
            st.error(f"下载 {report_info['name']} 失败: {e}")
        
        # 短暂延迟，避免请求过快
        time.sleep(1)
    
    # 5. 完成进度
    progress_bar.progress(1.0)
    status_text.text("下载完成!")
    
    



def get_stock_list_disk_cache(cache_file="stock_list_cache.pkl", max_age_hours=24):
    """
    使用磁盘文件缓存
    """
    # 检查缓存文件是否存在且未过期
    if os.path.exists(cache_file):
        file_mtime = datetime.fromtimestamp(os.path.getmtime(cache_file))
        if datetime.now() - file_mtime < timedelta(hours=max_age_hours):
            try:
                with open(cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    st.sidebar.info(f"从缓存加载 ({len(cached_data)} 只股票)")
                    return cached_data
            except:
                pass
    
    # 重新获取数据
    with st.spinner("正在从网络获取股票列表..."):
        try:
            stock_list = ak.stock_info_a_code_name()
            if not stock_list.empty:
                # 保存到缓存文件
                with open(cache_file, 'wb') as f:
                    pickle.dump(stock_list, f)
                st.sidebar.success(f"获取并缓存 {len(stock_list)} 只股票")
                return stock_list
        except Exception as e:
            st.error(f"获取股票列表失败: {e}")
    
    return pd.DataFrame(columns=["code", "name"])
def get_stock_list_from_source():
    """
    从多个数据源获取股票列表
    返回: 包含股票代码、名称、市场等信息的DataFrame
    """
    try:
        # 方法1: 使用AKShare获取A股列表
        stock_info_a_code_name_df = ak.stock_info_a_code_name()
        if not stock_info_a_code_name_df.empty:
            return stock_info_a_code_name_df
    except Exception as e:
        st.sidebar.warning(f"AKShare获取股票列表失败: {e}")
    
    # 返回空的DataFrame
    return pd.DataFrame(columns=["code", "name"])
