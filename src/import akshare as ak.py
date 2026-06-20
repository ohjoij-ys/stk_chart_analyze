import akshare as ak
from datetime import datetime, timedelta
end_data = datetime.now().strftime("%Y%m%d")
start_date = (datetime.now() - timedelta(days=365*10)).strftime("%Y%m%d")
#stock_zh_a_hist_df = ak.stock_zh_a_hist(symbol="000001", period="daily", start_date="20170301", end_date='20240528', adjust="")
stock_zh_a_hist_df = ak.stock_zh_a_hist(symbol="600276", period="daily", 
                                        start_date=start_date, end_date=end_data  , 
                                        adjust="") 
print(stock_zh_a_hist_df)