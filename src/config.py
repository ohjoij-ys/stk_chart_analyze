import os

# Base paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, '数据导出')

# Colors
COLORS = {
    'yellow': '#FFD700',       # 更亮的金色
    'blue': '#00B4D8',         # 更现代的青蓝色
    'white': '#FFFFFF',
    'orange': '#FF9F1C',
    'red': '#FF595E',
    'grey': '#CACACA',
    'green': '#2EC4B6',
    'background': '#0E1117',   # Streamlit 默认深色背景，更深邃
    'card_bg': '#1E2329',      # 卡片背景色
    'sidebar_bg': '#262730',   # 侧边栏背景
    'blue_v3': '#1eb1e2',
    'yellow_v3': '#efca55'
}

# Field Mappings
# Maps the user-friendly name to the potential column name in the excel file
# Since we don't know the exact internal names, we assume the first column contains these exact strings
FIELDS = {
    'main': {
        'revenue': '营业总收入(元)',
        'profit': '净利润(元)',
        'gross_margin': '销售毛利率'
    },
    'benefit': {
        'operating_cost': '其中：营业成本(元)',
        'sales_expense': '销售费用(元)',
        'manage_expense': '管理费用(元)',
        'rd_expense': '研发费用(元)'
    },
    'cash': {
        'sales_cash': '销售商品、提供劳务收到的现金(元)',
        'net_operate_cash': '经营活动产生的现金流量净额(元)'
    },
    'debt': {
        'equity': '*归属于母公司所有者权益合计(元)'
    },
    'balance_sheet': {
        # Assets
        'monetary_funds': '货币资金(元)',
        'trading_financial_assets': '交易性金融资产(元)',
        'receivables_total': '应收票据及应收账款(元)',
        'other_receivables_total': '其他应收款合计(元)',
        'prepayments': '预付款项(元)',
        'inventory': '存货(元)',
        'non_current_assets_due_1y': '一年内到期的非流动资产(元)',
        'other_current_assets': '其他流动资产(元)',
        'available_for_sale_financial_assets': '可供出售金融资产(元)',
        'long_term_equity_investment': '长期股权投资(元)',
        'other_equity_instruments': '其他权益工具投资(元)',
        'other_non_current_financial_assets': '其他非流动金融资产(元)',
        'investment_property': '投资性房地产(元)',
        'fixed_assets_total': '固定资产合计(元)',
        'construction_in_progress_total': '在建工程合计(元)',
        'intangible_assets': '无形资产(元)',
        'goodwill': '商誉(元)',
        'long_term_deferred_expenses': '长期待摊费用(元)',
        'deferred_tax_assets': '递延所得税资产(元)',
        'other_non_current_assets': '其他非流动资产(元)',
        
        # Liabilities
        'short_term_borrowings': '短期借款(元)',
        'notes_accounts_payable': '应付票据及应付账款(元)',
        'other_payables_total': '其他应付款合计(元)',
        'financial_liabilities_fair_value': '以公允价值计量且其变动计入当期损益的金融负债(元)', # Note space removed to match cleaned data
        'derivative_financial_liabilities': '衍生金融负债(元)',
        'advances_received': '预收款项(元)',
        'contract_liabilities': '合同负债(元)',
        'payroll_payable': '应付职工薪酬(元)',
        'taxes_payable': '应交税费(元)',
        'non_current_liabilities_due_1y': '一年内到期的非流动负债(元)',
        'other_current_liabilities': '其他流动负债(元)',
        'long_term_borrowings': '长期借款(元)',
        'bonds_payable': '应付债券(元)',
        'long_term_payables_total': '长期应付款合计(元)',
        'deferred_tax_liabilities': '递延所得税负债(元)',
        'other_non_current_liabilities': '其他非流动负债(元)'
    },
    'market': {
        'total_shares': '总股本(股)', # Column name in capitalization report
        'close_price': '收盘' # Column name in daily line file
    }
}

# Chart Definitions
CHARTS = {


    'revenue_cash_growth': {
        'name': '总营收与主业现金流增长趋势',
        'type': 'line_multi',
        'series': [
            {'name': '滚动营业总收入', 'field': 'revenue', 'color': COLORS['blue'], 'source': 'main', 'is_ttm': True},
            {'name': '滚动销售商品、提供劳务收到的现金', 'field': 'sales_cash', 'color': COLORS['yellow'], 'source': 'cash', 'is_ttm': True}
        ],
        'y_axis_label': '金额（亿元）'
    },
    'profit_cash_growth': {
        'name': '净利润和现金流净额趋势',
        'type': 'line_multi',
        'series': [
            {'name': '滚动净利润', 'field': 'profit', 'color': COLORS['blue'], 'source': 'main', 'is_ttm': True},
            {'name': '滚动现金流净额', 'field': 'net_operate_cash', 'color': COLORS['yellow'], 'source': 'cash', 'is_ttm': True}
        ],
        'y_axis_label': '金额（亿元）'
    },
    'market_cap_revenue_trend': {
        'name': '市值与业绩趋势',
        'type': 'mixed',
        'dual_axis': True,
        'data_source': 'mixed_weekly',
        'series': [
            {'name': '滚动营业总收入', 'field': 'revenue', 'type': 'bar', 'color': COLORS['blue_v3'], 'source': 'main', 'is_ttm': True, 'y_axis': 'y1'},
            {'name': '公司市值', 'field': 'market_cap', 'type': 'line', 'color': COLORS['yellow_v3'], 'source': 'market', 'is_ttm': False, 'y_axis': 'y2'}
        ],
        'y_axis_label': '业绩金额（亿元）',
        'y_axis_2_label': '市值金额（亿元）'
    },
    'cost_gross_margin_analysis': {
        'name': '成本与毛利率分析',
        'type': 'mixed',
        'dual_axis': True,
        'series': [
            {'name': '滚动营业总收入', 'field': 'revenue', 'type': 'bar', 'color': COLORS['blue_v3'], 'source': 'main', 'is_ttm': True, 'y_axis': 'y1'},
            {'name': '滚动营业总成本', 'field': 'operating_cost', 'type': 'bar', 'color': '#ee711a', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y1'},
            {'name': '当期毛利率', 'field': 'calculated_gross_margin', 'type': 'line', 'color': '#e2e1d8', 'source': 'calculated', 'is_ttm': False, 'y_axis': 'y2', 'is_percent': True}
        ],
        'y_axis_label': '金额（亿元）',
        'y_axis_2_label': '百分比（%）'
    },
    'three_expenses_analysis': {
        'name': '三费与业绩对比分析',
        'type': 'mixed',
        'dual_axis': True,
        'series': [
            {'name': '滚动营业总收入', 'field': 'revenue', 'type': 'bar', 'color': '#1eb1e2', 'source': 'main', 'is_ttm': True, 'y_axis': 'y1'},
            {'name': '滚动销售费用', 'field': 'sales_expense', 'type': 'line', 'color': '#fe0108', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y2'},
            {'name': '滚动管理费用', 'field': 'manage_expense', 'type': 'line', 'color': '#e2e1d8', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y2'},
            {'name': '滚动研发费用', 'field': 'rd_expense', 'type': 'line', 'color': '#efca55', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y2'}
        ],
        'y_axis_label': '总营收金额（亿元）',
        'y_axis_2_label': '费用金额（亿元）'
    },
    'five_year_growth_analysis': {
        'name': '近五年综合增长率',
        'type': 'bar',
        'data_source': 'cagr_data',
        'series': [
            {'name': '5年复合增长率', 'field': 'cagr', 'type': 'bar', 'color_field': 'color', 'is_percent': True}
        ],
        'y_axis_label': '年化增长率（%）'
    },
    'asset_liability_structure': {
        'name': '资产负债结构',
        'type': 'bar',
        'data_source': 'structure_data',
        'series': [
            {'name': '金额', 'field': 'value', 'type': 'bar', 'color_field': 'color', 'show_values': True}
        ],
        'y_axis_label': '金额（亿元）'
    },
    'revenue_payables_trend': {
        'name': '总营收和应付账款对比',
        'type': 'mixed',
        'series': [
            {'name': '滚动营业总收入', 'field': 'revenue', 'type': 'bar', 'color': '#1eb1e2', 'source': 'main', 'is_ttm': True, 'y_axis': 'y1'},
            {'name': '滚动应付款', 'field': 'calculated_payables', 'type': 'line', 'color': '#efca55', 'source': 'calculated', 'is_ttm': False, 'y_axis': 'y1'}
        ],
        'y_axis_label': '金额（亿元）'
    },
    'pe_trend': {
        'name': '市盈率趋势',
        'type': 'line_multi',
        'data_source': 'mixed_weekly',
        'series': [
            {'name': '滚动市盈率', 'field': 'pe_ttm', 'color': '#f9f9f9', 'source': 'mixed_weekly'},
            {'name': '区间中位数', 'field': 'pe_mean', 'color': '#efca55', 'source': 'mixed_weekly', 'dash': 'dash'},
            {'name': '30%分位数', 'field': 'pe_q30', 'color': '#07b556', 'source': 'mixed_weekly', 'dash': 'dot'},
            {'name': '70%分位数', 'field': 'pe_q70', 'color': '#fe0108', 'source': 'mixed_weekly', 'dash': 'dot'}
        ],
        'y_axis_label': '市盈率（倍）'
    },
    'ps_trend': {
        'name': '市销率趋势',
        'type': 'line_multi',
        'data_source': 'mixed_weekly',
        'series': [
            {'name': '滚动市销率', 'field': 'ps_ttm', 'color': '#f9f9f9', 'source': 'mixed_weekly'},
            {'name': '区间中位数', 'field': 'ps_mean', 'color': '#efca55', 'source': 'mixed_weekly', 'dash': 'dash'},
            {'name': '30%分位数', 'field': 'ps_q30', 'color': '#07b556', 'source': 'mixed_weekly', 'dash': 'dot'},
            {'name': '70%分位数', 'field': 'ps_q70', 'color': '#fe0108', 'source': 'mixed_weekly', 'dash': 'dot'}
        ],
        'y_axis_label': '市销率（倍）'
    },
    'multi_pe_trend': {
        'name': '多公司滚动市盈率趋势对比',
        'type': 'line_multi',
        'data_source': 'mixed_weekly',
        'series': [
            {'name': '滚动市盈率', 'field': 'pe_ttm', 'color': '#f9f9f9', 'source': 'mixed_weekly'},
            {'name': '滚动市盈率', 'field': 'pe_ttm', 'color': '#519aba', 'source': 'mixed_weekly'},
            {'name': '滚动市盈率', 'field': 'pe_ttm', 'color': '#0078d4', 'source': 'mixed_weekly'},
            {'name': '滚动市盈率', 'field': 'pe_ttm', 'color': '#5ce488', 'source': 'mixed_weekly'},
            {'name': '区间中位数', 'field': 'pe_mean', 'color': '#efca55', 'source': 'mixed_weekly', 'dash': 'dash'},
            {'name': '30%分位数', 'field': 'pe_q30', 'color': '#07b556', 'source': 'mixed_weekly', 'dash': 'dot'},
            {'name': '70%分位数', 'field': 'pe_q70', 'color': '#fe0108', 'source': 'mixed_weekly', 'dash': 'dot'}
        ],
        'y_axis_label': '市盈率（倍）'
    },

    # #新加部分
    # 'R&D_spend_analysis': {
    #     'name': '研发费用与业绩对比分析',
    #     'type': 'mixed',
    #     'dual_axis': True,
    #     'series': [
    #         {'name': '滚动净利润', 'field': 'profit', 'type': 'bar', 'color': '#1eb1e2', 'source': 'main', 'is_ttm': True, 'y_axis': 'y1'},
    #         {'name': '', 'field': 'sales_expense', 'type': 'line', 'color': '#fe0108', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y2'},
    #         {'name': '', 'field': 'manage_expense', 'type': 'line', 'color': '#e2e1d8', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y2'},
    #         {'name': '', 'field': 'rd_expense', 'type': 'line', 'color': '#efca55', 'source': 'benefit', 'is_ttm': True, 'y_axis': 'y2'}
    #     ],
    #     'y_axis_label': '总营收金额（亿元）',
    #     'y_axis_2_label': '费用金额（亿元）'
    # },
}
