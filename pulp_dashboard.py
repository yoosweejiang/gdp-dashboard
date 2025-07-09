import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta
try:
    from statsmodels.tsa.seasonal import seasonal_decompose
    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

st.set_page_config(page_title="Pulp Data Dashboard", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
    .css-1d391kg {width: 250px !important;}
    .stDataFrame table th:first-child, .stDataFrame table td:first-child {width: 30% !important; min-width: 150px !important;}
    .stDataFrame table th:not(:first-child), .stDataFrame table td:not(:first-child) {width: 17.5% !important; text-align: center !important;}
</style>
""", unsafe_allow_html=True)

# Configuration
TRANSLATIONS = {
    'zh': {'title': '📊 纸浆市场数据仪表板', 'upload': '上传Excel文件', 'time_period': '时间段',
           'last_5_years': '过去5年', 'last_24_months': '过去24个月', 'last_12_months': '过去12个月', 'last_6_months': '过去6个月', 'last_1_month': '过去1个月',
           'price_trends': '价格趋势', 'statistics': '统计数据', 'correlation': '相关性分析', 'trend_seasonality': '趋势与季节性分析', 'data_preview': '数据预览',
           'product': '产品', 'mean': '平均值', 'median': '中位数', 'std_dev': '标准差', 'min': '最小值', 'max': '最大值',
           'select_product': '选择产品', 'trend_direction': '趋势方向', 'seasonal_range': '季节性波动', 'residual_std': '残差标准差', 'upward': '上升', 'downward': '下降'},
    'en': {'title': '📊 Pulp Market Data Dashboard', 'upload': 'Upload Excel file', 'time_period': 'Time Period',
           'last_5_years': 'Last 5 Years', 'last_24_months': 'Last 24 Months', 'last_12_months': 'Last 12 Months', 'last_6_months': 'Last 6 Months', 'last_1_month': 'Last 1 Month',
           'price_trends': 'Price Trends', 'statistics': 'Statistics', 'correlation': 'Correlation Analysis', 'trend_seasonality': 'Trend & Seasonality Analysis', 'data_preview': 'Data Preview',
           'product': 'Product', 'mean': 'Mean', 'median': 'Median', 'std_dev': 'Std Dev', 'min': 'Min', 'max': 'Max',
           'select_product': 'Select Product', 'trend_direction': 'Trend Direction', 'seasonal_range': 'Seasonal Range', 'residual_std': 'Residual Std', 'upward': 'Upward', 'downward': 'Downward'}
}

COLUMNS = {
    'suzano': '山东市场Suzano阔叶浆（金鱼）日度市场价（元/吨）',
    'westfraser': '山东市场WestFraser化机浆（昆河）日度市场价（元/吨）',
    'white_card': '山东市场白卡纸（250-400g）日度市场价（元/吨）',
    'double_offset': '山东市场双胶纸（70g天阳）日度市场价（元/吨）'
}

KEY_TERMS = {
    'suzano': ['Suzano', '阔叶浆', '金鱼'],
    'westfraser': ['WestFraser', '化机浆', '昆河'],
    'white_card': ['白卡纸', '250-400g'],
    'double_offset': ['双胶纸', '70g', '天阳', '华夏太阳', '广东市场']
}

TIME_PERIODS = {
    'last_5_years': 5*365, 'last_24_months': 24*30, 'last_12_months': 12*30, 'last_6_months': 180, 'last_1_month': 30
}

# Initialize session state
for key in ['df_data', 'available_cols', 'date_col']:
    if key not in st.session_state:
        st.session_state[key] = None if key in ['df_data', 'date_col'] else {}

# Sidebar setup
with st.sidebar:
    lang_code = 'zh' if st.selectbox('Language / 语言', ['中文', 'English']) == '中文' else 'en'
    t = TRANSLATIONS[lang_code]
    st.markdown("---")
    uploaded_file = st.file_uploader(t['upload'], type=['xlsx', 'xls'])
    period = st.selectbox(t['time_period'], [t['last_5_years'], t['last_24_months'], t['last_12_months'], t['last_6_months'], t['last_1_month']])

st.title(t['title'])

def find_matching_columns(df_columns):
    """Find matching columns based on exact match or key terms"""
    available_cols = {}
    for key, col_name in COLUMNS.items():
        exact_match = [col for col in df_columns if col_name == col]
        if exact_match:
            available_cols[key] = exact_match[0]
        else:
            for col in df_columns:
                if any(term in str(col) for term in KEY_TERMS[key]):
                    available_cols[key] = col
                    break
    return available_cols

def get_filtered_data(df, date_col, period, t):
    """Filter data based on selected time period"""
    period_key = [k for k, v in t.items() if v == period][0]
    days = TIME_PERIODS[period_key]
    start_date = df[date_col].max() - timedelta(days=days)
    return df[df[date_col] >= start_date].copy()

def create_chart(data, date_col, available_cols, title, colors):
    """Create price trends chart"""
    fig = go.Figure()
    for i, (key, col) in enumerate(available_cols.items()):
        if col in data.columns:
            fig.add_trace(go.Scatter(
                x=data[date_col], y=data[col], mode='lines',
                name=COLUMNS[key], line=dict(color=colors[i % len(colors)])
            ))
    fig.update_layout(
        title=title, xaxis_title="Date", yaxis_title="Price (元/吨)", height=500, hovermode='x unified',
        legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5)
    )
    return fig

def create_stats_table(data, available_cols, t):
    """Create statistics table"""
    stats_data = []
    for key, col in available_cols.items():
        if col in data.columns:
            series = data[col].dropna()
            stats_data.append({
                t['product']: COLUMNS[key], t['mean']: f"{series.mean():.2f}",
                t['median']: f"{series.median():.2f}", t['std_dev']: f"{series.std():.2f}",
                t['min']: f"{series.min():.2f}", t['max']: f"{series.max():.2f}"
            })
    return pd.DataFrame(stats_data)

def create_correlation_analysis(data, available_cols, t):
    """Create correlation heatmap and table"""
    corr_cols = [col for col in available_cols.values() if col in data.columns]
    if len(corr_cols) < 2:
        return None, None
    
    corr_data = data[corr_cols].corr()
    short_names = {col: key for key, col in available_cols.items() if col in corr_cols}
    corr_data = corr_data.rename(index=short_names, columns=short_names)
    
    fig_corr = px.imshow(corr_data, text_auto='.3f', aspect="auto",
                        title=f"{t['correlation']}", color_continuous_scale='RdBu_r')
    fig_corr.update_layout(height=500)
    return fig_corr, corr_data

def create_trend_analysis(data, date_col, selected_product, available_cols, t):
    """Create trend and seasonality analysis"""
    if not HAS_STATSMODELS:
        return None, None
    
    analysis_col = available_cols[selected_product]
    monthly_data = data.set_index(date_col)[analysis_col].resample('M').mean().dropna()
    
    if len(monthly_data) < 24:
        return None, None
    
    try:
        decomposition = seasonal_decompose(monthly_data, model='additive', period=12)
        
        fig_decomp = go.Figure()
        fig_decomp.add_trace(go.Scatter(x=monthly_data.index, y=monthly_data.values, mode='lines', name='Original', line=dict(color='#1f77b4')))
        fig_decomp.add_trace(go.Scatter(x=decomposition.trend.index, y=decomposition.trend.values, mode='lines', name='Trend', line=dict(color='#ff7f0e')))
        fig_decomp.add_trace(go.Scatter(x=decomposition.seasonal.index, y=decomposition.seasonal.values, mode='lines', name='Seasonal', line=dict(color='#2ca02c')))
        
        fig_decomp.update_layout(
            title=f"Time Series Decomposition - {selected_product}", xaxis_title="Date", yaxis_title="Price (元/吨)",
            height=400, hovermode='x unified', legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5)
        )
        
        metrics = {
            'trend_direction': t['upward'] if decomposition.trend.dropna().iloc[-1] > decomposition.trend.dropna().iloc[0] else t['downward'],
            'seasonal_range': f"{decomposition.seasonal.max() - decomposition.seasonal.min():.2f}",
            'residual_std': f"{decomposition.resid.std():.2f}"
        }
        
        return fig_decomp, metrics
    except Exception:
        return None, None

# Process uploaded file
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        date_col = df.columns[0]
        df[date_col] = pd.to_datetime(df[date_col])
        df = df.sort_values(date_col)
        
        st.session_state.df_data = df
        st.session_state.available_cols = find_matching_columns(df.columns)
        st.session_state.date_col = date_col
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.stop()

# Main dashboard
if st.session_state.df_data is not None:
    df, available_cols, date_col = st.session_state.df_data, st.session_state.available_cols, st.session_state.date_col
    
    if not available_cols:
        st.error("No matching columns found. Please check your Excel format.")
        st.stop()
    
    current_df = get_filtered_data(df, date_col, period, t)
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']
    
    # Price Trends
    st.header(f"📊 {t['price_trends']} - {period}")
    fig = create_chart(current_df, date_col, available_cols, f"{t['price_trends']} - {period}", colors)
    st.plotly_chart(fig, use_container_width=True)
    
    # Statistics
    st.header(f"📈 {t['statistics']} - {period}")
    stats_df = create_stats_table(current_df, available_cols, t)
    st.dataframe(stats_df, use_container_width=True)
    
    # Correlation Analysis
    st.header(f"🔗 {t['correlation']} - {period}")
    fig_corr, corr_data = create_correlation_analysis(current_df, available_cols, t)
    if fig_corr:
        st.plotly_chart(fig_corr, use_container_width=True)
        st.dataframe(corr_data.round(3), use_container_width=True)
    else:
        st.warning("Need at least 2 columns for correlation analysis")
    
    # Trend & Seasonality Analysis
    if HAS_STATSMODELS and len(current_df) > 24:
        st.header(f"📈 {t['trend_seasonality']} - {period}")
        product_options = [key for key in available_cols.keys() if available_cols[key] in current_df.columns]
        if product_options:
            selected_product = st.selectbox(t['select_product'], product_options)
            fig_decomp, metrics = create_trend_analysis(current_df, date_col, selected_product, available_cols, t)
            if fig_decomp:
                st.plotly_chart(fig_decomp, use_container_width=True)
                col1, col2, col3 = st.columns(3)
                col1.metric(t['trend_direction'], metrics['trend_direction'])
                col2.metric(t['seasonal_range'], metrics['seasonal_range'])
                col3.metric(t['residual_std'], metrics['residual_std'])
            else:
                st.warning("Need at least 24 months of data for seasonality analysis")
    elif not HAS_STATSMODELS:
        st.info("Install statsmodels package: pip install statsmodels")
    
    # Data Preview and Summary
    with st.expander(f"📋 {t['data_preview']}"):
        st.dataframe(current_df.head(20), use_container_width=True)
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Records", len(current_df))
    col2.metric("Date Range", f"{current_df[date_col].min().strftime('%Y-%m-%d')} to {current_df[date_col].max().strftime('%Y-%m-%d')}")
    col3.metric("Available Products", len(available_cols))
    col4.metric("Time Period", period)

else:
    st.info(f"👆 {t['upload']}")
    st.subheader("Expected Excel Format:")
    sample_data = {'Date': ['2024-01-01', '2024-01-02', '2024-01-03'],
                   **{col: [5200+i*50, 5250+i*50, 5180+i*50] for i, col in enumerate(COLUMNS.values())}}
    st.dataframe(pd.DataFrame(sample_data))