import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# ==========================================
# 頁面設定
# ==========================================
st.set_page_config(
    page_title="所有料件庫存清單",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 頁面標題
# ==========================================
st.markdown('<h1 style="color: #1f77b4; text-align: center;">所有料件庫存清單</h1>', unsafe_allow_html=True)
st.markdown("---")

# ==========================================
# 資料載入（使用快取）
# ==========================================
@st.cache_data
def load_all_materials():
    """載入所有料件資料"""
    try:
        # 讀取 Excel 的「所有料件」工作表
        df = pd.read_excel('物料整合報表.xlsx', sheet_name='所有料件')
        
        # 選擇需要的欄位
        required_cols = ['料號', '品名', '單位', '倉庫', '庫存量']
        
        # 檢查欄位是否存在
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"缺少以下欄位：{', '.join(missing_cols)}")
            return None
        
        # 保留指定欄位
        df = df[required_cols].copy()
        
        # 處理庫存量：填補空值為 0，轉換為整數
        df['庫存量'] = pd.to_numeric(df['庫存量'], errors='coerce').fillna(0).astype(int)
        
        return df
    
    except FileNotFoundError:
        st.error("找不到 '物料整合報表.xlsx' 檔案")
        return None
    except Exception as e:
        st.error(f"讀取檔案失敗: {e}")
        return None

@st.cache_data
def load_delivery_period_data():
    """載入包含交貨期的料件資料"""
    try:
        # 讀取包含交貨期的工作表
        df = pd.read_excel('物料整合報表.xlsx', sheet_name='整合資料')
        
        # 選擇需要的欄位
        if '交貨期(L)' not in df.columns:
            return None
        
        # 保留需要的欄位
        df = df[['料號', '品名', '交貨期(L)']].copy()
        
        # 填補交貨期空值為 30
        df['交貨期(L)'] = pd.to_numeric(df['交貨期(L)'], errors='coerce').fillna(30).astype(int)
        
        return df
    
    except Exception:
        return None

# 載入資料
with st.spinner("正在載入資料..."):
    df = load_all_materials()
    delivery_df = load_delivery_period_data()

if df is None:
    st.stop()

# ==========================================
# 交貨期分佈圓餅圖
# ==========================================
if delivery_df is not None and len(delivery_df) > 0:
    st.subheader("📊 交貨期分佈統計")
    
    # 定義交貨期區間
    bins = [0, 7, 14, 30, 60, 90, float('inf')]
    labels = ['0-7天', '8-14天', '15-30天', '31-60天', '61-90天', '90天以上']
    
    # 分類交貨期
    delivery_df['交貨期區間'] = pd.cut(delivery_df['交貨期(L)'], bins=bins, labels=labels, right=True)
    
    # 計算各區間的料件數和比例
    delivery_stats = delivery_df['交貨期區間'].value_counts().sort_index()
    
    # 創建圓餅圖
    fig = go.Figure(data=[go.Pie(
        labels=delivery_stats.index.astype(str),
        values=delivery_stats.values,
        hovertemplate='<b>%{label}</b><br>料件數: %{value}<br>比例: %{percent:.1%}<extra></extra>',
        textposition='inside',
        textinfo='label+percent'
    )])
    
    fig.update_layout(
        title='各交貨期區間的料件分佈',
        height=400,
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.01
        )
    )
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("#### 統計摘要")
        st.metric("總料件數", len(delivery_df))
        st.metric("平均交貨期", f"{delivery_df['交貨期(L)'].mean():.1f} 天")
        st.metric("中位交貨期", f"{delivery_df['交貨期(L)'].median():.0f} 天")
        st.metric("最短交貨期", f"{delivery_df['交貨期(L)'].min()} 天")
        st.metric("最長交貨期", f"{delivery_df['交貨期(L)'].max()} 天")
    
    st.markdown("---")

# ==========================================
# 搜尋功能
# ==========================================
st.subheader("搜尋")

# 建立下拉式選單選項：料號 | 品名
df['選項標籤'] = df['料號'].astype(str) + " | " + df['品名'].astype(str).fillna('')
search_options = ["-- 全部 --"] + df['選項標籤'].tolist()

selected_material = st.selectbox(
    "選擇料號或品名",
    options=search_options,
    help="從下拉清單中選擇料件"
)

# ==========================================
# 資料篩選
# ==========================================
df_filtered = df.copy()

if selected_material != "-- 全部 --":
    # 從選項中提取料號 (格式：料號 | 品名)
    selected_code = selected_material.split(" | ")[0]
    df_filtered = df_filtered[df_filtered['料號'].astype(str) == selected_code]

# ==========================================
# 資訊摘要
# ==========================================
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "總料件數",
        len(df),
        help="所有料件總數"
    )

with col2:
    stock_zero = len(df[df['庫存量'] == 0])
    st.metric(
        "缺貨料件",
        stock_zero,
        delta=f"{stock_zero/len(df)*100:.1f}%" if len(df) > 0 else "0%",
        delta_color="inverse"
    )

with col3:
    if selected_material != "-- 全部 --":
        st.metric(
            "搜尋結果",
            len(df_filtered),
            help="符合搜尋條件的料件數"
        )

st.markdown("---")

# ==========================================
# 資料表顯示（含條件格式化）
# ==========================================
st.subheader("料件清單")

if len(df_filtered) > 0:
    # 創建顯示用的 DataFrame（複製以保持原始資料）
    display_df = df_filtered.copy()
    
    # 為了在 Streamlit 中顯示紅色警告，使用 st.dataframe 的內建樣式功能
    def highlight_zero_stock(row):
        """將庫存量為 0 的列著色為紅色"""
        if row['庫存量'] == 0:
            return ['background-color: #ffcccc'] * len(row)
        return [''] * len(row)
    
    # 應用樣式
    styled_df = display_df.style.apply(highlight_zero_stock, axis=1)
    
    # 顯示表格
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=600,
        hide_index=True
    )
    
    # 顯示篩選結果統計
    st.caption(f"顯示 {len(df_filtered)} / {len(df)} 筆料件")
    
    # 下載按鈕
    csv_data = df_filtered.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        label="下載搜尋結果 (CSV)",
        data=csv_data,
        file_name=f"料件清單_搜尋結果.csv",
        mime="text/csv"
    )

else:
    st.info("沒有資料可顯示")

st.markdown("---")

# ==========================================
# 統計資訊
# ==========================================
with st.expander("📊 詳細統計"):
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 按倉庫統計")
        warehouse_stats = df_filtered.groupby('倉庫')['庫存量'].agg(['count', 'sum']).rename(
            columns={'count': '料件數', 'sum': '總庫存'}
        )
        st.dataframe(warehouse_stats, use_container_width=True)
    
    with col2:
        st.markdown("### 按單位統計")
        unit_stats = df_filtered.groupby('單位')['庫存量'].agg(['count', 'sum']).rename(
            columns={'count': '料件數', 'sum': '總庫存'}
        )
        st.dataframe(unit_stats, use_container_width=True)

    # 庫存狀態分析
    st.markdown("### 庫存狀態分析")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        zero_count = len(df_filtered[df_filtered['庫存量'] == 0])
        st.metric("缺貨 (0件)", zero_count)
    
    with col2:
        low_count = len(df_filtered[(df_filtered['庫存量'] > 0) & (df_filtered['庫存量'] <= 10)])
        st.metric("低庫存 (1-10件)", low_count)
    
    with col3:
        normal_count = len(df_filtered[(df_filtered['庫存量'] > 10) & (df_filtered['庫存量'] <= 100)])
        st.metric("正常庫存 (11-100件)", normal_count)
    
    with col4:
        high_count = len(df_filtered[df_filtered['庫存量'] > 100])
        st.metric("充足庫存 (>100件)", high_count)
