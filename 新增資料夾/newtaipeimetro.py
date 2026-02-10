import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import re

# ==========================================
# 頁面設定
# ==========================================
st.set_page_config(
    page_title="新北捷運維修物料智慧庫存預警系統",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 自訂CSS樣式
# ==========================================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 0.5rem 0;
        margin-top: -30px;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #1f77b4;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 參數設定
# ==========================================
W = 365.0      # 年天數
r = 0.3        # 風險係數
CV = 0.25      # 預設變異係數
Z_MAP = {
    1: 1.28,   # 90%
    2: 1.65,   # 95%
    3: 1.96,   # 97.5%
    4: 2.33,   # 99%
    5: 3.09    # 99.9%
}

# 處理不同環境下的檔案路徑
import os
import sys

def get_file_path(filename):
    """取得檔案路徑，支援 PyInstaller 打包的環境"""
    # 如果是 PyInstaller 打包的 exe
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    file_path = os.path.join(base_path, filename)
    return file_path

FILE_NAME = get_file_path('物料整合報表.xlsx') 

# ==========================================
# 資料載入與計算（使用快取）
# ==========================================
@st.cache_data
def load_and_calculate_data():
    """載入資料並計算ROP與採購日期"""
    
    try:
        # 讀取 Excel
        df = pd.read_excel(FILE_NAME, sheet_name='整合資料')
    except Exception as e:
        st.error(f"讀取檔案失敗: {e}")
        st.info("請確認「物料整合報表.xlsx」在相同資料夾中")
        return None, []
    
    # 填補基本空值
    df['交貨期(L)'] = df['交貨期(L)'].fillna(30)
    df['採購前置時間(P)'] = df['採購前置時間(P)'].fillna(30)
    df['南機廠庫存量'] = df['南機廠庫存量'].fillna(0)
    df['安全等級'] = df['安全等級'].fillna(2)
    
    # 自動判斷供應風險
    df['Is_High_Risk'] = np.where(df['交貨期(L)'] > 60, 1, 0)
    df['Effective_T'] = df['交貨期(L)'] * (1 + r * df['Is_High_Risk']) + df['採購前置時間(P)']
    
    # 計算日用量 (d)
    df['Annual_Usage_Base'] = df['2025年使用量(U25)'].fillna(
        df['2024年使用量(U24)'].fillna(
            df['2023年使用量(U23)'].fillna(0)
        )
    )
    df['d'] = df['Annual_Usage_Base'] / W
    df['d'] = df['d'].replace(0, 0.0001)  # 防呆
    
    # 計算 ROP
    df['Z_Value'] = df['安全等級'].map(Z_MAP).fillna(1.65)
    df['Sigma_d'] = df['d'] * CV
    
    df['Calculated_ROP'] = (df['d'] * df['Effective_T']) + \
                           (df['Z_Value'] * df['Sigma_d'] * np.sqrt(df['Effective_T']))
    df['Calculated_ROP'] = np.ceil(df['Calculated_ROP'])
    
    # 預測採購日期
    today = pd.Timestamp.now().normalize()
    df['Surplus_Stock'] = df['南機廠庫存量'] - df['Calculated_ROP']
    df['Days_Until_Reorder'] = df['Surplus_Stock'] / df['d']
    
    # 狀態標記 
    df['Action_Status'] = '安全'
    df.loc[df['Days_Until_Reorder'] <= 0, 'Action_Status'] = '缺貨'
    df.loc[(df['Days_Until_Reorder'] > 0) & (df['Days_Until_Reorder'] <= 30), 'Action_Status'] = '需補貨'
    df.loc[df['Days_Until_Reorder'] > 730, 'Action_Status'] = '庫存充足'
    
    # 計算日期
    clean_days = np.where(df['Days_Until_Reorder'] < 0, 0, df['Days_Until_Reorder'])
    clean_days = np.where(clean_days > 1825, 1825, clean_days)
    df['Estimated_Reorder_Date'] = today + pd.to_timedelta(clean_days, unit='D')
    df['Estimated_Reorder_Date_Str'] = df['Estimated_Reorder_Date'].dt.strftime('%Y-%m-%d')
    
    # 抓取所有月份欄位
    date_cols = [col for col in df.columns if re.match(r'\d{4}年\d{1,2}月', str(col))]
    
    # 計算平均月用量
    if date_cols:
        month_usage = df[date_cols].copy()
        for col in month_usage.columns:
            month_usage[col] = pd.to_numeric(month_usage[col], errors='coerce')
        df['平均月用量'] = month_usage.mean(axis=1).fillna(0)
    else:
        df['平均月用量'] = 0
    
    return df, date_cols

# ==========================================
# 主程式
# ==========================================
st.markdown('<h1 class="main-header">新北捷運維修物料智慧庫存預警系統</h1>', unsafe_allow_html=True)

# 載入資料
with st.spinner("正在載入資料並計算請購點..."):
    df, date_cols = load_and_calculate_data()

if df is None:
    st.stop()

# ==========================================
# 側邊欄篩選
# ==========================================
st.sidebar.header("篩選與搜尋")

# 狀態篩選
status_options = ['全部'] + sorted(df['Action_Status'].unique().tolist())
selected_status = st.sidebar.selectbox("庫存狀態", status_options)

# 搜尋物料編號 
# 建立選單選項： "物料編號 | 品名"
df['Search_Label'] = df['物料編號'].astype(str) + " | " + df['品名'].astype(str).fillna('')
search_options = ["全部"] + df['Search_Label'].tolist()
selected_search_item = st.sidebar.selectbox("搜尋物料 (編號或品名)", search_options)

# 風險等級篩選
risk_options = ['全部', '高風險(交期>60天)', '一般']
selected_risk = st.sidebar.selectbox("供應風險", risk_options)

# 安全等級篩選
safety_options = ['全部', '1', '2', '3', '4']
selected_safety = st.sidebar.selectbox("安全等級", safety_options)

# 套用篩選
df_filtered = df.copy()

if selected_status != '全部':
    df_filtered = df_filtered[df_filtered['Action_Status'] == selected_status]

# 套用新的下拉式選單搜尋邏輯
if selected_search_item != '全部':
    # 從選項中切分出物料編號 (假設格式為 "ID | Name")
    selected_id = selected_search_item.split(" | ")[0]
    df_filtered = df_filtered[df_filtered['物料編號'].astype(str) == selected_id]

if selected_risk == '高風險(交期>60天)':
    df_filtered = df_filtered[df_filtered['Is_High_Risk'] == 1]
elif selected_risk == '一般':
    df_filtered = df_filtered[df_filtered['Is_High_Risk'] == 0]

if selected_safety != '全部':
    df_filtered = df_filtered[df_filtered['安全等級'] == int(selected_safety)]

st.sidebar.markdown("---")
st.sidebar.info(f"篩選結果: {len(df_filtered)} / {len(df)} 項")

# ==========================================
# 新增：庫存狀態視覺化區域（並排顯示）
# ==========================================

# 建立兩欄並排
viz_col1, viz_col2 = st.columns(2)

# 左側：缺貨料件占比分析（圓餅圖）
with viz_col1:
    st.markdown("#### 缺貨料件占比分析")
    
    # 計算缺貨統計
    total_items = len(df)
    # 修改：將缺貨定義從「庫存為0」改為「Action_Status 為 缺貨」
    shortage_items = len(df[df['Action_Status'] == '缺貨'])
    normal_items = total_items - shortage_items
    shortage_percent = (shortage_items / total_items) * 100 if total_items > 0 else 0
    normal_percent = 100 - shortage_percent
    
    # 建立圓餅圖（甜甜圈圖）
    fig_donut = go.Figure()
    
    colors = ['#EF5350', '#66BB6A']  # 紅色=缺貨，綠色=正常
    
    fig_donut.add_trace(go.Pie(
        labels=['缺貨料件', '正常庫存'],
        values=[shortage_items, normal_items],
        hole=0.5,
        marker=dict(colors=colors, line=dict(color='white', width=3)),
        textposition='inside',
        textinfo='percent',
        textfont=dict(size=14, color='white', family='Arial Black'),
        hovertemplate='<b>%{label}</b><br>數量: %{value:,}<br>占比: %{percent}<extra></extra>',
        pull=[0.05, 0]  # 突出缺貨部分
    ))
    
    # 在中心添加文字
    fig_donut.add_annotation(
        text=f'<b>{total_items:,}</b>',
        x=0.5, y=0.55,
        font=dict(size=28, color='#212121', family='Arial Black'),
        showarrow=False,
        xref="paper", yref="paper"
    )
    
    fig_donut.add_annotation(
        text='總料件數',
        x=0.5, y=0.45,
        font=dict(size=12, color='#757575'),
        showarrow=False,
        xref="paper", yref="paper"
    )
    
    fig_donut.update_layout(
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.1,
            xanchor="center",
            x=0.5,
            font=dict(size=11)
        ),
        height=400,
        margin=dict(l=20, r=20, t=30, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
    )
    
    st.plotly_chart(fig_donut, use_container_width=True)
    
    # 狀態評估（緊湊顯示）
    if shortage_percent > 30:
        st.error(f"[高風險] - 缺貨率 {shortage_percent:.1f}%")
    elif shortage_percent > 20:
        st.warning(f"[中風險] - 缺貨率 {shortage_percent:.1f}%")
    else:
        st.success(f"[狀態良好] - 缺貨率 {shortage_percent:.1f}%")

# 右側：交貨期分布統計（直方圖）
with viz_col2:
    st.markdown("#### 交貨期分布統計")
    
    # 讀取原始Excel以獲取真實的空白值
    try:
        df_original = pd.read_excel(FILE_NAME, sheet_name='整合資料')
        # 取得原始交貨期數據（包含真實的空白/NaN）
        delivery_data_original = df_original['交貨期(L)']
    except:
        delivery_data_original = df['交貨期(L)']
    
    # 定義交貨期區間
    bins_info = [
        ('null', 'null', '無資料'),  # 空白值
        (1, 30, '1-30'),
        (31, 60, '31-60'),
        (61, 90, '61-90'),
        (91, 120, '91-120'),
        (121, 150, '121-150'),
        (151, 180, '151-180'),
        (181, 240, '181-240'),
        (241, 330, '241-330'),
        (331, float('inf'), '360+')
    ]
    
    # 統計各區間的數量
    counts = []
    labels = []
    
    for start, end, label in bins_info:
        if start == 'null':
            # 統計原始數據中的空白值（NaN或None）
            count = delivery_data_original.isna().sum()
        elif end == float('inf'):
            count = np.sum(delivery_data_original >= start)
        else:
            count = np.sum((delivery_data_original >= start) & (delivery_data_original <= end))
        
        counts.append(int(count))
        labels.append(label)
    
    # 計算百分比
    total_count = sum(counts)
    percentages = [(c / total_count * 100) if total_count > 0 else 0 for c in counts]
    
    # 建立直方圖
    # 使用顏色突出重點區間
    colors_bar = ['#0D47A1' if c == max(counts) else '#1976D2' for c in counts]
    
    fig_hist = go.Figure()
    
    fig_hist.add_trace(go.Bar(
        x=labels,
        y=counts,
        text=[f'{c}<br>({p:.1f}%)' for c, p in zip(counts, percentages)],
        textposition='outside',
        textfont=dict(size=10, color='#424242', family='Arial Black'),
        marker=dict(
            color=colors_bar,
            line=dict(color='white', width=2)
        ),
        hovertemplate='<b>%{x}</b><br>數量: %{y:,}<extra></extra>',
        showlegend=False
    ))
    
    fig_hist.update_layout(
        xaxis=dict(
            title=dict(text='<b>交貨期區間 (天)</b>', font=dict(size=12, color='#424242')),
            tickfont=dict(size=10, color='#424242'),
            showgrid=False
        ),
        yaxis=dict(
            title=dict(text='<b>物料數量</b>', font=dict(size=12, color='#424242')),
            tickfont=dict(size=10, color='#616161'),
            showgrid=True,
            gridcolor='rgba(0,0,0,0.1)',
            gridwidth=1
        ),
        height=400,
        margin=dict(l=50, r=20, t=30, b=50),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='#FAFAFA',
        hovermode='x unified'
    )
    
    st.plotly_chart(fig_hist, use_container_width=True)
    
    # 簡要統計（緊湊顯示）
    # 計算平均交貨期時排除空白值
    avg_delivery = delivery_data_original.dropna().mean()
    high_risk = len(df[df['交貨期(L)'] > 60])
    no_data_count = delivery_data_original.isna().sum()
    st.info(f"平均交貨期: {avg_delivery:.0f}天 | 高風險物料: {high_risk}項 ({high_risk/len(df)*100:.1f}%) | 無資料: {no_data_count}項")

st.markdown("---")

# ==========================================
# 儀表板摘要
# ==========================================
st.subheader("庫存概況總覽")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "總物料數", 
        f"{len(df):,}",
        help="系統中所有物料品項數量"
    )

with col2:
    out_of_stock = len(df[df['Action_Status'] == '缺貨'])
    st.metric(
        "缺貨", 
        f"{out_of_stock:,}",
        delta=f"{out_of_stock/len(df)*100:.1f}%" if len(df) > 0 else "0%",
        delta_color="inverse"
    )

with col3:
    safe_stock = len(df[df['Action_Status'].str.contains('安全')])
    st.metric(
        "安全", 
        f"{safe_stock:,}",
        delta=f"{safe_stock/len(df)*100:.1f}%" if len(df) > 0 else "0%",
        delta_color="normal"
    )

with col4:
    high_risk_count = len(df[df['Is_High_Risk'] == 1])
    st.metric(
        "高風險", 
        f"{high_risk_count:,}",
        help="交貨期超過60天的物料"
    )

st.markdown("---")

# ==========================================
# 主要資料表 (含線上編輯功能)
# ==========================================
st.subheader("物料庫存清單與盤點")
st.info("提示：您可以直接點擊「當前庫存」欄位修改數值，系統將自動重新計算狀態。")

# 1. 準備編輯用的資料
edit_cols = [
    '物料編號', '品名', '南機廠庫存量', 'Calculated_ROP', 
    'Action_Status', 'Days_Until_Reorder', '交貨期(L)', 'Estimated_Reorder_Date_Str'
]
df_editor_source = df_filtered[edit_cols].copy()

# 新增格式化函數
def format_days_status(x):
    if not pd.notna(x): return "N/A"
    if x <= 0: return "[應補貨]"
    if x >= 365: return "[可用量充足]"
    return f"{x:.0f}"

# 套用格式化
df_editor_source['Days_Until_Reorder'] = df_editor_source['Days_Until_Reorder'].apply(format_days_status)

# 2. 顯示線上編輯器
edited_df = st.data_editor(
    df_editor_source,
    column_config={
        "物料編號": st.column_config.TextColumn("物料編號", disabled=True),
        "品名": st.column_config.TextColumn("品名", disabled=True),
        "南機廠庫存量": st.column_config.NumberColumn(
            "當前庫存 (可編輯)", 
            min_value=0, 
            step=1, 
            format="%d",
            help="點擊此處輸入最新的盤點數量"
        ),
        "Calculated_ROP": st.column_config.NumberColumn("請購點 ROP", disabled=True, format="%d"),
        "Action_Status": st.column_config.TextColumn("當前狀態", disabled=True),
        "Days_Until_Reorder": st.column_config.TextColumn("可用天數", disabled=True),
        "交貨期(L)": st.column_config.NumberColumn("交貨期", disabled=True),
        "Estimated_Reorder_Date_Str": st.column_config.TextColumn("建議日期", disabled=True),
    },
    use_container_width=True,
    height=450,
    hide_index=True,
    key="inventory_editor"
)

# ==========================================
# 3. 接收編輯後的結果並即時運算
# ==========================================
if not edited_df.equals(df_editor_source):
    surplus = edited_df['南機廠庫存量'] - edited_df['Calculated_ROP']
    
    def update_status(row):
        stock = row['南機廠庫存量']
        rop = row['Calculated_ROP']
        if stock <= 0: return '缺貨'
        if stock <= rop: return '需補貨' 
        return '安全'

    edited_df['New_Status'] = edited_df.apply(update_status, axis=1)
    
    changed_rows = edited_df[edited_df['Action_Status'] != edited_df['New_Status']]
    if len(changed_rows) > 0:
        st.success(f"偵測到 {len(changed_rows)} 筆物料狀態已改變！")
        st.dataframe(changed_rows[['物料編號', '品名', '南機廠庫存量', 'Action_Status', 'New_Status']])

# ==========================================
# 4. 下載更新後的檔案
# ==========================================
st.markdown("### 儲存變更")
col_save1, col_save2 = st.columns([2, 5])

with col_save1:
    df_download = df.copy()
    # 確保原始資料沒有重複的物料編號，若有則取第一筆
    df_download = df_download.drop_duplicates(subset=['物料編號'])
    df_download.set_index('物料編號', inplace=True)
    
    edited_df_temp = edited_df.copy()
    # 確保編輯後的資料沒有重複的物料編號
    edited_df_temp = edited_df_temp.drop_duplicates(subset=['物料編號'])
    edited_df_temp.set_index('物料編號', inplace=True)
    
    df_download.update(edited_df_temp[['南機廠庫存量']])
    df_download.reset_index(inplace=True)
    
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_download.to_excel(writer, index=False, sheet_name='整合資料')
        
    st.download_button(
        label="[下載] 更新後的 Excel 報表",
        data=output.getvalue(),
        file_name=f"物料庫存_盤點更新_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col_save2:
    st.caption("說明：網頁上的修改僅保留於當前畫面。若要永久保存盤點結果，請點擊下載按鈕。")

st.markdown("---")

# ==========================================
# 用量趨勢圖表區
# ==========================================
st.subheader("物料用量趨勢分析")

if len(df_filtered) > 0 and len(date_cols) > 0:
    # 修改：建立「物料編號 | 品名」的選項清單
    df_filtered['Trend_Search_Label'] = df_filtered['物料編號'].astype(str) + " | " + df_filtered['品名'].astype(str).fillna('')
    material_options = df_filtered['Trend_Search_Label'].tolist()
    
    col_select1, col_select2 = st.columns([2, 3])
    
    with col_select1:
        selected_label = st.selectbox(
            "選擇物料",
            options=material_options,
            help="選擇要查看用量趨勢圖的物料"
        )
    
    # 從選中的標籤中解析出物料編號
    selected_material = selected_label.split(" | ")[0] if selected_label else None
    
    with col_select2:
        if selected_material:
            material_name = df[df['物料編號'] == selected_material].iloc[0]['品名']
            st.info(f"**{material_name[:50]}{'...' if len(material_name) > 50 else ''}**")
    
    if selected_material:
        material_data = df[df['物料編號'] == selected_material].iloc[0]
        
        col_info1, col_info2, col_info3, col_info4 = st.columns(4)
        
        with col_info1:
            st.metric("當前庫存", f"{material_data['南機廠庫存量']:.0f}")
        with col_info2:
            st.metric("請購點ROP", f"{material_data['Calculated_ROP']:.0f}")
        with col_info3:
            st.metric("平均月用量", f"{material_data['平均月用量']:.1f}")
        with col_info4:
            status = material_data['Action_Status']
            if '缺貨' in status:
                st.error(f"{status}")
            elif '補貨' in status:
                st.warning(f"{status}")
            else:
                st.success(f"{status}")
        
        usage_values = []
        for col in date_cols:
            val = material_data[col]
            usage_values.append(float(val) if pd.notna(val) else 0)
        
        avg_usage = np.mean([v for v in usage_values if v > 0]) if any(v > 0 for v in usage_values) else 0
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=date_cols,
            y=usage_values,
            mode='lines+markers',
            name='月用量',
            line=dict(color='#1f77b4', width=3),
            marker=dict(size=8, color='#1f77b4'),
            hovertemplate='<b>%{x}</b><br>用量: %{y:.1f}<extra></extra>',
            fill='tozeroy',
            fillcolor='rgba(31, 119, 180, 0.1)'
        ))
        
        if avg_usage > 0:
            fig.add_hline(
                y=avg_usage,
                line_dash="dash",
                line_color="orange",
                line_width=2,
                annotation_text=f"平均月用量: {avg_usage:.1f}",
                annotation_position="right"
            )
        
        max_usage = max(usage_values) if usage_values else 0
        if max_usage > 0:
            fig.add_hline(
                y=max_usage,
                line_dash="dot",
                line_color="red",
                line_width=1,
                annotation_text=f"最大值: {max_usage:.0f}",
                annotation_position="left"
            )
        
        fig.update_layout(
            title=dict(
                text=f"<b>{selected_material}</b> 用量趨勢",
                font=dict(size=18)
            ),
            xaxis_title="月份",
            yaxis_title="消耗量",
            hovermode='x unified',
            height=500,
            showlegend=True,
            xaxis_tickangle=-45,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
        )
        
        fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
        
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("##### 統計摘要")
        col_stat1, col_stat2, col_stat3, col_stat4, col_stat5 = st.columns(5)
        
        non_zero_values = [v for v in usage_values if v > 0]
        
        with col_stat1:
            st.metric("最大用量", f"{max(usage_values):.0f}")
        with col_stat2:
            st.metric("最小用量", f"{min(non_zero_values):.0f}" if non_zero_values else "0")
        with col_stat3:
            st.metric("平均用量", f"{avg_usage:.1f}")
        with col_stat4:
            total_usage = sum(usage_values)
            st.metric("總用量", f"{total_usage:.0f}")
        with col_stat5:
            std_usage = np.std(usage_values)
            st.metric("標準差", f"{std_usage:.1f}")

else:
    if len(df_filtered) == 0:
        st.warning("沒有符合篩選條件的物料資料")
    else:
        st.info("資料中沒有月份用量資訊")

st.markdown("---")

# ==========================================
# 採購優先順序排行
# ==========================================
st.subheader("採購優先順序排行榜")

tab1, tab2 = st.tabs(["[緊急] 最緊急缺口分析 (TOP 10)", "[分布] 狀態分布"])

with tab1:
    df_urgent = df[df['Days_Until_Reorder'] <= 0].sort_values('Days_Until_Reorder').head(10).copy()
    
    if len(df_urgent) > 0:
        st.caption("缺貨物料清單 (TOP 10) - 按緊急程度排序")
        
        for idx, (item_idx, row) in enumerate(df_urgent.iterrows()):
            with st.container():
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.markdown(f"**{row['物料編號']}** - {row['品名'][:40]}{'...' if len(row['品名']) > 40 else ''}")
                
                with col2:
                    st.markdown('<div style="background-color: #ff4444; color: white; padding: 8px; border-radius: 5px; text-align: center; font-weight: bold;">缺貨</div>', unsafe_allow_html=True)
                
                current_stock = row['南機廠庫存量']
                rop = row['Calculated_ROP']
                progress_ratio = 0.0
                
                st.markdown(f"**庫存水位**: {current_stock:.0f} / {rop:.0f} 單位")
                st.progress(progress_ratio)
            
            if idx < len(df_urgent) - 1:
                st.markdown("---")
        
        st.markdown("詳細數據表")
        priority_display = df_urgent[[
            '物料編號', '品名', 'Action_Status', '南機廠庫存量', 
            'Calculated_ROP', 'Days_Until_Reorder', 'Estimated_Reorder_Date_Str'
        ]].copy()
        
        priority_display.columns = ['物料編號', '品名', '狀態', '當前庫存', 'ROP', '可用天數(估)', '建議請購日']
        
        priority_display['當前庫存'] = priority_display['當前庫存'].astype(float)
        priority_display['ROP'] = priority_display['ROP'].astype(float)
        priority_display['可用天數(估)'] = priority_display['可用天數(估)'].astype(float)
        
        def format_days(x):
            if x < 0: return "[0天]"
            if x <= 30: return f"[{x:.0f}天]"
            return f"{x:.0f}天"
            
        priority_display['可用天數(估)'] = priority_display['可用天數(估)'].apply(format_days)
        
        st.dataframe(priority_display, use_container_width=True, hide_index=True)
    else:
        st.success("[良好] 目前沒有需要緊急採購的物料！所有物料庫存充足。")

with tab2:
    status_counts = df['Action_Status'].value_counts()
    
    fig_pie = px.pie(
        values=status_counts.values,
        names=status_counts.index,
        title="物料狀態分布",
        color=status_counts.index,
        color_discrete_map={
            '缺貨': '#ff4444',
            '需補貨': '#ffaa00',
            '安全': '#44ff44',
            '庫存充足': '#00cc00'
        },
        hole=0.4
    )
    
    fig_pie.update_traces(textposition='inside', textinfo='percent+label+value')
    fig_pie.update_layout(height=450)
    
    st.plotly_chart(
        fig_pie,
        use_container_width=True,
        key="material_status_pie"
    )

# ==========================================
# 頁尾資訊
# ==========================================
st.markdown("---")
st.caption(f"""
資料更新時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
分析物料總數：{len(df):,} 項  
目前顯示：{len(df_filtered):,} 項  
計算參數：CV={CV} | 年天數={W:.0f} | 風險係數={r}  
資料來源：{FILE_NAME}
""")

# 側邊欄說明
with st.sidebar:
    st.markdown("---")
    st.markdown("### 使用說明")
    st.markdown("""
    **狀態說明：**
    - 缺貨：庫存為0
    - 需補貨：可用天數≤30天
    - 安全：庫存正常
    - 庫存充足：可用天數>2年
    

    """)
    
    st.markdown("---")