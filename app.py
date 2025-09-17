import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import warnings
warnings.filterwarnings('ignore')

# 設定頁面配置
st.set_page_config(
    page_title="退貨建議分析系統",
    page_icon="📦",
    layout="wide"
)

def convert_to_string_format(value):
    """確保 Article 欄位為 12 位字符串格式"""
    if pd.isna(value):
        return ""
    value_str = str(value).strip()
    # 移除小數點（如果是浮點數）
    if '.' in value_str:
        value_str = value_str.split('.')[0]
    # 確保是 12 位數字
    if value_str.isdigit() and len(value_str) <= 12:
        return value_str.zfill(12)
    return value_str

def safe_convert_to_int(value):
    """安全轉換為整數，異常值設為 0"""
    try:
        if pd.isna(value):
            return 0
        return max(0, int(float(value)))
    except:
        return 0

def safe_convert_to_string(value):
    """安全轉換為字符串"""
    if pd.isna(value):
        return ""
    return str(value).strip()

def preprocess_data(df):
    """數據預處理與驗證"""
    df_processed = df.copy()
    
    # 確保 Article 欄位為 12 位字符串格式
    df_processed['Article'] = df_processed['Article'].apply(convert_to_string_format)
    
    # 字符串欄位處理
    string_columns = ['OM', 'RP Type', 'Site']
    for col in string_columns:
        if col in df_processed.columns:
            df_processed[col] = df_processed[col].apply(safe_convert_to_string)
    
    # 整數欄位處理
    int_columns = ['SaSa Net Stock', 'Pending Received', 'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty']
    for col in int_columns:
        if col in df_processed.columns:
            df_processed[col] = df_processed[col].apply(safe_convert_to_int)
    
    # 銷量異常值校正
    df_processed['Notes'] = ""
    
    for idx, row in df_processed.iterrows():
        notes = []
        
        # 檢查銷量數據範圍
        for col in ['Last Month Sold Qty', 'MTD Sold Qty']:
            if col in df_processed.columns:
                if row[col] > 100000:
                    df_processed.loc[idx, col] = 100000
                    notes.append(f'{col}銷量數據超出範圍')
        
        df_processed.loc[idx, 'Notes'] = '; '.join(notes)
    
    return df_processed

def calculate_effective_sold_qty(row):
    """計算有效銷量"""
    last_month = row.get('Last Month Sold Qty', 0)
    mtd = row.get('MTD Sold Qty', 0)
    
    # 優先使用 Last Month Sold Qty，若為 0 則使用 MTD Sold Qty
    if last_month > 0:
        return last_month
    else:
        return mtd

def get_top20_percent_threshold(df, article):
    """計算該 Article 的銷量前 20% 門檻"""
    article_data = df[df['Article'] == article]
    if article_data.empty:
        return float('inf')
    
    sold_quantities = []
    for _, row in article_data.iterrows():
        effective_qty = calculate_effective_sold_qty(row)
        sold_quantities.append(effective_qty)
    
    if not sold_quantities:
        return float('inf')
    
    # 計算 80% 分位數（前 20% 的門檻）
    threshold = np.percentile(sold_quantities, 80)
    return threshold

def generate_return_recommendations(df, calculation_type="both"):
    """生成退貨建議
    
    Args:
        df: 數據框架
        calculation_type: 計算類型 ('nd_only', 'rf_only', 'both')
    """
    recommendations = []
    
    for _, row in df.iterrows():
        article = row['Article']
        om = row['OM']
        rp_type = row['RP Type']
        site = row['Site']
        net_stock = row['SaSa Net Stock']
        pending_received = row['Pending Received']
        safety_stock = row['Safety Stock']
        
        effective_sold_qty = calculate_effective_sold_qty(row)
        
        # 優先級 1: ND 類型退倉
        if rp_type == "ND" and net_stock > 0 and calculation_type in ["nd_only", "both"]:
            recommendations.append({
                'Article': article,
                'Product Desc': row.get('Article Description', ''),
                'OM': om,
                'Transfer Site': site,
                'Receive Site': 'D001',
                'Transfer Qty': net_stock,
                'Notes': f'ND類型退倉 - 優先級1' + (f'; {row.get("Notes", "")}' if row.get("Notes") else ""),
                'Priority': 1,
                'Type': 'ND'
            })
        
        # 優先級 2: RF 類型過剩退倉
        elif rp_type == "RF" and calculation_type in ["rf_only", "both"]:
            total_available = net_stock + pending_received
            
            # 檢查條件：庫存充足
            if total_available > safety_stock:
                # 檢查銷量是否非最高 20%
                top20_threshold = get_top20_percent_threshold(df, article)
                
                if effective_sold_qty < top20_threshold:
                    # 計算可轉數量
                    potential_transfer = total_available - safety_stock
                    
                    # 限制條件：退貨後必須高於安全庫存的 20%
                    min_remaining = max(int(safety_stock * 0.2), 0)
                    max_transfer = total_available - min_remaining
                    
                    # 最終轉移數量（至少 2 件）
                    transfer_qty = min(potential_transfer, max_transfer)
                    
                    if transfer_qty >= 2 and transfer_qty <= net_stock:
                        notes_parts = [f'RF類型過剩退倉 - 優先級2']
                        if row.get("Notes"):
                            notes_parts.append(row.get("Notes"))
                        
                        recommendations.append({
                            'Article': article,
                            'Product Desc': row.get('Article Description', ''),
                            'OM': om,
                            'Transfer Site': site,
                            'Receive Site': 'D001',
                            'Transfer Qty': transfer_qty,
                            'Notes': '; '.join(notes_parts),
                            'Priority': 2,
                            'Type': 'RF'
                        })
    
    return pd.DataFrame(recommendations)

def create_excel_report(recommendations_df, df_original, calculation_type="both"):
    """創建 Excel 報告
    
    Args:
        recommendations_df: 退貨建議數據框架
        df_original: 原始數據框架
        calculation_type: 計算類型 ('nd_only', 'rf_only', 'both')
    """
    # 創建工作簿
    wb = Workbook()
    
    # 定義樣式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'), 
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # 工作表 1: 退貨建議
    ws1 = wb.active
    ws1.title = "退貨建議"
    
    # 寫入標題行
    headers = ['Article', 'Product Desc', 'OM', 'Transfer Site', 'Receive Site', 'Transfer Qty', 'Notes']
    for col_num, header in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
    # 寫入數據
    for row_num, (_, row) in enumerate(recommendations_df.iterrows(), 2):
        for col_num, header in enumerate(headers, 1):
            cell = ws1.cell(row=row_num, column=col_num, value=row[header])
            cell.border = border
    
    # 調整列寬
    column_widths = [15, 30, 10, 15, 15, 12, 40]
    for col_num, width in enumerate(column_widths, 1):
        ws1.column_dimensions[ws1.cell(row=1, column=col_num).column_letter].width = width
    
    # 工作表 2: 統計摘要
    ws2 = wb.create_sheet("統計摘要")
    
    # KPI 橫幅
    total_recommendations = len(recommendations_df)
    total_transfer_qty = recommendations_df['Transfer Qty'].sum() if not recommendations_df.empty else 0
    
    # 分析類型說明
    type_descriptions = {
        "nd_only": "ND 類型退倉分析",
        "rf_only": "RF 類型過剩退倉分析",
        "both": "綜合退貨分析 (ND + RF)"
    }
    analysis_type_desc = type_descriptions.get(calculation_type, "綜合分析")
    
    ws2.cell(row=1, column=1, value="KPI 摘要").font = Font(size=16, bold=True)
    ws2.cell(row=2, column=1, value=f"分析類型: {analysis_type_desc}").font = Font(size=12, bold=True)
    ws2.cell(row=4, column=1, value="總退貨建議數量（條數）:").font = Font(bold=True)
    ws2.cell(row=4, column=2, value=total_recommendations)
    ws2.cell(row=5, column=1, value="總退貨件數:").font = Font(bold=True)
    ws2.cell(row=5, column=2, value=total_transfer_qty)
    
    # 詳細統計表
    current_row = 8
    
    if not recommendations_df.empty:
        # 按 Article 統計
        ws2.cell(row=current_row, column=1, value="按 Article 統計").font = Font(size=14, bold=True)
        current_row += 2
        
        ws2.cell(row=current_row, column=1, value="Article").font = header_font
        ws2.cell(row=current_row, column=1).fill = header_fill
        ws2.cell(row=current_row, column=2, value="總退貨件數").font = header_font
        ws2.cell(row=current_row, column=2).fill = header_fill
        ws2.cell(row=current_row, column=3, value="涉及OM數量").font = header_font
        ws2.cell(row=current_row, column=3).fill = header_fill
        current_row += 1
        
        article_stats = recommendations_df.groupby('Article').agg({
            'Transfer Qty': 'sum',
            'OM': 'nunique'
        }).reset_index()
        
        for _, row in article_stats.iterrows():
            ws2.cell(row=current_row, column=1, value=row['Article'])
            ws2.cell(row=current_row, column=2, value=row['Transfer Qty'])
            ws2.cell(row=current_row, column=3, value=row['OM'])
            current_row += 1
        
        current_row += 2
        
        # 按 OM 統計
        ws2.cell(row=current_row, column=1, value="按 OM 統計").font = Font(size=14, bold=True)
        current_row += 2
        
        ws2.cell(row=current_row, column=1, value="OM").font = header_font
        ws2.cell(row=current_row, column=1).fill = header_fill
        ws2.cell(row=current_row, column=2, value="總退貨件數").font = header_font
        ws2.cell(row=current_row, column=2).fill = header_fill
        ws2.cell(row=current_row, column=3, value="涉及Article數量").font = header_font
        ws2.cell(row=current_row, column=3).fill = header_fill
        current_row += 1
        
        om_stats = recommendations_df.groupby('OM').agg({
            'Transfer Qty': 'sum',
            'Article': 'nunique'
        }).reset_index()
        
        for _, row in om_stats.iterrows():
            ws2.cell(row=current_row, column=1, value=row['OM'])
            ws2.cell(row=current_row, column=2, value=row['Transfer Qty'])
            ws2.cell(row=current_row, column=3, value=row['Article'])
            current_row += 1
        
        current_row += 2
        
        # 轉出類型分布
        ws2.cell(row=current_row, column=1, value="轉出類型分布").font = Font(size=14, bold=True)
        current_row += 2
        
        type_stats = recommendations_df.groupby('Type').agg({
            'Transfer Qty': ['count', 'sum']
        }).round(2)
        type_stats.columns = ['建議數量', '總件數']
        type_stats = type_stats.reset_index()
        
        ws2.cell(row=current_row, column=1, value="類型").font = header_font
        ws2.cell(row=current_row, column=1).fill = header_fill
        ws2.cell(row=current_row, column=2, value="建議數量").font = header_font
        ws2.cell(row=current_row, column=2).fill = header_fill
        ws2.cell(row=current_row, column=3, value="總件數").font = header_font
        ws2.cell(row=current_row, column=3).fill = header_fill
        current_row += 1
        
        for _, row in type_stats.iterrows():
            ws2.cell(row=current_row, column=1, value=row['Type'])
            ws2.cell(row=current_row, column=2, value=row['建議數量'])
            ws2.cell(row=current_row, column=3, value=row['總件數'])
            current_row += 1
        
        current_row += 2
        
        # 優先級分布
        ws2.cell(row=current_row, column=1, value="優先級分布").font = Font(size=14, bold=True)
        current_row += 2
        
        priority_stats = recommendations_df.groupby('Priority').agg({
            'Transfer Qty': ['count', 'sum']
        }).round(2)
        priority_stats.columns = ['建議數量', '總件數']
        priority_stats = priority_stats.reset_index()
        
        ws2.cell(row=current_row, column=1, value="優先級").font = header_font
        ws2.cell(row=current_row, column=1).fill = header_fill
        ws2.cell(row=current_row, column=2, value="建議數量").font = header_font
        ws2.cell(row=current_row, column=2).fill = header_fill
        ws2.cell(row=current_row, column=3, value="總件數").font = header_font
        ws2.cell(row=current_row, column=3).fill = header_fill
        current_row += 1
        
        for _, row in priority_stats.iterrows():
            ws2.cell(row=current_row, column=1, value=f"優先級 {row['Priority']}")
            ws2.cell(row=current_row, column=2, value=row['建議數量'])
            ws2.cell(row=current_row, column=3, value=row['總件數'])
            current_row += 1
    
    # 調整列寬
    for col in ['A', 'B', 'C']:
        ws2.column_dimensions[col].width = 20
    
    return wb

def quality_check(recommendations_df, original_df):
    """質量檢查"""
    checks = []
    
    if recommendations_df.empty:
        checks.append("✅ 無退貨建議生成")
        return checks
    
    # 檢查 1: Article 和 OM 一致性
    for _, row in recommendations_df.iterrows():
        original_row = original_df[
            (original_df['Article'] == row['Article']) & 
            (original_df['Site'] == row['Transfer Site'])
        ]
        if not original_row.empty and original_row.iloc[0]['OM'] == row['OM']:
            continue
        else:
            checks.append(f"❌ Article {row['Article']} 和 OM {row['OM']} 不一致")
            break
    else:
        checks.append("✅ Article 和 OM 一致性檢查通過")
    
    # 檢查 2: Transfer Qty 為正整數
    if all(recommendations_df['Transfer Qty'] > 0):
        checks.append("✅ 所有 Transfer Qty 為正整數")
    else:
        checks.append("❌ 存在非正整數的 Transfer Qty")
    
    # 檢查 3: Transfer Qty 不超過原庫存
    exceeded = False
    for _, row in recommendations_df.iterrows():
        original_row = original_df[
            (original_df['Article'] == row['Article']) & 
            (original_df['Site'] == row['Transfer Site'])
        ]
        if not original_row.empty:
            original_stock = original_row.iloc[0]['SaSa Net Stock']
            if row['Transfer Qty'] > original_stock:
                exceeded = True
                break
    
    if not exceeded:
        checks.append("✅ Transfer Qty 不超過原庫存")
    else:
        checks.append("❌ 存在 Transfer Qty 超過原庫存的情況")
    
    # 檢查 4: Article 格式檢查
    if all(len(str(art)) <= 12 for art in recommendations_df['Article']):
        checks.append("✅ Article 格式正確")
    else:
        checks.append("❌ Article 格式異常")
    
    return checks

def main():
    st.title("📦 退貨建議分析系統")
    st.markdown("---")
    
    # 側邊欄
    st.sidebar.header("🔧 系統設置")
    st.sidebar.markdown("**接收站點**: D001")
    
    # 文件上傳
    st.header("📤 數據上傳")
    
    uploaded_file = st.file_uploader(
        "選擇 Excel 文件",
        type=['xlsx'],
        help="支持 .xlsx 格式的 Excel 文件"
    )
    
    # 處理上傳的文件
    current_file = None
    file_source = ""
    
    if uploaded_file is not None:
        try:
            current_file = pd.read_excel(uploaded_file, dtype={'Article': str})
            file_source = f"上傳文件 ({uploaded_file.name})"
            st.success(f"✅ 文件上傳成功: {uploaded_file.name}")
        except Exception as e:
            st.error(f"❌ 文件讀取失敗: {str(e)}")
    
    if current_file is not None:
        # 數據預覽
        st.header("🔍 數據預覽")
        st.markdown(f"**數據來源**: {file_source}")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("總記錄數", current_file.shape[0])
        with col2:
            st.metric("欄位數", current_file.shape[1])
        with col3:
            st.metric("RP Type 分布", f"ND: {(current_file['RP Type'] == 'ND').sum()}, RF: {(current_file['RP Type'] == 'RF').sum()}")
        
        # 顯示數據表預覽
        st.subheader("數據表預覽 (前5行)")
        key_columns = ['Article', 'Article Description', 'OM', 'RP Type', 'Site', 'SaSa Net Stock', 'Pending Received', 'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty']
        available_columns = [col for col in key_columns if col in current_file.columns]
        
        if available_columns:
            st.dataframe(current_file[available_columns].head(), use_container_width=True)
        else:
            st.warning("⚠️ 未找到關鍵欄位，顯示所有欄位前5行")
            st.dataframe(current_file.head(), use_container_width=True)
        
        # 計算類型選擇
        st.header("⚙️ 分析設置")
        
        calculation_type = st.radio(
            "選擇計算類型 (Select Calculation Type)",
            options=[
                ("both", "ND 和 RF 都計算 (Calculate Both)"),
                ("nd_only", "只計算 ND 類型 (ND Only)"), 
                ("rf_only", "只計算 RF 類型 (RF Only)")
            ],
            format_func=lambda x: x[1],
            index=0,
            help="選擇要進行分析的退貨類型"
        )
        
        selected_type = calculation_type[0]  # 獲取選中的值
        
        st.markdown("---")
        
        if st.button("🚀 生成退貨建議", type="primary", help="點擊開始分析並生成退貨建議"):
            with st.spinner("正在處理數據..."):
                # 數據預處理
                processed_df = preprocess_data(current_file)
                
                # 生成退貨建議
                recommendations_df = generate_return_recommendations(processed_df, selected_type)
                
                # 顯示結果
                st.success("✅ 分析完成！")
                
                # 基本統計
                st.header("📊 分析結果")
                
                if not recommendations_df.empty:
                    # 基本統計說明
                    type_description = {
                        "nd_only": "ND 類型退倉",
                        "rf_only": "RF 類型過剩退倉",
                        "both": "綜合退貨分析"
                    }
                    st.markdown(f"**分析類型**: {type_description[selected_type]}")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("退貨建議總數", len(recommendations_df))
                    with col2:
                        st.metric("總退貨件數", recommendations_df['Transfer Qty'].sum())
                    with col3:
                        nd_count = (recommendations_df['Type'] == 'ND').sum()
                        st.metric("ND 類型", nd_count)
                    with col4:
                        rf_count = (recommendations_df['Type'] == 'RF').sum()
                        st.metric("RF 類型", rf_count)
                    
                    # 顯示退貨建議表
                    st.subheader("🔄 退貨建議表")
                    display_columns = ['Article', 'Product Desc', 'OM', 'Transfer Site', 'Receive Site', 'Transfer Qty', 'Notes']
                    st.dataframe(recommendations_df[display_columns], use_container_width=True)
                    
                    # 統計圖表
                    st.subheader("📈 統計圖表")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # OM 分布
                        om_stats = recommendations_df.groupby('OM')['Transfer Qty'].sum().reset_index()
                        st.bar_chart(om_stats.set_index('OM'))
                        st.caption("各 OM 退貨件數分布")
                    
                    with col2:
                        # 類型分布
                        type_stats = recommendations_df.groupby('Type')['Transfer Qty'].sum().reset_index()
                        st.bar_chart(type_stats.set_index('Type'))
                        st.caption("退貨類型分布")
                    
                else:
                    st.info("📝 未生成任何退貨建議")
                    st.markdown("**可能原因:**")
                    st.markdown("- 所有商品均未達到退貨條件")
                    st.markdown("- ND 類型商品庫存為 0")
                    st.markdown("- RF 類型商品不滿足過剩條件或屬於高銷量商品")
                
                # 質量檢查
                st.header("✅ 質量檢查")
                quality_results = quality_check(recommendations_df, processed_df)
                
                for check in quality_results:
                    if "✅" in check:
                        st.success(check)
                    elif "❌" in check:
                        st.error(check)
                    else:
                        st.info(check)
                
                # 生成並提供下載
                if not recommendations_df.empty:
                    st.header("💾 下載報告")
                    
                    # 創建 Excel 文件
                    wb = create_excel_report(recommendations_df, processed_df, selected_type)
                    
                    # 生成文件名
                    current_date = datetime.now().strftime("%Y%m%d")
                    filename = f"退貨建議_{current_date}.xlsx"
                    
                    # 保存到內存
                    buffer = io.BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)
                    
                    # 提供下載按鈕
                    st.download_button(
                        label="📥 下載退貨建議報告",
                        data=buffer.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help=f"下載包含退貨建議和統計摘要的 Excel 文件"
                    )
                    
                    st.success(f"✅ 報告已準備完成: {filename}")
    else:
        st.info("👆 請上傳 Excel 文件或點擊使用默認文件開始分析")
        
        # 顯示使用說明
        st.header("📋 使用說明")
        
        with st.expander("💡 系統功能", expanded=True):
            st.markdown("""
            **主要功能:**
            - 📤 支持 Excel 文件上傳或使用默認文件
            - 🔍 數據預處理與驗證
            - ⚙️ 自動生成退貨建議（支持 ND 和 RF 類型）
            - 📊 統計分析與圖表展示
            - ✅ 質量檢查與驗證
            - 💾 Excel 報告下載
            """)
        
        with st.expander("🔧 退貨規則說明"):
            st.markdown("""
            **優先級 1 - ND 類型退倉:**
            - 條件：RP Type = "ND"
            - 退貨數量：全部現有庫存
            
            **優先級 2 - RF 類型過剩退倉:**
            - 條件：RP Type = "RF"
            - 庫存充足：現有庫存 + 在途訂單 > 安全庫存
            - 銷量限制：不屬於該商品的前 20% 高銷量店鋪
            - 最少退貨：2 件
            - 安全限制：退貨後庫存需高於安全庫存的 20%
            """)
        
        with st.expander("📋 必需欄位"):
            st.markdown("""
            **Excel 文件必須包含以下欄位:**
            - Article (產品編號)
            - Article Description (產品描述)
            - OM (營運管理單位)
            - RP Type (轉出類型: ND/RF)
            - Site (店鋪編號)
            - SaSa Net Stock (現有庫存)
            - Pending Received (在途訂單)
            - Safety Stock (安全庫存)
            - Last Month Sold Qty (上月銷量)
            - MTD Sold Qty (本月至今銷量)
            """)
    
    # 底部水印
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #888; font-size: 14px; margin-top: 30px;'>
            由 Ricky 開發 | © 2025
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
