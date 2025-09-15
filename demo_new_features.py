#!/usr/bin/env python3
"""
退貨建議分析系統 - 新功能演示腳本
演示計算類型選擇功能
"""

import pandas as pd
import numpy as np
import sys

# 添加當前路徑
sys.path.append('/workspace')

# 導入主應用模塊的核心函數
from app import preprocess_data, generate_return_recommendations

def demo_calculation_types():
    """演示計算類型選擇功能"""
    print("🎯 退貨建議分析系統 - 新功能演示")
    print("=" * 60)
    
    try:
        # 讀取真實數據
        print("📊 載入真實數據...")
        real_data = pd.read_excel('/workspace/user_input_files/ELE_15Sep2025.XLSX', dtype={'Article': str})
        processed_data = preprocess_data(real_data)
        print(f"✅ 成功載入 {processed_data.shape[0]} 行數據")
        
        print("\n" + "=" * 60)
        print("🔍 分析不同計算類型的結果對比")
        print("=" * 60)
        
        # 1. 所有類型分析
        print("\n1️⃣  所有類型分析 (ND + RF)")
        print("-" * 30)
        recommendations_all = generate_return_recommendations(processed_data, "both")
        
        if len(recommendations_all) > 0:
            nd_count = (recommendations_all['Type'] == 'ND').sum()
            rf_count = (recommendations_all['Type'] == 'RF').sum()
            total_qty = recommendations_all['Transfer Qty'].sum()
            
            print(f"📈 退貨建議總數: {len(recommendations_all)} 條")
            print(f"📦 總退貨件數: {total_qty} 件")
            print(f"🔵 ND 類型: {nd_count} 條")
            print(f"🟡 RF 類型: {rf_count} 條")
            
            # 按 OM 統計
            om_stats = recommendations_all.groupby('OM')['Transfer Qty'].sum().sort_values(ascending=False)
            print(f"\n📊 按 OM 退貨件數分布:")
            for om, qty in om_stats.head(5).items():
                print(f"   • {om}: {qty} 件")
        else:
            print("❌ 未生成任何退貨建議")
        
        # 2. 只 ND 類型分析
        print("\n2️⃣  只計算 ND 類型退倉")
        print("-" * 30)
        recommendations_nd = generate_return_recommendations(processed_data, "nd_only")
        
        if len(recommendations_nd) > 0:
            total_qty_nd = recommendations_nd['Transfer Qty'].sum()
            print(f"📈 ND 退倉建議: {len(recommendations_nd)} 條")
            print(f"📦 ND 退倉件數: {total_qty_nd} 件")
            
            # 顯示具體建議
            print(f"\n📋 ND 類型退倉詳情:")
            for _, row in recommendations_nd.iterrows():
                print(f"   • Article {row['Article']} - Site {row['Transfer Site']}: {row['Transfer Qty']} 件")
        else:
            print("ℹ️  暫無 ND 類型退倉建議")
        
        # 3. 只 RF 類型分析
        print("\n3️⃣  只計算 RF 類型過剩退倉")
        print("-" * 30)
        recommendations_rf = generate_return_recommendations(processed_data, "rf_only")
        
        if len(recommendations_rf) > 0:
            total_qty_rf = recommendations_rf['Transfer Qty'].sum()
            print(f"📈 RF 過剩退倉建議: {len(recommendations_rf)} 條")
            print(f"📦 RF 退倉件數: {total_qty_rf} 件")
            
            # 按 OM 統計
            om_stats_rf = recommendations_rf.groupby('OM')['Transfer Qty'].sum().sort_values(ascending=False)
            print(f"\n📊 RF 類型按 OM 分布:")
            for om, qty in om_stats_rf.head(5).items():
                print(f"   • {om}: {qty} 件")
        else:
            print("ℹ️  暫無 RF 類型過剩退倉建議")
        
        # 4. 結果對比分析
        print("\n" + "=" * 60)
        print("📊 計算類型對比分析")
        print("=" * 60)
        
        print(f"╭─────────────────────┬──────────┬──────────┬──────────╮")
        print(f"│ 分析類型            │ 建議條數 │ 退貨件數 │ 百分比   │")
        print(f"├─────────────────────┼──────────┼──────────┼──────────┤")
        
        all_count = len(recommendations_all)
        all_qty = recommendations_all['Transfer Qty'].sum() if all_count > 0 else 0
        nd_count = len(recommendations_nd)
        nd_qty = recommendations_nd['Transfer Qty'].sum() if nd_count > 0 else 0
        rf_count = len(recommendations_rf)
        rf_qty = recommendations_rf['Transfer Qty'].sum() if rf_count > 0 else 0
        
        print(f"│ 所有類型 (ND + RF)  │ {all_count:8d} │ {all_qty:8d} │ {100.0:7.1f}% │")
        nd_pct = (nd_qty / all_qty * 100) if all_qty > 0 else 0
        print(f"│ 只計算 ND 類型      │ {nd_count:8d} │ {nd_qty:8d} │ {nd_pct:7.1f}% │")
        rf_pct = (rf_qty / all_qty * 100) if all_qty > 0 else 0
        print(f"│ 只計算 RF 類型      │ {rf_count:8d} │ {rf_qty:8d} │ {rf_pct:7.1f}% │")
        print(f"╰─────────────────────┴──────────┴──────────┴──────────╯")
        
        # 5. 業務洞察
        print("\n💡 業務洞察與建議:")
        print("-" * 30)
        
        if nd_count > 0:
            nd_avg = nd_qty / nd_count
            print(f"🔵 ND 類型平均退倉量: {nd_avg:.1f} 件/商品")
            print(f"   建議：優先處理 ND 類型商品的完全清倉")
        
        if rf_count > 0:
            rf_avg = rf_qty / rf_count
            print(f"🟡 RF 類型平均退倉量: {rf_avg:.1f} 件/商品")
            print(f"   建議：RF 類型商品可分批處理，注意維持安全庫存")
        
        if all_count > 0:
            print(f"📈 整體效率：通過退貨優化可釋放 {all_qty} 件庫存空間")
            
            # 最活躍的 OM
            if len(recommendations_all) > 0:
                top_om = recommendations_all.groupby('OM')['Transfer Qty'].sum().idxmax()
                top_om_qty = recommendations_all.groupby('OM')['Transfer Qty'].sum().max()
                print(f"🏆 最需要退貨調整的 OM：{top_om} ({top_om_qty} 件)")
        
        print("\n" + "=" * 60)
        print("🎉 演示完成！")
        print("💡 用戶可根據具體業務需求選擇相應的計算類型進行分析")
        print("🌐 現在可以訪問 http://localhost:8501 使用完整的網頁界面")
        print("=" * 60)
        
    except Exception as e:
        print(f"❌ 演示失敗: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    demo_calculation_types()
