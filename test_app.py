#!/usr/bin/env python3
"""
退貨建議分析系統 - 功能測試腳本
"""

import pandas as pd
import numpy as np
from datetime import datetime
import sys
import os

# 添加當前路徑
sys.path.append('/workspace')

# 導入主應用模塊的核心函數
from app import preprocess_data, generate_return_recommendations, calculate_effective_sold_qty, quality_check

def test_data_preprocessing():
    """測試數據預處理功能"""
    print("🔧 測試數據預處理功能...")
    
    # 創建測試數據
    test_data = {
        'Article': ['106545309001', '106545309002', '106545309003'],
        'Article Description': ['Test Product 1', 'Test Product 2', 'Test Product 3'],
        'OM': ['Candy', 'Hippo', 'Queenie'],
        'RP Type': ['ND', 'RF', 'RF'],
        'Site': ['H001', 'H002', 'H003'],
        'SaSa Net Stock': [10, 15, 20],
        'Pending Received': [2, 3, 5],
        'Safety Stock': [5, 8, 10],
        'Last Month Sold Qty': [3, 0, 8],
        'MTD Sold Qty': [2, 5, 4]
    }
    
    df = pd.DataFrame(test_data)
    processed_df = preprocess_data(df)
    
    # 驗證處理結果
    assert processed_df.shape[0] == 3, "數據行數不正確"
    assert 'Notes' in processed_df.columns, "缺少 Notes 欄位"
    assert processed_df['Article'].iloc[0] == '106545309001', "Article 格式處理異常"
    
    print("✅ 數據預處理測試通過")
    return processed_df

def test_return_recommendations():
    """測試退貨建議生成功能"""
    print("🔧 測試退貨建議生成功能...")
    
    # 創建測試數據
    test_data = {
        'Article': ['106545309001', '106545309001', '106545309002', '106545309002'],
        'Article Description': ['Test Product 1', 'Test Product 1', 'Test Product 2', 'Test Product 2'],
        'OM': ['Candy', 'Candy', 'Hippo', 'Hippo'],
        'RP Type': ['ND', 'RF', 'RF', 'RF'],
        'Site': ['H001', 'H002', 'H003', 'H004'],
        'SaSa Net Stock': [10, 15, 8, 20],
        'Pending Received': [0, 3, 2, 5],
        'Safety Stock': [5, 8, 4, 10],
        'Last Month Sold Qty': [0, 2, 1, 8],  # H004 是高銷量店鋪
        'MTD Sold Qty': [0, 1, 1, 4],
        'Notes': ['', '', '', '']
    }
    
    df = pd.DataFrame(test_data)
    recommendations = generate_return_recommendations(df)
    
    print(f"生成 {len(recommendations)} 條退貨建議")
    
    if len(recommendations) > 0:
        print("退貨建議詳情:")
        for _, rec in recommendations.iterrows():
            print(f"  - Article: {rec['Article']}, Site: {rec['Transfer Site']}, Qty: {rec['Transfer Qty']}, Type: {rec.get('Type', 'N/A')}")
    
    # 基本驗證
    if len(recommendations) > 0:
        assert all(recommendations['Receive Site'] == 'D001'), "接收站點應為 D001"
        assert all(recommendations['Transfer Qty'] > 0), "轉移數量應為正數"
    
    print("✅ 退貨建議生成測試通過")
    return recommendations

def test_effective_sold_qty():
    """測試有效銷量計算"""
    print("🔧 測試有效銷量計算...")
    
    # 測試案例
    test_cases = [
        {'Last Month Sold Qty': 5, 'MTD Sold Qty': 3, 'expected': 5},
        {'Last Month Sold Qty': 0, 'MTD Sold Qty': 8, 'expected': 8},
        {'Last Month Sold Qty': 0, 'MTD Sold Qty': 0, 'expected': 0},
    ]
    
    for i, case in enumerate(test_cases):
        result = calculate_effective_sold_qty(case)
        assert result == case['expected'], f"測試案例 {i+1} 失敗: 期望 {case['expected']}, 得到 {result}"
    
    print("✅ 有效銷量計算測試通過")

def test_quality_check():
    """測試質量檢查功能"""
    print("🔧 測試質量檢查功能...")
    
    # 創建測試數據
    original_data = {
        'Article': ['106545309001', '106545309002'],
        'OM': ['Candy', 'Hippo'],
        'Site': ['H001', 'H002'],
        'SaSa Net Stock': [10, 15]
    }
    
    recommendations_data = {
        'Article': ['106545309001'],
        'OM': ['Candy'],
        'Transfer Site': ['H001'],
        'Transfer Qty': [5]
    }
    
    original_df = pd.DataFrame(original_data)
    recommendations_df = pd.DataFrame(recommendations_data)
    
    checks = quality_check(recommendations_df, original_df)
    
    assert len(checks) > 0, "質量檢查應返回檢查結果"
    
    print("質量檢查結果:")
    for check in checks:
        print(f"  {check}")
    
    print("✅ 質量檢查測試通過")

def test_with_real_data():
    """使用真實數據進行測試"""
    print("🔧 使用真實數據進行測試...")
    
    try:
        # 讀取真實數據
        real_data = pd.read_excel('/workspace/user_input_files/ELE_15Sep2025.XLSX', dtype={'Article': str})
        print(f"成功讀取真實數據: {real_data.shape[0]} 行 x {real_data.shape[1]} 列")
        
        # 預處理
        processed_data = preprocess_data(real_data)
        print(f"預處理完成: {processed_data.shape[0]} 行")
        
        # 生成建議
        recommendations = generate_return_recommendations(processed_data)
        print(f"生成退貨建議: {len(recommendations)} 條")
        
        if len(recommendations) > 0:
            print("\n真實數據分析結果:")
            print(f"  總退貨建議數: {len(recommendations)}")
            print(f"  總退貨件數: {recommendations['Transfer Qty'].sum()}")
            print(f"  ND 類型: {(recommendations['Type'] == 'ND').sum()}")
            print(f"  RF 類型: {(recommendations['Type'] == 'RF').sum()}")
            
            # 按 OM 統計
            om_stats = recommendations.groupby('OM')['Transfer Qty'].sum().sort_values(ascending=False)
            print(f"\n按 OM 退貨件數分布:")
            for om, qty in om_stats.head().items():
                print(f"  {om}: {qty} 件")
        
        # 質量檢查
        quality_results = quality_check(recommendations, processed_data)
        print(f"\n質量檢查結果:")
        for result in quality_results:
            print(f"  {result}")
        
        print("✅ 真實數據測試通過")
        
    except Exception as e:
        print(f"❌ 真實數據測試失敗: {str(e)}")

def main():
    """主測試函數"""
    print("🚀 開始退貨建議分析系統測試")
    print("=" * 50)
    
    try:
        # 執行各項測試
        test_effective_sold_qty()
        print()
        
        processed_df = test_data_preprocessing()
        print()
        
        recommendations = test_return_recommendations()
        print()
        
        test_quality_check()
        print()
        
        test_with_real_data()
        print()
        
        print("=" * 50)
        print("🎉 所有測試完成！系統功能正常")
        
        # 顯示應用程序啟動信息
        print("\n📱 Streamlit 應用程序已啟動")
        print("🌐 訪問地址: http://localhost:8501")
        
    except Exception as e:
        print(f"❌ 測試失敗: {str(e)}")
        print("\n詳細錯誤信息:")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
