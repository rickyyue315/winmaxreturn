#!/bin/bash
# 退貨建議分析系統 - 快速啟動腳本

echo "🚀 退貨建議分析系統啟動中..."
echo "=" * 50

# 檢查 Python 版本
python_version=$(python --version 2>&1)
echo "📋 Python 版本: $python_version"

# 安裝依賴包
echo "📦 安裝依賴包..."
pip install -r requirements.txt

# 檢查文件完整性
echo "🔍 檢查文件完整性..."
if [ -f "app.py" ]; then
    echo "✅ app.py 存在"
else
    echo "❌ app.py 不存在"
    exit 1
fi

if [ -f "user_input_files/ELE_15Sep2025.XLSX" ]; then
    echo "✅ 默認數據文件存在"
else
    echo "⚠️  默認數據文件不存在，請準備 Excel 文件"
fi

# 運行測試
echo "🧪 運行功能測試..."
python test_app.py

# 啟動應用程序
echo ""
echo "🌐 啟動 Streamlit 應用程序..."
echo "📱 訪問地址: http://localhost:8501"
echo "🛑 按 Ctrl+C 停止應用程序"
echo ""

streamlit run app.py --server.port 8501 --server.address 0.0.0.0
