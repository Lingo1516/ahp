import streamlit as st
import pandas as pd
import numpy as np

# --- 頁面設定 ---
st.set_page_config(page_title="AHP 層級分析系統 V5.0", layout="wide")

# --- 核心數學函式 ---
def repair_matrix(matrix):
    """修復矩陣：強制對角線為1，補全左下角"""
    # 確保矩陣是浮點數
    matrix = np.array(matrix, dtype=float)
    rows, cols = matrix.shape
    
    for i in range(rows):
        for j in range(cols):
            if i == j: 
                matrix[i, j] = 1.0
            elif i < j:
                # 右上角：如果讀到 0 或 NaN，預設補 1
                if matrix[i, j] == 0 or np.isnan(matrix[i, j]): 
                    matrix[i, j] = 1.0
                # 左下角：強制倒數
                if matrix[i, j] != 0:
                    matrix[j, i] = 1.0 / matrix[i, j]
                else:
                    matrix[j, i] = 1.0 # 避免除以零
    return matrix

def calculate_ahp(matrix):
    """計算 AHP，回傳權重與 CR"""
    matrix = repair_matrix(matrix)
    n = matrix.shape[0]
    col_sums = matrix.sum(axis=0)
    with np.errstate(divide='ignore', invalid='ignore'):
        normalized_matrix = matrix / col_sums
    weights = normalized_matrix.mean(axis=1)
    
    lambda_max = np.dot(col_sums, weights)
    ci = (lambda_max - n) / (n - 1) if n > 1 else 0
    ri_table = {1:0, 2:0, 3:0.58, 4:0.90, 5:1.12, 6:1.24, 7:1.32, 8:1.41, 9:1.45, 10:1.49}
    ri = ri_table.get(n, 1.49)
    cr = ci / ri if n > 2 else 0
    return weights, cr, matrix

def geometric_mean_matrix(matrices):
    """多專家幾何平均"""
    if not matrices: return None
    stack = np.array(matrices)
    prod = np.prod(stack, axis=0)
    geo_mean = np.power(prod, 1/len(matrices))
    return geo_mean

# --- 主程式介面 ---

st.title("⚖️ AHP 層級分析系統 (V5.0 強制裁切版)")
st.markdown("解決「權重一樣」與「讀到空白格」的問題。")

tab1, tab2 = st.tabs(["Step 1: 計算局部權重", "Step 2: 整合全球權重"])

# === Tab 1: 權重計算器 ===
with tab1:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.info("💡 操作提示：\n1. 上傳 Excel。\n2. 若讀取範圍錯誤 (例如出現 8 個指標)，請調整下方的「裁切設定」。")
        uploaded_file = st.file_uploader("上傳 Excel 檔", type=['xlsx', 'xls'])
        
        # --- 關鍵功能：手動裁切 ---
        st.write("---")
        st.markdown("**✂️ 矩陣裁切設定**")
        manual_n = st.number_input("強制設定指標數量 (N)", min_value=0, max_value=15, value=0, help="設為 0 代表自動偵測。若您只填了 3 個指標卻跑出 8 個，請手動改成 3。")

    with col2:
        if uploaded_file is not None:
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                valid_matrices = []
                
                st.write(f"📄 偵測到 {len(sheet_names)} 位專家資料")

                for sheet in sheet_names:
                    # 1. 讀取資料
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                    
                    # 2. 資料清理：轉數字
                    df = df.apply(pd.to_numeric, errors='coerce')
                    
                    # 3. 抓取矩陣 (自動移除全空的行列)
                    df_clean = df.dropna(how='all').dropna(axis=1, how='all')
                    raw_matrix = df_clean.values
                    
                    # 4. 強制裁切 (關鍵步驟！)
                    if manual_n > 0:
                        # 如果使用者指定了 N，就只取左上角的 NxN
                        if raw_matrix.shape[0] >= manual_n and raw_matrix.shape[1] >= manual_n:
                            raw_matrix = raw_matrix[:manual_n, :manual_n]
                    
                    rows, cols = raw_matrix.shape
                    
                    # 5. 驗證形狀
                    if rows == cols and rows > 1:
                        valid_matrices.append(raw_matrix)
                    else:
                        st.warning(f"⚠️ 工作表 {sheet} 格式異常 (大小 {rows}x{cols})，已略過。")

                if valid_matrices:
                    # 顯示它到底讀到了什麼 (Debug 視窗)
                    with st.expander("🔍 點此檢查：系統讀到的矩陣數據 (第一位專家)", expanded=True):
                        st.write(f"目前矩陣大小：**{valid_matrices[0].shape[0]} x {valid_matrices[0].shape[0]}**")
                        st.dataframe(pd.DataFrame(valid_matrices[0]))
                        if valid_matrices[0].shape[0] > 3 and manual_n == 0:
                            st.error("❗ 注意：如果您只填了 3 個指標，但上面顯示 8x8 或更大，請將左側的「強制設定指標數量」改為 3！")

                    # 進行計算
                    final_matrix = geometric_mean_matrix(valid_matrices)
                    weights, cr, _ = calculate_ahp(final_matrix)
                    
                    st.success("✅ 計算完成！")
                    
                    # 結果顯示
                    res_col1, res_col2 = st.columns(2)
                    with res_col1:
                        st.metric("整合後 CR 值", f"{cr:.4f}", delta="合格" if cr < 0.1 else "不一致", delta_color="inverse")
                    
                    # 表格
                    df_res = pd.DataFrame({
                        "指標": [f"指標 {i+1}" for i in range(len(weights))],
                        "權重": weights
                    })
                    st.dataframe(df_res.style.format({"權重": "{:.2%}"}).background_gradient(cmap="Blues"))
                    
                    st.caption("請複製此處權重，填入 Step 2 進行整合。")

                else:
                    st.error("無法讀取有效矩陣。請確認 Excel 內容或嘗試調整裁切設定。")

            except Exception as e:
                st.error(f"發生錯誤：{e}")

# === Tab 2: 全球權重整合 ===
with tab2:
    st.markdown("### 🌍 全球權重計算表")
    st.info("請將 Step 1 算出的「構面權重」與「準則權重」填入下方。")

    if "grid_data" not in st.session_state:
        st.session_state.grid_data = pd.DataFrame(
            [
                {"構面": "構面A", "構面權重": 0.5, "準則": "準則A1", "準則局部權重": 0.6},
                {"構面": "構面A", "構面權重": 0.5, "準則": "準則A2", "準則局部權重": 0.4},
            ]
        )

    edited_df = st.data_editor(st.session_state.grid_data, num_rows="dynamic", use_container_width=True)

    if st.button("計算最終排名"):
        res = edited_df.copy()
        res["構面權重"] = pd.to_numeric(res["構面權重"], errors='coerce').fillna(0)
        res["準則局部權重"] = pd.to_numeric(res["準則局部權重"], errors='coerce').fillna(0)
        res["全球權重"] = res["構面權重"] * res["準則局部權重"]
        res = res.sort_values("全球權重", ascending=False).reset_index(drop=True)
        
        st.dataframe(res.style.format({
            "構面權重": "{:.2%}", "準則局部權重": "{:.2%}", "全球權重": "{:.2%}"
        }))
