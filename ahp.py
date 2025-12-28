import streamlit as st
import pandas as pd
import numpy as np

# --- 頁面設定 (極簡模式) ---
st.set_page_config(page_title="AHP 層級權重計算", layout="centered") 

# --- 核心數學函式 (不變，但移除囉唆的檢查) ---
def repair_matrix(matrix):
    """修復矩陣：強制對角線為1，補全左下角"""
    rows, cols = matrix.shape
    matrix = matrix.astype(float)
    for i in range(rows):
        for j in range(cols):
            if i == j: matrix[i, j] = 1.0
            elif i < j:
                if matrix[i, j] == 0 or np.isnan(matrix[i, j]): matrix[i, j] = 1.0
                matrix[j, i] = 1.0 / matrix[i, j]
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
    ci = (lambda_max - n) / (n - 1)
    ri_table = {1:0, 2:0, 3:0.58, 4:0.90, 5:1.12, 6:1.24, 7:1.32, 8:1.41, 9:1.45, 10:1.49}
    ri = ri_table.get(n, 1.49)
    cr = ci / ri if n > 2 else 0
    return weights, cr

def geometric_mean_matrix(matrices):
    """多專家幾何平均"""
    stack = np.array(matrices)
    prod = np.prod(stack, axis=0)
    geo_mean = np.power(prod, 1/len(matrices))
    return geo_mean

# --- 主程式介面 ---

st.title("⚖️ AHP 極簡權重計算器")

# 使用 Tab 分流，讓畫面不雜亂
tab1, tab2 = st.tabs(["Step 1: 上傳計算權重", "Step 2: 計算全球權重"])

# === Tab 1: 單一檔案計算器 ===
with tab1:
    st.markdown("### 📥 單層權重計算")
    st.info("說明：請依序上傳「構面」或「各準則」的 Excel 檔。計算出權重後，請抄寫或複製下來，填入 Step 2。")
    
    uploaded_file = st.file_uploader("上傳 Excel 檔 (支援多專家 Sheet)", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            valid_matrices = []
            
            # 靜默處理所有 Sheet
            for sheet in sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                df_numeric = df.apply(pd.to_numeric, errors='coerce')
                df_clean = df_numeric.dropna(how='all').dropna(axis=1, how='all')
                raw_matrix = df_clean.values
                
                rows, cols = raw_matrix.shape
                if rows > 2 and rows == cols:
                    # 嘗試修復並加入
                    try:
                        repaired = repair_matrix(raw_matrix)
                        valid_matrices.append(repaired)
                    except:
                        pass

            if valid_matrices:
                # 直接進行群體整合
                final_matrix = geometric_mean_matrix(valid_matrices)
                final_weights, final_cr = calculate_ahp(final_matrix)
                
                # --- 結果顯示區 (極簡化) ---
                st.success(f"✅ 計算完成 (整合了 {len(valid_matrices)} 位專家)")
                
                if final_cr > 0.1:
                    st.warning(f"⚠️ 注意：整合後 CR 值為 {final_cr:.4f} (大於 0.1)，但下方仍顯示權重供參考。")
                else:
                    st.caption(f"一致性檢定通過 (CR = {final_cr:.4f})")

                # 只顯示純淨的表格
                df_res = pd.DataFrame({
                    "項目名稱 (自行對照)": [f"項目 {i+1}" for i in range(len(final_weights))],
                    "權重 (Weight)": final_weights
                })
                # 格式化顯示百分比，但保留原始數值方便複製
                st.dataframe(df_res.style.format({"權重 (Weight)": "{:.4%}"}))

            else:
                st.error("無法讀取有效矩陣，請檢查 Excel 格式。")

        except Exception as e:
            st.error(f"錯誤：{e}")

# === Tab 2: 全球權重整合表 ===
with tab2:
    st.markdown("### 🌍 全球權重 (Global Weight) 整合")
    st.markdown("請將 Step 1 算出的數據填入下方表格：")

    # 初始化預設表格數據
    if "grid_data" not in st.session_state:
        st.session_state.grid_data = pd.DataFrame(
            [
                {"構面名稱": "構面A", "構面權重": 0.5, "準則名稱": "準則A-1", "準則局部權重": 0.6},
                {"構面名稱": "構面A", "構面權重": 0.5, "準則名稱": "準則A-2", "準則局部權重": 0.4},
                {"構面名稱": "構面B", "構面權重": 0.5, "準則名稱": "準則B-1", "準則局部權重": 0.3},
                {"構面名稱": "構面B", "構面權重": 0.5, "準則名稱": "準則B-2", "準則局部權重": 0.7},
            ]
        )

    # 可編輯的表格
    edited_df = st.data_editor(st.session_state.grid_data, num_rows="dynamic", use_container_width=True)

    # 自動計算按鈕
    if st.button("計算最終排名"):
        # 計算全球權重
        result_df = edited_df.copy()
        # 確保是數字
        result_df["構面權重"] = pd.to_numeric(result_df["構面權重"], errors='coerce').fillna(0)
        result_df["準則局部權重"] = pd.to_numeric(result_df["準則局部權重"], errors='coerce').fillna(0)
        
        # 核心公式：全球權重 = 構面權重 * 準則局部權重
        result_df["全球權重"] = result_df["構面權重"] * result_df["準則局部權重"]
        
        # 排序
        result_df = result_df.sort_values(by="全球權重", ascending=False).reset_index(drop=True)
        
        # 顯示結果
        st.write("### 🏆 最終分析結果")
        st.dataframe(result_df.style.format({
            "構面權重": "{:.4%}", 
            "準則局部權重": "{:.4%}", 
            "全球權重": "{:.4%}"
        }))
