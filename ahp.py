import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# 設定網頁標題
st.set_page_config(page_title="AHP 論文數據分析系統", layout="wide")

st.title("🏆 AHP 論文數據分析系統 (V3.0)")
st.markdown("### 支援 Excel 即時 CR 檢測 • 強化版讀取引擎")

# --- 數學運算核心函式 (Python 端) ---

def repair_matrix(matrix):
    """
    修復矩陣：
    1. 強制對角線為 1
    2. 強制左下角數值 = 1 / 右上角數值 (避免 Excel 公式讀取錯誤)
    """
    rows, cols = matrix.shape
    # 確保是浮點數型態
    matrix = matrix.astype(float)
    
    for i in range(rows):
        for j in range(cols):
            if i == j:
                matrix[i, j] = 1.0
            elif i < j:
                # 這是右上角 (使用者填寫區)，如果使用者留空或填0，預設為1
                if matrix[i, j] == 0 or np.isnan(matrix[i, j]):
                    matrix[i, j] = 1.0
                # 同步更新左下角
                matrix[j, i] = 1.0 / matrix[i, j]
                
    return matrix

def calculate_ahp(matrix):
    """計算單一矩陣的 AHP 權重與 CR"""
    # 先修復矩陣 (這是最關鍵的一步！)
    matrix = repair_matrix(matrix)
    
    n = matrix.shape[0]
    col_sums = matrix.sum(axis=0)
    
    with np.errstate(divide='ignore', invalid='ignore'):
        normalized_matrix = matrix / col_sums
        
    weights = normalized_matrix.mean(axis=1)
    
    lambda_max = np.dot(col_sums, weights)
    ci = (lambda_max - n) / (n - 1)
    
    ri_table = {1:0, 2:0, 3:0.58, 4:0.90, 5:1.12, 6:1.24, 7:1.32, 8:1.41, 9:1.45, 10:1.49, 11:1.51, 12:1.48, 13:1.56, 14:1.57, 15:1.59}
    ri = ri_table.get(n, 1.49)
    cr = ci / ri if n > 2 else 0
    
    return weights, cr, ci, matrix

def geometric_mean_matrix(matrices):
    """計算多個矩陣的幾何平均"""
    stack = np.array(matrices)
    prod = np.prod(stack, axis=0)
    geo_mean = np.power(prod, 1/len(matrices))
    return geo_mean

def generate_smart_excel(n_criteria, n_experts):
    """產生智慧型 Excel 範例"""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    # --- 定義格式 ---
    fmt_yellow = workbook.add_format({'bg_color': '#FFFFCC', 'border': 1, 'align': 'center'}) 
    fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#E0E0E0', 'border': 1})
    fmt_formula = workbook.add_format({'bg_color': '#F9F9F9', 'border': 1, 'align': 'center', 'font_color': '#555555'})
    fmt_result_good = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True, 'border': 1})
    fmt_result_bad = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'border': 1})
    fmt_guide = workbook.add_format({'italic': True, 'font_color': 'blue'})

    # RI 查表
    ri_values = [0, 0, 0, 0.58, 0.90, 1.12, 1.24, 1.32, 1.41, 1.45, 1.49, 1.51, 1.48, 1.56, 1.57, 1.59]
    current_ri = ri_values[n_criteria] if n_criteria < len(ri_values) else 1.59

    for i in range(n_experts):
        sheet_name = f'專家{i+1}'
        worksheet = workbook.add_worksheet(sheet_name)
        
        worksheet.write('A1', '請填寫黃色區域 (輸入 1~9 或小數)', fmt_guide)
        
        # --- 建立矩陣 ---
        start_row = 2
        start_col = 1
        
        # 標題
        for k in range(n_criteria):
            label = f"指標{k+1}"
            worksheet.write(start_row - 1, start_col + k, label, fmt_header)
            worksheet.write(start_row + k, start_col - 1, label, fmt_header)

        col_sums_refs = []
        weight_refs = []

        # 填寫格子
        for r in range(n_criteria):
            for c in range(n_criteria):
                cell_row = start_row + r
                cell_col = start_col + c
                
                if r == c:
                    worksheet.write(cell_row, cell_col, 1, fmt_formula)
                elif r < c:
                    worksheet.write(cell_row, cell_col, 1, fmt_yellow) # 使用者填寫區
                else:
                    target_str = xl_rowcol_to_cell(start_row + c, start_col + r)
                    worksheet.write_formula(cell_row, cell_col, f'=1/{target_str}', fmt_formula)

        # --- Excel 內部運算 (隱藏區) ---
        calc_start_row = start_row + n_criteria + 2
        worksheet.write(calc_start_row, 0, "中間運算區 (請勿更動)", fmt_guide)
        
        # 行加總
        for c in range(n_criteria):
            range_start = xl_rowcol_to_cell(start_row, start_col + c)
            range_end = xl_rowcol_to_cell(start_row + n_criteria - 1, start_col + c)
            sum_cell = xl_rowcol_to_cell(calc_start_row + 1, start_col + c)
            worksheet.write_formula(sum_cell, f'=SUM({range_start}:{range_end})', fmt_formula)
            col_sums_refs.append(sum_cell)

        # 權重計算
        norm_start_row = calc_start_row + 3
        for r in range(n_criteria):
            row_norm_refs = []
            for c in range(n_criteria):
                raw_val = xl_rowcol_to_cell(start_row + r, start_col + c)
                col_sum = col_sums_refs[c]
                norm_cell = xl_rowcol_to_cell(norm_start_row + r, start_col + c)
                worksheet.write_formula(norm_cell, f'={raw_val}/{col_sum}', fmt_formula)
                row_norm_refs.append(norm_cell)
            
            weight_cell = xl_rowcol_to_cell(norm_start_row + r, start_col + n_criteria + 1)
            range_norm_start = row_norm_refs[0]
            range_norm_end = row_norm_refs[-1]
            worksheet.write_formula(weight_cell, f'=AVERAGE({range_norm_start}:{range_norm_end})', fmt_formula)
            weight_refs.append(weight_cell)

        # CR 計算
        lambda_formula_parts = []
        for i in range(n_criteria):
            lambda_formula_parts.append(f"{col_sums_refs[i]}*{weight_refs[i]}")
        lambda_formula = "=" + "+".join(lambda_formula_parts)
        
        lambda_cell = xl_rowcol_to_cell(start_row, start_col + n_criteria + 2) 
        ci_cell = xl_rowcol_to_cell(start_row + 1, start_col + n_criteria + 2)
        cr_cell = xl_rowcol_to_cell(start_row + 2, start_col + n_criteria + 2)
        status_cell = xl_rowcol_to_cell(start_row + 3, start_col + n_criteria + 2)

        worksheet.write(start_row, start_col + n_criteria + 1, "Lambda Max:", fmt_header)
        worksheet.write(start_row + 1, start_col + n_criteria + 1, "CI:", fmt_header)
        worksheet.write(start_row + 2, start_col + n_criteria + 1, "CR 值 (即時):", fmt_header)
        worksheet.write(start_row + 3, start_col + n_criteria + 1, "狀態:", fmt_header)

        worksheet.write_formula(lambda_cell, lambda_formula, fmt_formula)
        worksheet.write_formula(ci_cell, f'=({lambda_cell}-{n_criteria})/({n_criteria}-1)', fmt_formula)
        worksheet.write_formula(cr_cell, f'={ci_cell}/{current_ri}', fmt_yellow)
        worksheet.write_formula(status_cell, f'=IF({cr_cell}<0.1, "✅ 有效", "❌ 矛盾")', fmt_formula)
        worksheet.conditional_format(cr_cell, {'type': 'cell', 'criteria': '<', 'value': 0.1, 'format': fmt_result_good})
        worksheet.conditional_format(cr_cell, {'type': 'cell', 'criteria': '>=', 'value': 0.1, 'format': fmt_result_bad})

        # 作弊建議值
        hint_start_col = start_col + n_criteria + 5
        worksheet.write(start_row - 1, hint_start_col, "💡 參考建議值 (完美一致性)", fmt_header)
        for r in range(n_criteria):
            for c in range(n_criteria):
                hint_cell = xl_rowcol_to_cell(start_row + r, hint_start_col + c)
                if r == c:
                     worksheet.write(hint_cell, 1, fmt_formula)
                else:
                    w_r = weight_refs[r]
                    w_c = weight_refs[c]
                    worksheet.write_formula(hint_cell, f'={w_r}/{w_c}', fmt_formula)

    writer.close()
    return output.getvalue()

# --- 介面佈局 ---

st.sidebar.header("📥 步驟 1：下載智慧型 Excel")
criteria_count = st.sidebar.number_input("指標數量 (N)", min_value=3, max_value=15, value=4)
expert_count = st.sidebar.number_input("專家數量", min_value=1, max_value=20, value=3)

if st.sidebar.button("產生 Excel 範例檔 (V3.0)"):
    excel_data = generate_smart_excel(criteria_count, expert_count)
    st.sidebar.download_button(
        label="點此下載智慧 Excel",
        data=excel_data,
        file_name=f"AHP_智慧問卷_{criteria_count}x{criteria_count}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.write("---")
st.header("📂 步驟 2：上傳分析")
st.info("請上傳您的 Excel 檔，系統將會自動修復讀取錯誤並進行計算。")

uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        valid_matrices = []
        expert_results = []
        
        progress_bar = st.progress(0)
        
        for idx, sheet in enumerate(sheet_names):
            # 讀取數據
            df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
            df_numeric = df.apply(pd.to_numeric, errors='coerce')
            df_clean = df_numeric.dropna(how='all').dropna(axis=1, how='all')
            
            # 取得原始矩陣 (可能含有讀取錯誤的公式)
            raw_matrix = df_clean.values
            
            rows, cols = raw_matrix.shape
            
            if rows > 2 and rows == cols:
                # 呼叫 AHP 計算，這裡面會自動修復矩陣
                weights, cr, ci, fixed_matrix = calculate_ahp(raw_matrix)
                is_pass = cr < 0.1
                
                expert_results.append({
                    "專家代號": sheet,
                    "CR 值": round(cr, 4),
                    "狀態": "✅ 有效" if is_pass else "❌ 剔除 (CR>0.1)",
                })
                
                if is_pass:
                    valid_matrices.append(fixed_matrix)
            
            progress_bar.progress((idx + 1) / len(sheet_names))

        st.success(f"分析完成！共讀取 {len(sheet_names)} 個工作表，其中 {len(valid_matrices)} 份有效。")

        if expert_results:
            st.subheader("1. 專家問卷檢定報告")
            st.table(pd.DataFrame(expert_results))

        if valid_matrices:
            st.subheader("2. 群體決策整合結果 (幾何平均法)")
            final_matrix = geometric_mean_matrix(valid_matrices)
            # 再次經過 AHP 計算取得最終權重
            final_weights, final_cr, final_ci, _ = calculate_ahp(final_matrix)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("整合後 CR 值", f"{final_cr:.4f}")
            col2.metric("一致性狀態", "極佳" if final_cr < 0.05 else ("合格" if final_cr < 0.1 else "不合格"))
            col3.metric("有效樣本數", len(valid_matrices))
            
            st.markdown("#### 各指標最終權重排名")
            chart_data = pd.DataFrame({
                "指標": [f"指標 {i+1}" for i in range(len(final_weights))],
                "權重": final_weights
            }).sort_values(by="權重", ascending=True)
            
            st.bar_chart(chart_data.set_index("指標"))
            
            rank_df = chart_data.sort_values(by="權重", ascending=False).reset_index(drop=True)
            rank_df.index += 1
            rank_df["權重"] = rank_df["權重"].apply(lambda x: f"{x:.2%}")
            st.dataframe(rank_df)
        else:
            st.error("⚠️ 警告：沒有任何一份問卷通過一致性檢定。請檢查 Excel 填寫邏輯。")

    except Exception as e:
        st.error(f"檔案解析發生錯誤：{e}")
