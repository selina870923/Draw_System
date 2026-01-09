import streamlit as st
import pandas as pd
import random
import io

# ==========================================
# 核心抽籤邏輯 (整合動態排除與最終校驗)
# ==========================================
def table_tennis_draw(df, interval_starts, col_map):
    COL_UNIT = col_map['單位']
    COL_SEED = col_map['種子序號']
    COL_DRAW = col_map['抽籤序號']
    
    df.columns = [str(c).strip() for c in df.columns]
    players = df.to_dict('records')
    num_players = len(players)
    
    # 1. 定義區間
    intervals = []
    if interval_starts:
        sorted_starts = sorted(list(set(interval_starts)))
        for i in range(len(sorted_starts)):
            start = sorted_starts[i]
            end = sorted_starts[i+1] - 1 if i+1 < len(sorted_starts) else num_players
            if start <= num_players:
                intervals.append(list(range(start, min(end, num_players) + 1)))

    # 2. 合法性檢查：單位人數 vs 區間數
    if intervals:
        unit_counts = df[COL_UNIT].value_counts()
        if unit_counts.max() > len(intervals):
            return None, f"⚠️ 非法序列！「{unit_counts.idxmax()}」有 {unit_counts.max()} 人，但僅有 {len(intervals)} 個區間。"

    available_slots = list(range(1, num_players + 1))

    # 3. 第一步：種子佔位 (確保 抽籤序號 = 種子序號)
    for p in players:
        p[COL_DRAW] = None
        seed_val = p.get(COL_SEED)
        parsed_seed = None
        if pd.notna(seed_val):
            str_val = str(seed_val).strip()
            if str_val.isdigit(): parsed_seed = int(str_val)
            else:
                try: parsed_seed = int(float(str_val))
                except: parsed_seed = None
        
        if parsed_seed is not None:
            if 1 <= parsed_seed <= num_players:
                p[COL_DRAW] = parsed_seed
                if parsed_seed in available_slots:
                    available_slots.remove(parsed_seed)
            else:
                return None, f"⚠️ 種子序號 {parsed_seed} 超出範圍。"

    # 4. 第二步：線性動態抽籤 (按 Excel 順序)
    for p in players:
        if p[COL_DRAW] is not None: continue
        unit = p[COL_UNIT]
        
        forbidden = []
        if intervals:
            for p_check in players:
                if p_check[COL_UNIT] == unit and p_check[COL_DRAW] is not None:
                    for inv in intervals:
                        if p_check[COL_DRAW] in inv:
                            forbidden.extend(inv)
                            break
        
        current_pool = [s for s in available_slots if s not in forbidden]
        if not current_pool: current_pool = available_slots
            
        picked = random.choice(current_pool)
        p[COL_DRAW] = picked
        available_slots.remove(picked)

    # 5. 第三步：最終校驗
    if intervals:
        for inv in intervals:
            units_in_this_inv = [p[COL_UNIT] for p in players if p[COL_DRAW] in inv]
            if any(pd.Series(units_in_this_inv).value_counts() > 1):
                return None, f"🚨 校驗失敗：同單位在區間 {inv[0]}-{inv[-1]} 重複，請重新執行。"

    return pd.DataFrame(players), None

# ==========================================
# Streamlit 介面設計
# ==========================================
def main():
    st.set_page_config(page_title="賽事抽籤系統", layout="wide")
    
    # 標題與簡介
    st.title("🏆 賽事抽籤系統")
    
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.markdown("""
        ### 📌 系統支援功能
        * **Excel 批次處理：** 一次上傳，同時處理多個比賽項目。
        * **種子位功能：** 確保「種子序號」等於最終籤位。
        * **同區避開機制：** 自定義區間（如 1-16 號），自動確保同單位不互撞。
        """)
    with col_info2:
        st.markdown("""
        ### 🛡️ 自動檢查機制
        * **欄位自動辨識：** 自動掃描「單位、姓名、種子」等關鍵字。
        * **種子合法性：** 檢查序號是否超出總人數。
        * **區間合法性：** 檢查單一單位人數是否超過可分配區間數。
        """)

    st.divider()

    st.markdown("""
    ### 📖 使用說明
    1. **整理 Excel 表格：**
        * 將不同組別（如：男單、女雙）放在**不同分頁**。
        * 欄位依序建議為：`序號`、`單位`、`名稱`、`種子序號`、`抽籤序號`。
        * **種子功能：** 若有種子，請在「種子序號」欄位填入最終籤位。
        * **排序建議：** 建議將「人數較多單位」與「種子選手」排在該分頁的前方。
    2. **設定區間序號（第一支籤功能）：**
        * 於側邊欄輸入各組的區間起點（如 `1, 17, 33`）。
        * 若該項目為循環賽或無須避開，請**保持空白**。
    3. **開始抽籤與下載結果**。
    """)

    # 檔案上傳
    uploaded_file = st.file_uploader("1. 上傳 Excel 檔案", type=["xlsx"])

    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        st.sidebar.header("⚙️ 第一支籤區間設定")
        
        configs = {}
        for s in xl.sheet_names:
            configs[s] = st.sidebar.text_input(f"📍 {s} 區間起點 (留白則不限)", "", key=s)

        if st.button("🚀 開始抽籤並執行校驗"):
            output = io.BytesIO()
            success_count = 0
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name in xl.sheet_names:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                    
                    # 欄位自動辨識
                    col_map = {
                        '單位': next((c for c in df.columns if any(k in c for k in ["單位", "學校", "團體"])), "單位"),
                        '名稱': next((c for c in df.columns if any(k in c for k in ["名稱", "姓名", "選手"])), "名稱"),
                        '種子序號': next((c for c in df.columns if "種子" in c), "種子序號"),
                        '抽籤序號': next((c for c in df.columns if "抽籤" in c or "結果" in c), "抽籤序號")
                    }

                    # 解析區間輸入
                    raw_input = configs[sheet_name].strip()
                    interval_starts = [int(x.strip()) for x in raw_input.split(",") if x.strip().isdigit()] if raw_input else []
                    
                    result_df, err = table_tennis_draw(df, interval_starts, col_map)
                    
                    if err:
                        st.error(f"❌ {sheet_name}: {err}")
                    else:
                        st.success(f"✅ {sheet_name} 抽籤成功！")
                        st.dataframe(result_df, hide_index=True)
                        result_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        success_count += 1

            if success_count > 0:
                st.download_button("📥 下載最終抽籤結果", output.getvalue(), "抽籤結果.xlsx")

    # --- 新增免責聲明 ---
    st.markdown("---")
    st.caption("⚠️ 本系統僅供參考，請務必於下載後人工核對抽籤結果。")

if __name__ == "__main__":
    main()
