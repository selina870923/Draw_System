import streamlit as st
import pandas as pd
import random
import io

# ==========================================
# 核心演算法：符合「序號、單位、名稱、抽籤序號」格式
# ==========================================
def table_tennis_draw(df):
    # 定義標準欄位名稱
    COL_ID = "序號"
    COL_UNIT = "單位"
    COL_NAME = "名稱"
    COL_SEED = "抽籤序號"
    
    players = df.to_dict('records')
    num_players = len(players)
    
    # 建立空的籤位表 (1 到 N)
    slots = {i: None for i in range(1, num_players + 1)}
    
    # 1. 處理預排種子 (讀取「抽籤序號」欄位)
    remaining_players = []
    for p in players:
        try:
            seed_val = p.get(COL_SEED)
            if pd.notna(seed_val) and str(seed_val).strip().isdigit():
                pos = int(seed_val)
                if 1 <= pos <= num_players and slots[pos] is None:
                    slots[pos] = p
                else:
                    remaining_players.append(p)
            else:
                remaining_players.append(p)
        except:
            remaining_players.append(p)
            
    # 2. 隨機打亂其餘選手
    random.shuffle(remaining_players)
    empty_slots = [i for i, v in slots.items() if v is None]
    
    # 3. 填入剩餘選手並避開同單位
    for p in remaining_players:
        my_unit = str(p.get(COL_UNIT, ''))
        best_slot = None
        
        for slot_idx in empty_slots:
            # 判斷對手位置 (1-2, 3-4, 5-6...)
            opponent_idx = slot_idx + 1 if slot_idx % 2 != 0 else slot_idx - 1
            
            if opponent_idx in slots and slots[opponent_idx] is not None:
                # 檢查對手是否來自同一單位
                if str(slots[opponent_idx].get(COL_UNIT, '')) != my_unit:
                    best_slot = slot_idx
                    break
            else:
                best_slot = slot_idx
                break
        
        if best_slot is None:
            best_slot = empty_slots[0]
            
        slots[best_slot] = p
        empty_slots.remove(best_slot)

    # 4. 產出最終結果
    result_list = []
    for i in range(1, num_players + 1):
        if slots[i] is not None:
            p_data = slots[i].copy()
            p_data['結果籤號'] = i
            result_list.append(p_data)
        
    return pd.DataFrame(result_list)

# ==========================================
# Streamlit 介面層
# ==========================================
def main():
    st.set_page_config(page_title="桌球比賽抽籤系統", layout="wide")
    st.title("🏓 專業桌球抽籤系統")
    st.markdown("---")
    st.write("📋 **Excel 格式要求**：需包含 `序號`、`單位`、`名稱`、`抽籤序號` 四個欄位。")

    uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        output = io.BytesIO()
        processed_any = False
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet in xl.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                
                # 嚴格檢查欄位
                required = ["序號", "單位", "名稱", "抽籤序號"]
                if all(c in df.columns for c in required):
                    st.success(f"正在處理分頁：{sheet}")
                    
                    result_df = table_tennis_draw(df)
                    
                    # 重新排序顯示欄位
                    display_order = ['結果籤號', '單位', '名稱', '序號', '抽籤序號']
                    st.dataframe(result_df[display_order], use_container_width=True)
                    
                    result_df.to_excel(writer, sheet_name=sheet, index=False)
                    processed_any = True
                else:
                    st.error(f"分頁 `{sheet}` 格式不符！請確認欄位名稱完全正確。")

        if processed_any:
            st.download_button(
                label="📥 下載最終抽籤結果",
                data=output.getvalue(),
                file_name="桌球抽籤結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()

