import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter
from PIL import Image

# ==========================================
# 1. データベース定義（部材・単価マスタ）
# ==========================================

# シリーズ定義
SERIES_NAMES = {
    "fullcolor": "Panasonic フルカラー(モダン)",
    "cosmo": "Panasonic コスモワイド21",
    "advance": "Panasonic アドバンス",
    "adv_metal": "Panasonic アドバンス(新金属)",
    "select": "Panasonic セレクトプレート",
    "sostyle": "Panasonic SO-STYLE",
    "classic": "Panasonic クラシック",
    "extra": "Panasonic エクストラ",
    "jimbo": "JIMBO NKシリーズ"
}

# ▼ ハンドル（操作板）基本単価マスタ (標準・表示なし・ネームなし)
HANDLES_DB = {
    "single": {
        "name": "シングルハンドル", 
        "fullcolor": 0, "cosmo": 110, "advance": 320, "adv_metal": 320, 
        "select": 320, "sostyle": 450, "classic": 0, "extra": 320, "jimbo": 0
    },
    "double": {
        "name": "ダブルハンドル(一式)", 
        "fullcolor": 0, "cosmo": 220, "advance": 640, "adv_metal": 640, 
        "select": 640, "sostyle": 900, "classic": 0, "extra": 640, "jimbo": 0
    },
    "triple": {
        "name": "トリプルハンドル(一式)", 
        "fullcolor": 0, "cosmo": 330, "advance": 960, "adv_metal": 960, 
        "select": 960, "sostyle": 1350, "classic": 0, "extra": 960, "jimbo": 0
    },
}

# ▼ ハンドル仕様による加算額マスタ (標準価格に対する差額)
# キーの意味: [window:窓の有無][name:ネームの有無]
# 例: "win_name" = 表示窓あり・ネームあり
HANDLE_ADDERS = {
    # 標準 (窓なし・ネームなし)
    "std_std":  {"name": "標準", "cosmo": 0, "advance": 0, "sostyle": 0, "fullcolor": 0, "other": 0},
    # ネームのみ (窓なし・ネームあり)
    "std_name": {"name": "ネーム付", "cosmo": 70, "advance": 0, "sostyle": 0, "fullcolor": 20, "other": 0},
    # 表示のみ (窓あり・ネームなし)
    "win_std":  {"name": "表示付", "cosmo": 70, "advance": 0, "sostyle": 0, "fullcolor": 50, "other": 0},
    # 両方 (窓あり・ネームあり)
    "win_name": {"name": "表示+ネーム付", "cosmo": 140, "advance": 20, "sostyle": 0, "fullcolor": 70, "other": 0},
}

# ▼ 部材マスタ（has_lamp: ホタルなど窓が必要なものはTrue）
ITEMS_DB = {
    # --- スイッチ類 ---
    "sw_b_mech": {
        "name": "片切スイッチ", "icon": "🔘", "img_file": "sw_b.jpg", "has_lamp": False,
        "fullcolor": 250, "cosmo": 270, "advance": 610, "adv_metal": 610, 
        "select": 610, "sostyle": 900, "classic": 1430, "extra": 610, "jimbo": 1800
    },
    "sw_h_mech": {
        "name": "ほたるスイッチ", "icon": "🟢", "img_file": "sw_b.jpg", "has_lamp": True,
        "fullcolor": 630, "cosmo": 1050, "advance": 1495, "adv_metal": 1495, 
        "select": 1495, "sostyle": 1970, "classic": 1430, "extra": 1495, "jimbo": 2900
    },
    "sw_3_mech": {
        "name": "3路スイッチ", "icon": "🔄", "img_file": "sw_3.jpg", "has_lamp": False,
        "fullcolor": 430, "cosmo": 420, "advance": 930, "adv_metal": 930, 
        "select": 930, "sostyle": 1500, "classic": 2040, "extra": 930, "jimbo": 2200
    },
    "sw_3h_mech": {
        "name": "3路ほたるSW", "icon": "🔄🟢", "img_file": "sw_3.jpg", "has_lamp": True,
        "fullcolor": 850, "cosmo": 1650, "advance": 2300, "adv_metal": 2300, 
        "select": 2300, "sostyle": 2900, "classic": 2040, "extra": 2300, "jimbo": 3300
    },
    "sw_4_mech": {
        "name": "4路スイッチ", "icon": "🔀", "img_file": "sw_4.jpg", "has_lamp": False,
        "fullcolor": 1600, "cosmo": 1600, "advance": 2400, "adv_metal": 2400, 
        "select": 2400, "sostyle": 3500, "classic": 3960, "extra": 2400, "jimbo": 3200
    },
    "sw_4h_mech": {
        "name": "4路ほたるSW", "icon": "🔀🟢", "img_file": "sw_4.jpg", "has_lamp": True,
        "fullcolor": 2100, "cosmo": 3800, "advance": 4600, "adv_metal": 4600, 
        "select": 4600, "sostyle": 5300, "classic": 3960, "extra": 4600, "jimbo": 4200
    },

    # --- コンセント類 ---
    "outlet_w": {
        "name": "ダブルコンセント", "icon": "🔌", "img_file": "outlet_w.jpg", "has_lamp": False,
        "fullcolor": 380, "cosmo": 550, "advance": 800, "adv_metal": 800, 
        "select": 800, "sostyle": 1200, "classic": 380, "extra": 800, "jimbo": 1300
    },
    "outlet_e": {
        "name": "アース付コンセント", "icon": "⏚", "img_file": "outlet_e.jpg", "has_lamp": False,
        "fullcolor": 450, "cosmo": 600, "advance": 900, "adv_metal": 900, 
        "select": 900, "sostyle": 1300, "classic": 450, "extra": 900, "jimbo": 1500
    },
    "tv_4k": {
        "name": "TV端子(4K8K)", "icon": "📺", "img_file": "tv_4k.jpg", "has_lamp": False,
        "fullcolor": 1400, "cosmo": 1400, "advance": 1700, "adv_metal": 1700, 
        "select": 1700, "sostyle": 2100, "classic": 1400, "extra": 1700, "jimbo": 2300
    },
    "lan_6": {
        "name": "LAN(CAT6)", "icon": "💻", "img_file": "lan_6.jpg", "has_lamp": False,
        "fullcolor": 2090, "cosmo": 2090, "advance": 2500, "adv_metal": 2500, 
        "select": 2500, "sostyle": 3500, "classic": 2090, "extra": 2500, "jimbo": 3200
    },
    "blank": {
        "name": "空白・ブランク", "icon": "⬜", "img_file": "blank.jpg", "has_lamp": False,
        "fullcolor": 0, "cosmo": 0, "advance": 0, "adv_metal": 0, 
        "select": 0, "sostyle": 0, "classic": 0, "extra": 0, "jimbo": 300
    },
}

# プレート・取付枠マスタ
PARTS_DB = {
    "plate_1": {
        "name": "1連プレート", 
        "fullcolor": 220, "cosmo": 130, "advance": 600, "adv_metal": 730, 
        "select": 1400, "sostyle": 900, "classic": 1100, "extra": 5000, "jimbo": 600
    },
    "plate_2": {
        "name": "2連プレート", 
        "fullcolor": 440, "cosmo": 260, "advance": 1200, "adv_metal": 1460, 
        "select": 2800, "sostyle": 1800, "classic": 2200, "extra": 10000, "jimbo": 1200
    },
    "plate_3": {
        "name": "3連プレート", 
        "fullcolor": 660, "cosmo": 390, "advance": 1800, "adv_metal": 2920, 
        "select": 4200, "sostyle": 2700, "classic": 4400, "extra": 15000, "jimbo": 2000
    },
    "frame": {
        "name": "取付枠", 
        "fullcolor": 60, "cosmo": 70, "advance": 120, "adv_metal": 120, 
        "select": 120, "sostyle": 150, "classic": 60, "extra": 120, "jimbo": 100
    },
}

# ==========================================
# 2. アプリ設定・関数
# ==========================================
st.set_page_config(page_title="電材差額見積アプリ(Pro)", layout="wide")
if 'estimate_list' not in st.session_state:
    st.session_state.estimate_list = []

# 画像表示関数
def show_item_image(item_key):
    if item_key in ITEMS_DB:
        item_data = ITEMS_DB[item_key]
        img_filename = item_data.get("img_file", "")
        img_path = os.path.join("img", img_filename)
        if os.path.exists(img_path):
            try:
                st.image(Image.open(img_path), use_column_width=True)
            except:
                st.write(item_data["icon"])
        else:
            st.markdown(f"<h1 style='text-align: center;'>{item_data['icon']}</h1>", unsafe_allow_html=True)

# 差額計算ヘルパー
def get_handle_price_diff(handle_type, series_key, needs_window, needs_name):
    # 基本ハンドル価格
    base_price_src = HANDLES_DB[handle_type].get(source_series_key, 0)
    base_price_tgt = HANDLES_DB[handle_type].get(target_series_key, 0)
    
    # オプション加算キーの生成
    opt_key_window = "win" if needs_window else "std"
    opt_key_name = "name" if needs_name else "std"
    full_opt_key = f"{opt_key_window}_{opt_key_name}"
    
    # 加算額の取得 (シリーズごとに異なる)
    def get_adder(series, key):
        adder_data = HANDLE_ADDERS.get(key, {})
        if series in adder_data: return adder_data[series]
        return adder_data.get("other", 0)

    adder_src = get_adder(source_series_key, full_opt_key)
    adder_tgt = get_adder(target_series_key, full_opt_key)
    
    return (base_price_tgt + adder_tgt) - (base_price_src + adder_src)

# ==========================================
# 3. サイドバー
# ==========================================
st.sidebar.header("🏠 物件情報")
client_name = st.sidebar.text_input("施主名", placeholder="例：山田 太郎 様")
hm_name = st.sidebar.text_input("HM名", placeholder="例：〇〇工務店 様")

st.sidebar.markdown("---")
source_series_key = st.sidebar.selectbox("【現在】変更元", list(SERIES_NAMES.keys()), index=1, format_func=lambda x: SERIES_NAMES[x])
target_series_key = st.sidebar.selectbox("【変更】変更先", list(SERIES_NAMES.keys()), index=8, format_func=lambda x: SERIES_NAMES[x])

# ==========================================
# 4. メイン画面
# ==========================================
st.title("⚡ 電材差額見積りアプリ Pro")
st.caption(f"現在の設定： {SERIES_NAMES[source_series_key]} ➡ {SERIES_NAMES[target_series_key]}")

tab1, tab2, tab3 = st.tabs(["📝 基本(1連)クイック", "🏗️ 多連・詳細ビルダー", "📄 見積書発行"])

# ------------------------------------------
# TAB 1: 簡易入力
# ------------------------------------------
with tab1:
    st.header("基本スイッチ・コンセント入力")
    is_name_req_simple = st.checkbox("📛 すべて「ネーム付」にする（+差額）", value=False)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("基本スイッチ")
        qty_sw_b = st.number_input("片切スイッチ", min_value=0)
        qty_sw_h = st.number_input("ほたるスイッチ", min_value=0)
        qty_out_w = st.number_input("ダブルコンセント", min_value=0)
    with col2:
        st.subheader("多路・機能スイッチ")
        qty_sw_3 = st.number_input("3路スイッチ", min_value=0)
        qty_sw_3h = st.number_input("3路ほたるスイッチ", min_value=0)
        qty_sw_4 = st.number_input("4路スイッチ", min_value=0)
        qty_sw_4h = st.number_input("4路ほたるスイッチ", min_value=0)

    if st.button("STEP1 追加"):
        def add_simple(item_key, qty, handle_type="single"):
            if qty > 0:
                item = ITEMS_DB[item_key]
                # 本体差額
                d_body = item[target_series_key] - item[source_series_key]
                
                # ハンドル差額（自動判定）
                # ホタル(has_lamp=True)なら窓必須。ネームはチェックボックス依存。
                needs_window = item.get("has_lamp", False)
                
                # コンセント類はハンドルがないので0円
                if "outlet" in item_key or "tv" in item_key or "lan" in item_key:
                    d_hdl = 0
                else:
                    d_hdl = get_handle_price_diff(handle_type, source_series_key, needs_window, is_name_req_simple)

                d_frm = PARTS_DB['frame'][target_series_key] - PARTS_DB['frame'][source_series_key]
                d_plt = PARTS_DB['plate_1'][target_series_key] - PARTS_DB['plate_1'][source_series_key]
                
                unit = d_body + d_hdl + d_frm + d_plt
                
                # 詳細文字列の作成
                detail_txt = "標準セット"
                if needs_window: detail_txt += "(表示付)"
                if is_name_req_simple and d_hdl != 0: detail_txt += "(ネーム付)"

                st.session_state.estimate_list.append({
                    "type": "1連(基本)", "name": item['name'], "detail": detail_txt,
                    "unit_diff": unit, "qty": qty, "total_diff": unit * qty
                })
        add_simple("sw_b_mech", qty_sw_b); add_simple("sw_h_mech", qty_sw_h)
        add_simple("sw_3_mech", qty_sw_3); add_simple("sw_3h_mech", qty_sw_3h)
        add_simple("sw_4_mech", qty_sw_4); add_simple("sw_4h_mech", qty_sw_4h)
        add_simple("outlet_w", qty_out_w) 
        st.success("基本項目を追加しました！")

# ------------------------------------------
# TAB 2: 詳細ビルダー (自動判定ロジック強化)
# ------------------------------------------
with tab2:
    st.header("詳細ビルダー：画像確認モード")
    plate_size = st.radio("プレートサイズ", ["1連", "2連", "3連"], horizontal=True)
    cols_num = 1
    if plate_size == "2連": cols_num = 2
    elif plate_size == "3連": cols_num = 3
    
    st.markdown("---")
    ui_cols = st.columns(cols_num)
    column_configs = []
    
    for i in range(cols_num):
        with ui_cols[i]:
            st.info(f"【{i+1}列目】")
            layout_type = st.selectbox("割り付け",["シングル(1個)", "ダブル(2個)", "トリプル(3個)", "コンセント(一体)"], key=f"layout_{i}")
            
            # ▼ ハンドルオプション（列ごと）
            is_name_col = False
            if layout_type != "コンセント(一体)":
                is_name_col = st.checkbox("📛 ネーム付にする", key=f"name_opt_{i}")

            items_in_col = []
            opt_list = list(ITEMS_DB.keys())
            
            def item_selector_with_image(label, k):
                c_input, c_img = st.columns([3, 1])
                with c_input:
                    sel = st.selectbox(label, opt_list, format_func=lambda x: ITEMS_DB[x]['name'], key=k)
                with c_img:
                    show_item_image(sel)
                return sel

            if layout_type == "シングル(1個)":
                item = item_selector_with_image("中身", f"c{i}_1")
                items_in_col.append(item); handle_key = "single"
            elif layout_type == "ダブル(2個)":
                item1 = item_selector_with_image("上段", f"c{i}_1")
                item2 = item_selector_with_image("下段", f"c{i}_2")
                items_in_col.extend([item1, item2]); handle_key = "double"
            elif layout_type == "トリプル(3個)":
                item1 = item_selector_with_image("上段", f"c{i}_1")
                item2 = item_selector_with_image("中段", f"c{i}_2")
                item3 = item_selector_with_image("下段", f"c{i}_3")
                items_in_col.extend([item1, item2, item3]); handle_key = "triple"
            else: # コンセント
                c_input, c_img = st.columns([3, 1])
                with c_input:
                    item = st.selectbox("種別", ["outlet_w", "outlet_e", "tv_4k", "lan_6"], format_func=lambda x: ITEMS_DB[x]['name'], key=f"c{i}_1")
                with c_img:
                    show_item_image(item)
                items_in_col.append(item); handle_key = "single"
            
            column_configs.append({"items": items_in_col, "handle": handle_key, "is_name": is_name_col})

    st.markdown("---")
    qty_build = st.number_input("個数", min_value=1, value=1)
    
    if st.button("見積に追加", key="add_build"):
        p_key = "plate_1"
        if plate_size == "2連": p_key = "plate_2"
        elif plate_size == "3連": p_key = "plate_3"
        diff_plate = PARTS_DB[p_key][target_series_key] - PARTS_DB[p_key][source_series_key]
        
        diff_cols_total = 0
        details_str = []
        for idx, config in enumerate(column_configs):
            d_frame = PARTS_DB['frame'][target_series_key] - PARTS_DB['frame'][source_series_key]
            
            # --- ハンドル自動判定ロジック ---
            # 1. コンセントならハンドル代は0
            if "コンセント" in str(config['handle']) or "outlet" in str(config['items'][0]):
                d_handle = 0
            else:
                # 2. 列の中に「ホタル」など窓必須アイテムがあるかチェック
                # any()を使って、選ばれたアイテムのどれか1つでも has_lamp=True なら窓ありハンドルにする
                needs_window = any(ITEMS_DB[itm].get("has_lamp", False) for itm in config['items'])
                
                # 3. 差額計算（窓の有無 + ネームの有無）
                d_handle = get_handle_price_diff(config['handle'], source_series_key, needs_window, config['is_name'])

            d_items = 0
            item_names = []
            for itm in config['items']:
                d_items += ITEMS_DB[itm][target_series_key] - ITEMS_DB[itm][source_series_key]
                item_names.append(ITEMS_DB[itm]['name'])
            
            diff_cols_total += (d_frame + d_handle + d_items)
            
            # 詳細表記の作成
            h_type_str = config['handle']
            if config.get('is_name'): h_type_str += "(ネーム)"
            col_detail = f"[{idx+1}列目:{h_type_str}] " + ",".join(item_names)
            details_str.append(col_detail)
            
        total_unit_diff = diff_plate + diff_cols_total
        st.session_state.estimate_list.append({
            "type": f"{plate_size}カスタム", "name": "詳細構成セット", "detail": " / ".join(details_str),
            "unit_diff": total_unit_diff, "qty": qty_build, "total_diff": total_unit_diff * qty_build
        })
        st.success("追加しました！")

# ------------------------------------------
# TAB 3: 見積書発行
# ------------------------------------------
with tab3:
    st.header("見積りプレビュー")
    if st.session_state.estimate_list:
        df = pd.DataFrame(st.session_state.estimate_list)
        st.dataframe(df[["type", "name", "detail", "unit_diff", "qty", "total_diff"]], use_container_width=True)
        grand_total = df["total_diff"].sum()
        st.metric("総計(税抜)", f"¥ {grand_total:,.0f}")
        
        def to_excel(df, client, hm, src, tgt, total):
            output = io.BytesIO()
            wb = xlsxwriter.Workbook(output, {'in_memory': True})
            ws = wb.add_worksheet("差額見積")
            fmt_head = wb.add_format({'bold': True, 'bg_color': '#ddd', 'border': 1})
            ws.write(0, 0, f"施主: {client}")
            ws.write(1, 0, f"HM: {hm}")
            ws.write(2, 0, f"{src} ➡ {tgt}")
            headers = ["種類", "品名", "詳細", "単価差額", "数量", "差額合計"]
            for c, h in enumerate(headers): ws.write(4, c, h, fmt_head)
            for r, row in enumerate(df.to_dict('records')):
                ws.write(5+r, 0, row['type'])
                ws.write(5+r, 1, row['name'])
                ws.write(5+r, 2, row['detail'])
                ws.write(5+r, 3, row['unit_diff'])
                ws.write(5+r, 4, row['qty'])
                ws.write(5+r, 5, row['total_diff'])
            ws.write(5+len(df), 5, total, wb.add_format({'bold':True}))
            wb.close()
            return output.getvalue()

        xl = to_excel(df, client_name, hm_name, SERIES_NAMES[source_series_key], SERIES_NAMES[target_series_key], grand_total)
        st.download_button("Excelダウンロード", xl, "見積.xlsx")
        
        if st.button("リセット"):
            st.session_state.estimate_list = []
            st.rerun()