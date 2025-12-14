import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter
from PIL import Image

# ==========================================
# 0. デザイン設定
# ==========================================
st.set_page_config(page_title="電材差額見積アプリ(Pro)", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f8fb; }
    h1, h2, h3 { color: #003366 !important; font-family: "Helvetica", sans-serif; }
    /* 通常ボタン（青） */
    div.stButton > button {
        background: linear-gradient(to bottom, #0066cc, #004499);
        color: white; border: none; border-radius: 5px; font-weight: bold;
    }
    div.stButton > button:hover { background: linear-gradient(to bottom, #0055bb, #003388); color: white;}
    
    /* リセットボタン（赤系）のスタイル定義用のクラスなどはStreamlit標準では難しいので配置で工夫 */
    
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {
        background-color: #e6f2ff; border-bottom-color: #0066cc; color: #0066cc; font-weight: bold;
    }
    section[data-testid="stSidebar"] { background-color: #eef4f9; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. データベース定義
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

# ▼ ハンドル単価マスタ
HANDLE_PRICES_SINGLE = {
    ("cosmo", "std"):   {(False, False): 115, (False, True): 185, (True, False): 185, (True, True): 255},
    ("advance", "std"): {(False, False): 230, (False, True): 300, (True, False): 300, (True, True): 370},
    ("advance", "black"):{(False, False): 330, (False, True): 400, (True, False): 400, (True, True): 440},
    ("fullcolor", "std"):{(False, False): 0, (False, True): 20, (True, False): 50, (True, True): 70},
    ("sostyle", "std"):  {(False, False): 450, (False, True): 450, (True, False): 450, (True, True): 450},
    ("sostyle", "black"):{(False, False): 450, (False, True): 450, (True, False): 450, (True, True): 450},
    ("other", "std"):    {(False, False): 0, (False, True): 0, (True, False): 0, (True, True): 0},
}

# ▼ プレート単価マスタ (1連)
PLATE_PRICES_1 = {
    ("cosmo", "std"): 170,     ("advance", "std"): 330,   ("advance", "black"): 430,
    ("sostyle", "std"): 900,   ("sostyle", "black"): 900, ("fullcolor", "std"): 220,
    ("jimbo", "std"): 600,
}

# ▼ 部材マスタ
ITEMS_DB = {
    "sw_b_mech": {"name": "片切スイッチ", "icon": "🔘", "img_file": "sw_b.jpg", "has_lamp": False,
                  "fullcolor": 250, "cosmo": 250, "advance": 250, "sostyle": 900, "jimbo": 1800},
    "sw_h_mech": {"name": "ほたるスイッチ", "icon": "🟢", "img_file": "sw_b.jpg", "has_lamp": True,
                  "fullcolor": 630, "cosmo": 550, "advance": 550, "sostyle": 1970, "jimbo": 2900},
    "sw_3_mech": {"name": "3路スイッチ", "icon": "🔄", "img_file": "sw_3.jpg", "has_lamp": False,
                  "fullcolor": 430, "cosmo": 420, "advance": 420, "sostyle": 1500, "jimbo": 2200},
    "sw_3h_mech": {"name": "3路ほたるSW", "icon": "🔄🟢", "img_file": "sw_3.jpg", "has_lamp": True,
                   "fullcolor": 850, "cosmo": 760, "advance": 700, "sostyle": 2900, "jimbo": 3300},
    "sw_4_mech": {"name": "4路スイッチ", "icon": "🔀", "img_file": "sw_4.jpg", "has_lamp": False,
                  "fullcolor": 1600, "cosmo": 1600, "advance": 1600, "sostyle": 3500, "jimbo": 3200},
    "sw_4h_mech": {"name": "4路ほたるSW", "icon": "🔀🟢", "img_file": "sw_4.jpg", "has_lamp": True,
                   "fullcolor": 2100, "cosmo": 1800, "advance": 1600, "sostyle": 5300, "jimbo": 4200},
    "outlet_w": {"name": "ダブルコンセント", "icon": "🔌", "img_file": "outlet_w.jpg", "has_lamp": False,
                 "fullcolor": 380, "cosmo": 550, "advance": 800, "sostyle": 1200, "jimbo": 1300},
    "outlet_e": {"name": "アース付コンセント", "icon": "⏚", "img_file": "outlet_e.jpg", "has_lamp": False,
                 "fullcolor": 450, "cosmo": 600, "advance": 900, "sostyle": 1300, "jimbo": 1500},
    "tv_4k": {"name": "TV端子(4K8K)", "icon": "📺", "img_file": "tv_4k.jpg", "has_lamp": False,
              "fullcolor": 1400, "cosmo": 1400, "advance": 1700, "sostyle": 2100, "jimbo": 2300},
    "lan_6": {"name": "LAN(CAT6)", "icon": "💻", "img_file": "lan_6.jpg", "has_lamp": False,
              "fullcolor": 2090, "cosmo": 2090, "advance": 2500, "sostyle": 3500, "jimbo": 3200},
}

FRAME_PRICES = {"fullcolor": 60, "cosmo": 70, "advance": 70, "sostyle": 150, "jimbo": 100}

# ==========================================
# 2. 関数ロジック
# ==========================================
if 'estimate_list' not in st.session_state:
    st.session_state.estimate_list = []

def show_item_image(item_key):
    if item_key in ITEMS_DB:
        item_data = ITEMS_DB[item_key]
        img_filename = item_data.get("img_file", "")
        img_path = os.path.join("img", img_filename)
        if os.path.exists(img_path):
            st.image(Image.open(img_path), use_column_width=True)
        else:
            st.markdown(f"<h1 style='text-align: center; color: #ccc;'>{item_data['icon']}</h1>", unsafe_allow_html=True)

def get_db_price(db, series_key, color_type, *args):
    if (series_key, color_type) in db: val = db[(series_key, color_type)]
    elif (series_key, "std") in db: val = db[(series_key, "std")]
    else: return db.get(series_key, 0)
    if args and isinstance(val, dict): return val.get(args[0], 0)
    return val

def calculate_single_unit(item_key, src_series, tgt_series, tgt_color, needs_window, needs_name, handle_type="single"):
    item = ITEMS_DB[item_key]
    p_body_src = item.get(src_series, 0)
    p_body_tgt = item.get(tgt_series, 0)
    p_frame_src = FRAME_PRICES.get(src_series, 0)
    p_frame_tgt = FRAME_PRICES.get(tgt_series, 0)
    p_plate_src = get_db_price(PLATE_PRICES_1, src_series, "std")
    p_plate_tgt = get_db_price(PLATE_PRICES_1, tgt_series, tgt_color)
    
    if "outlet" in item_key or "tv" in item_key or "lan" in item_key:
        p_handle_src = 0; p_handle_tgt = 0
    else:
        h_key = (needs_window, needs_name)
        p_h_src = get_db_price(HANDLE_PRICES_SINGLE, src_series, "std", h_key)
        p_h_tgt = get_db_price(HANDLE_PRICES_SINGLE, tgt_series, tgt_color, h_key)
        if handle_type == "double":
            adder_src = 110 if src_series == "cosmo" else 320
            adder_tgt = 110 if tgt_series == "cosmo" else 320
            p_h_src += adder_src; p_h_tgt += adder_tgt
        elif handle_type == "triple":
            adder_src = 220 if src_series == "cosmo" else 640
            adder_tgt = 220 if tgt_series == "cosmo" else 640
            p_h_src += adder_src; p_h_tgt += adder_tgt
        p_handle_src = p_h_src; p_handle_tgt = p_h_tgt

    total_src = p_body_src + p_frame_src + p_plate_src + p_handle_src
    total_tgt = p_body_tgt + p_frame_tgt + p_plate_tgt + p_handle_tgt
    return total_tgt - total_src

# ==========================================
# 3. UI - サイドバー
# ==========================================
st.sidebar.header("🏠 物件情報")
client_name = st.sidebar.text_input("施主名")
hm_name = st.sidebar.text_input("HM名")
st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ 設定")
source_series_key = st.sidebar.selectbox("【現在】変更元", list(SERIES_NAMES.keys()), index=1, format_func=lambda x: SERIES_NAMES[x])
target_series_key = st.sidebar.selectbox("【変更】変更先", list(SERIES_NAMES.keys()), index=2, format_func=lambda x: SERIES_NAMES[x])
target_color_mode = "std"
if target_series_key in ["advance", "sostyle"]:
    color_opt = st.sidebar.radio(f"{SERIES_NAMES[target_series_key]}の色", ["標準色 (白・グレー等)", "マットブラック (黒)"], index=0)
    if "ブラック" in color_opt: target_color_mode = "black"

# ==========================================
# 4. メイン画面
# ==========================================
st.title("⚡ 電材差額見積りアプリ Pro")
st.info(f"計算モード： {SERIES_NAMES[source_series_key]} ➡ {SERIES_NAMES[target_series_key]} ({'黒' if target_color_mode=='black' else '標準色'})")

tab1, tab2, tab3 = st.tabs(["📝 基本(1連)クイック", "🏗️ 多連・詳細ビルダー", "📄 見積書発行"])

# ------------------------------------------
# TAB 1: 簡易入力（リセット機能追加）
# ------------------------------------------
with tab1:
    st.markdown("### 基本スイッチ・コンセント入力")
    is_name_req_simple = st.checkbox("📛 すべて「ネーム付」にする（+差額）", value=False)
    
    # リセット用のコールバック関数
    def clear_inputs():
        keys_to_reset = ["q_sw_b", "q_sw_h", "q_out_w", "q_sw_3", "q_sw_3h", "q_sw_4"]
        for k in keys_to_reset:
            st.session_state[k] = 0

    col1, col2 = st.columns(2)
    # keyを指定することで、プログラムから値を操作できるようにする
    with col1:
        qty_sw_b = st.number_input("片切スイッチ", min_value=0, key="q_sw_b")
        qty_sw_h = st.number_input("ほたるスイッチ", min_value=0, key="q_sw_h")
        qty_out_w = st.number_input("ダブルコンセント", min_value=0, key="q_out_w")
    with col2:
        qty_sw_3 = st.number_input("3路スイッチ", min_value=0, key="q_sw_3")
        qty_sw_3h = st.number_input("3路ほたるスイッチ", min_value=0, key="q_sw_3h")
        qty_sw_4 = st.number_input("4路スイッチ", min_value=0, key="q_sw_4")

    st.markdown("---")
    c_btn1, c_btn2 = st.columns([1, 1])
    with c_btn1:
        if st.button("STEP1 追加", key="btn_simple"):
            def add_simple(item_key, qty):
                if qty > 0:
                    item = ITEMS_DB[item_key]
                    needs_window = item.get("has_lamp", False)
                    diff = calculate_single_unit(item_key, source_series_key, target_series_key, target_color_mode, needs_window, is_name_req_simple)
                    detail_txt = "標準セット"
                    if needs_window: detail_txt += "(表示付)"
                    if is_name_req_simple: detail_txt += "(ネーム付)"
                    st.session_state.estimate_list.append({
                        "type": "1連(基本)", "name": item['name'], "detail": detail_txt,
                        "unit_diff": diff, "qty": qty, "total_diff": diff * qty
                    })
            add_simple("sw_b_mech", qty_sw_b); add_simple("sw_h_mech", qty_sw_h)
            add_simple("sw_3_mech", qty_sw_3); add_simple("sw_3h_mech", qty_sw_3h)
            add_simple("sw_4_mech", qty_sw_4); add_simple("outlet_w", qty_out_w) 
            st.success("追加しました！")
    
    with c_btn2:
        # 入力値クリアボタン
        st.button("🗑️ 入力値を「0」にリセット", on_click=clear_inputs)

# ------------------------------------------
# TAB 2: 詳細ビルダー
# ------------------------------------------
with tab2:
    st.markdown("### 詳細ビルダー：画像確認モード")
    plate_size = st.radio("プレートサイズ", ["1連", "2連", "3連"], horizontal=True)
    cols_num = {"1連":1, "2連":2, "3連":3}[plate_size]
    st.markdown("---")
    ui_cols = st.columns(cols_num)
    column_configs = []
    
    for i in range(cols_num):
        with ui_cols[i]:
            st.markdown(f"**【{i+1}列目】**")
            layout_type = st.selectbox("割り付け", ["シングル(1個)", "ダブル(2個)", "トリプル(3個)", "コンセント(一体)"], key=f"layout_{i}")
            is_name_col = False
            if layout_type != "コンセント(一体)": is_name_col = st.checkbox("📛 ネーム付", key=f"name_opt_{i}")
            items_in_col = []
            opt_list = list(ITEMS_DB.keys())
            def item_selector(label, k):
                c_in, c_im = st.columns([3, 1])
                with c_in: sel = st.selectbox(label, opt_list, format_func=lambda x: ITEMS_DB[x]['name'], key=k)
                with c_im: show_item_image(sel)
                return sel

            if "シングル" in layout_type:
                items_in_col.append(item_selector("中身", f"c{i}_1")); h_type = "single"
            elif "ダブル" in layout_type:
                items_in_col.append(item_selector("上段", f"c{i}_1")); items_in_col.append(item_selector("下段", f"c{i}_2")); h_type = "double"
            elif "トリプル" in layout_type:
                items_in_col.append(item_selector("上段", f"c{i}_1")); items_in_col.append(item_selector("中段", f"c{i}_2")); items_in_col.append(item_selector("下段", f"c{i}_3")); h_type = "triple"
            else:
                items_in_col.append(item_selector("種別", f"c{i}_1")); h_type = "single"
            column_configs.append({"items": items_in_col, "handle": h_type, "is_name": is_name_col})

    st.markdown("---")
    qty_build = st.number_input("この構成のセット数", min_value=1, value=1)
    
    if st.button("見積に追加", key="add_build"):
        p_unit_src = get_db_price(PLATE_PRICES_1, source_series_key, "std")
        p_unit_tgt = get_db_price(PLATE_PRICES_1, target_series_key, target_color_mode)
        plate_factor = 1.0 if cols_num == 1 else (1.8 if cols_num == 2 else 2.6)
        if target_series_key == "cosmo": plate_factor = cols_num * 1.5
        diff_plate = (p_unit_tgt - p_unit_src) * plate_factor
        total_unit_diff = diff_plate
        details_str = []
        for idx, config in enumerate(column_configs):
            d_body = sum([ITEMS_DB[itm].get(target_series_key,0) - ITEMS_DB[itm].get(source_series_key,0) for itm in config['items']])
            if "outlet" in str(column_configs[0]['items'][0]) or "コンセント" in str(config['handle']): d_handle = 0
            else:
                needs_window = any(ITEMS_DB[itm].get("has_lamp", False) for itm in config['items'])
                h_key = (needs_window, config['is_name'])
                p_h_src = get_db_price(HANDLE_PRICES_SINGLE, source_series_key, "std", h_key)
                p_h_tgt = get_db_price(HANDLE_PRICES_SINGLE, target_series_key, target_color_mode, h_key)
                d_handle = p_h_tgt - p_h_src
            d_frame = FRAME_PRICES.get(target_series_key,0) - FRAME_PRICES.get(source_series_key,0)
            total_unit_diff += (d_body + d_handle + d_frame)
            item_names = [ITEMS_DB[itm]['name'] for itm in config['items']]
            details_str.append(f"[{config['handle']}]{','.join(item_names)}")

        st.session_state.estimate_list.append({
            "type": f"{plate_size}カスタム", "name": "詳細構成セット", "detail": " / ".join(details_str),
            "unit_diff": total_unit_diff, "qty": qty_build, "total_diff": total_unit_diff * qty_build
        })
        st.success("追加しました！")

# ------------------------------------------
# TAB 3: 見積書発行
# ------------------------------------------
with tab3:
    st.markdown("### 見積りプレビュー")
    if st.session_state.estimate_list:
        df = pd.DataFrame(st.session_state.estimate_list)
        st.dataframe(df[["type", "name", "detail", "unit_diff", "qty", "total_diff"]], use_container_width=True)
        grand_total = df["total_diff"].sum()
        st.metric("総計(税抜)", f"¥ {grand_total:,.0f}")
        
        def to_excel(df, client, hm, src, tgt, total):
            output = io.BytesIO()
            wb = xlsxwriter.Workbook(output, {'in_memory': True})
            ws = wb.add_worksheet("差額見積")
            fmt_head = wb.add_format({'bold': True, 'bg_color': '#cceeff', 'border': 1})
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
        
        if st.button("見積リストを全消去", key="btn_reset"):
            st.session_state.estimate_list = []
            st.rerun()
