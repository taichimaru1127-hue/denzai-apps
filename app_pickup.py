import streamlit as st
import pandas as pd
from PIL import Image
from streamlit_drawable_canvas import st_canvas

# ==========================================
# 1. アプリ基本設定
# ==========================================
st.set_page_config(page_title="図面拾い出しツール", layout="wide")

# ==========================================
# 2. UI改善：スクロールバー強制表示CSS
# ==========================================
st.markdown("""
    <style>
    /* 1. アプリ全体の横スクロールを許可する設定 */
    .stApp > header {background-color: transparent;}
    .main .block-container {
        max-width: 100%;
        padding-left: 2rem;
        padding-right: 2rem;
        overflow-x: auto !important; /* 強制的にスクロールさせる */
    }

    /* 2. スクロールバー自体のデザイン（太く、見やすく） */
    ::-webkit-scrollbar {
        height: 20px !important; /* バーの高さ(太さ) */
        width: 20px !important;
    }
    ::-webkit-scrollbar-track {
        background: #f0f0f0; 
        border-radius: 10px;
    }
    ::-webkit-scrollbar-thumb {
        background: #888; 
        border-radius: 10px;
        border: 4px solid #f0f0f0; /* 余白を持たせて浮き出るように */
    }
    ::-webkit-scrollbar-thumb:hover {
        background: #555; 
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 3. マーカーの色定義
# ==========================================
PICKUP_ITEMS = {
    "sw_b": {"name": "① 片切スイッチ", "color": "rgba(255, 0, 0, 0.4)"},      # 赤
    "sw_3way": {"name": "② 3路スイッチ", "color": "rgba(0, 0, 255, 0.4)"},   # 青
    "sw_4way": {"name": "③ 4路スイッチ", "color": "rgba(0, 128, 0, 0.4)"},   # 緑
    "outlet": {"name": "④ コンセント類", "color": "rgba(255, 165, 0, 0.4)"}, # オレンジ
    "tv_lan": {"name": "⑤ TV/LAN/TEL", "color": "rgba(128, 0, 128, 0.4)"},   # 紫
}

# ==========================================
# 4. サイドバー
# ==========================================
st.sidebar.header("🛠️ 拾い出し操作パネル")

# ズーム機能
st.sidebar.subheader("🔍 表示設定")
zoom_rate = st.sidebar.slider("図面のズーム倍率", 0.5, 3.0, 1.0, 0.1)

st.sidebar.info("""
**💡 ヒント**
スクロールバーは**画面の一番下**に表示されます。
図面が縦に長い場合は、まず下までスクロールしてバーを確認してください。
""")

st.sidebar.markdown("---")

st.sidebar.subheader("1. 何を数えますか？")
target_item_key = st.sidebar.radio(
    "アイテムを選択",
    list(PICKUP_ITEMS.keys()),
    format_func=lambda x: PICKUP_ITEMS[x]["name"],
    key="target_radio"
)

current_color = PICKUP_ITEMS[target_item_key]["color"]

st.sidebar.markdown(f"""
<div style="background-color: {current_color}; padding: 10px; border-radius: 5px; color: black; font-weight: bold; text-align: center; border: 1px solid #ccc;">
    現在のマーカー色
</div>
""", unsafe_allow_html=True)

stroke_width = st.sidebar.slider("マーカーの大きさ", 5, 40, 20)

# ==========================================
# 5. メイン画面
# ==========================================
st.title("🗺️ 図面デジタル拾い出しツール")

uploaded_file = st.file_uploader("図面画像をアップロード (PNG, JPG)", type=["png", "jpg", "jpeg"])

if uploaded_file:
    image = Image.open(uploaded_file)
    
    # ズーム計算
    base_width = 800
    canvas_width = int(base_width * zoom_rate)
    w, h = image.size
    canvas_height = int(canvas_width * (h / w))

    st.markdown("---")
    st.caption(f"▼ 図面エリア（現在の倍率: {zoom_rate}倍）")
    
    # ここにスクロール可能なコンテナを作成（念のため）
    with st.container():
        # キャンバス設定
        canvas_result = st_canvas(
            fill_color=current_color,
            stroke_color=current_color,
            stroke_width=stroke_width,
            background_image=image,
            update_streamlit=True,
            height=canvas_height,
            width=canvas_width,
            drawing_mode="point",
            display_toolbar=True,
            key=f"canvas_pickup_{zoom_rate}", 
        )

    # ==========================================
    # 6. 集計ロジック
    # ==========================================
    if canvas_result.json_data is not None:
        objects = pd.json_normalize(canvas_result.json_data["objects"])
        
        counts = {key: 0 for key in PICKUP_ITEMS.keys()}
        
        if not objects.empty and "fill" in objects.columns:
            for key, info in PICKUP_ITEMS.items():
                target_color = info["color"]
                match_count = objects[objects["fill"] == target_color].shape[0]
                counts[key] = match_count
        
        # 結果表示
        st.sidebar.markdown("---")
        st.sidebar.header("📊 集計結果")
        results_df = pd.DataFrame([
            {"アイテム": PICKUP_ITEMS[k]["name"], "個数": v} for k, v in counts.items()
        ])
        st.sidebar.dataframe(results_df, hide_index=True, use_container_width=True)
        
        total = results_df["個数"].sum()
        st.sidebar.metric("合計マーク数", f"{total} 個")
        
        csv = results_df.to_csv(index=False).encode('utf-8_sig')
        st.sidebar.download_button(
            "📥 CSVをダウンロード",
            csv,
            "pickup_result.csv",
            "text/csv"
        )

else:
    st.info("👆 画像ファイルをアップロードしてください。")