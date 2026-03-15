"""
冷卻水塔水蒸發量計算與分析系統
主程式 - Streamlit 介面（繁體中文版）
"""
import streamlit as st
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import io
import os
from datetime import datetime

from calculations import (
    calculate_water_losses, simple_evaporation,
    enthalpy_saturated, humidity_ratio_saturated,
    saturation_pressure, sensitivity_analysis
)
from export_pdf import generate_pdf
from export_docx import generate_docx
from export_xlsx import generate_xlsx

# ===== 頁面設定 =====
st.set_page_config(
    page_title="冷卻水塔蒸發量計算系統",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===== 自訂 CSS =====
st.markdown("""
<style>
    .main-title {
        font-size: 2rem; font-weight: 700; color: #1a5276;
        text-align: center; padding: 1rem 0;
        border-bottom: 3px solid #1a5276; margin-bottom: 1.5rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #1a5276 0%, #2c3e50 100%);
        padding: 1.2rem; border-radius: 10px; color: white;
        text-align: center; margin-bottom: 0.5rem;
    }
    .metric-card h3 { font-size: 0.85rem; opacity: 0.9; margin: 0; }
    .metric-card h2 { font-size: 1.6rem; margin: 0.3rem 0; }
    .metric-card p { font-size: 0.75rem; opacity: 0.7; margin: 0; }
    .section-header {
        color: #1a5276; font-size: 1.3rem; font-weight: 600;
        border-left: 4px solid #1a5276; padding-left: 0.8rem;
        margin: 1.5rem 0 1rem 0;
    }
    .stDownloadButton>button { width: 100%; border-radius: 8px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


# ===== 中文字型設定 =====
def setup_chinese_font():
    """設定 matplotlib 中文字型"""
    font_candidates = [
        # Windows
        'C:/Windows/Fonts/msjh.ttc',
        'C:/Windows/Fonts/mingliu.ttc',
        'C:/Windows/Fonts/kaiu.ttf',
        'C:/Windows/Fonts/simsun.ttc',
        # Linux
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
        # macOS
        '/System/Library/Fonts/PingFang.ttc',
        '/Library/Fonts/Arial Unicode.ttf',
    ]
    for fp in font_candidates:
        if os.path.exists(fp):
            try:
                font_prop = fm.FontProperties(fname=fp)
                # 全域設定
                plt.rcParams['font.family'] = font_prop.get_name()
                plt.rcParams['axes.unicode_minus'] = False
                return font_prop
            except Exception:
                continue
    # fallback
    plt.rcParams['axes.unicode_minus'] = False
    return fm.FontProperties()


font_prop = setup_chinese_font()


def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    buf.seek(0)
    return buf.getvalue()


def generate_chart_conclusions(results, params):
    """根據計算結果，為每張圖表產生結語"""
    md = results['merkel_details']
    COC = params['COC']
    conclusions = {}

    # 圖1：水損失佔比
    c1 = f"蒸發損失佔總補給水量的 {results['evap_pct']:.1f}%，為主要水損失來源。"
    if results['blowdown_pct'] > 30:
        c1 += f"排污損失佔 {results['blowdown_pct']:.1f}%，比例偏高，建議提高濃縮倍數以降低排污量。"
    else:
        c1 += f"排污損失佔 {results['blowdown_pct']:.1f}%，在合理範圍內。"
    c1 += f"飛濺損失佔 {results['splash_pct']:.1f}%，屬正常範圍。"
    conclusions['water_loss_pie'] = c1

    # 圖2：冷卻曲線
    c2 = f"冷卻範圍為 {md['T_range']:.1f}°C，逼近溫度為 {md['T_approach']:.1f}°C，冷卻效率 {md['efficiency']:.1f}%。"
    if md['efficiency'] > 70:
        c2 += "飽和焓與空氣焓之間的驅動力充足，水塔熱交換效果良好。"
    elif md['efficiency'] > 50:
        c2 += "焓差驅動力尚可，但仍有優化空間，可考慮增加填料面積或風量。"
    else:
        c2 += "焓差驅動力不足，建議檢查填料狀態、風扇運轉效率及水分配是否均勻。"
    conclusions['cooling_curve'] = c2

    # 圖3：COC vs 補給水量
    bd_at_coc3 = results['E_merkel'] / (3 - 1) if COC != 3 else results['E_blowdown']
    bd_at_coc6 = results['E_merkel'] / (6 - 1)
    save_pct = (1 - (results['E_merkel'] + results['E_splash'] + bd_at_coc6) /
                (results['E_merkel'] + results['E_splash'] + bd_at_coc3)) * 100
    c3 = f"目前濃縮倍數為 {COC:.1f}。"
    if COC < 3:
        c3 += f"若提升至 COC=6，總補給水量可減少約 {save_pct:.1f}%。建議在水質允許的條件下提高濃縮倍數以節約用水。"
    elif COC < 5:
        c3 += f"濃縮倍數在合理範圍，若進一步提高至 6 可再節水約 {save_pct:.1f}%，但需注意結垢與腐蝕風險。"
    else:
        c3 += "濃縮倍數已較高，繼續提升節水效益遞減，需特別注意水質管理避免結垢。"
    conclusions['coc_trend'] = c3

    # 圖4：溫度-焓值關係
    c4 = f"工作溫度範圍為 {params['T_out']:.1f}~{params['T_in']:.1f}°C。"
    c4 += "在此區間內，飽和空氣焓值隨溫度呈指數上升，表示高溫段的蒸發驅動力顯著大於低溫段。"
    if md['T_range'] > 8:
        c4 += "溫差較大，水塔承擔的熱負荷高，需確保足夠的風量與填料面積。"
    else:
        c4 += "溫差適中，水塔熱負荷在正常範圍。"
    conclusions['temp_enthalpy'] = c4

    return conclusions


# ===== 側邊欄 - 輸入參數 =====
with st.sidebar:
    st.markdown("## 📋 專案資訊")
    project_name = st.text_input("專案名稱", value="冷卻水塔分析報告")
    engineer = st.text_input("計算人員", value="工程師")
    calc_date = st.date_input("日期", value=datetime.now())

    st.markdown("---")
    st.markdown("## 🌊 運轉參數")
    Q = st.number_input("循環水量 Q (m³/h)", min_value=1.0, max_value=50000.0, value=500.0, step=50.0)
    T_in = st.number_input("進水溫度 T_in (°C)", min_value=10.0, max_value=60.0, value=37.0, step=0.5)
    T_out = st.number_input("出水溫度 T_out (°C)", min_value=5.0, max_value=50.0, value=32.0, step=0.5)

    st.markdown("## 🌡️ 環境條件")
    Tdb = st.number_input("乾球溫度 Tdb (°C)", min_value=-10.0, max_value=50.0, value=35.0, step=0.5)
    Twb = st.number_input("濕球溫度 Twb (°C)", min_value=-10.0, max_value=45.0, value=28.0, step=0.5)

    st.markdown("## 💨 風量與填料")
    G = st.number_input("空氣流量 G (m³/h)", min_value=100.0, max_value=500000.0, value=400000.0, step=10000.0)
    KaV_L = st.number_input("填料特性係數 KaV/L", min_value=0.1, max_value=5.0, value=1.5, step=0.1)

    st.markdown("## ⚙️ 水塔設定")
    tower_type = st.selectbox("水塔類型", ['counter', 'cross'],
        format_func=lambda x: '逆流式 (Counter-flow)' if x == 'counter' else '橫流式 (Cross-flow)')
    season = st.selectbox("季節", ['summer', 'winter'],
        format_func=lambda x: '夏季 (C=1.0)' if x == 'summer' else '冬季 (C=0.8)')
    COC = st.number_input("濃縮倍數 (COC)", min_value=1.1, max_value=15.0, value=3.0, step=0.5)

if T_in <= T_out:
    st.error("⚠️ 進水溫度必須大於出水溫度！")
    st.stop()
if Twb >= Tdb:
    st.error("⚠️ 濕球溫度必須小於乾球溫度！")
    st.stop()

# ===== 計算 =====
params = dict(Q=Q, T_in=T_in, T_out=T_out, G=G, tower_type=tower_type,
              season=season, COC=COC, KaV_L=KaV_L, Tdb=Tdb, Twb=Twb, P=101.325,
              project_name=project_name, engineer=engineer,
              date=calc_date.strftime('%Y-%m-%d'))

results = calculate_water_losses(**{k: v for k, v in params.items()
                                    if k not in ('project_name', 'engineer', 'date')})
md = results['merkel_details']
chart_conclusions = generate_chart_conclusions(results, params)

# ===== 主畫面 =====
st.markdown('<div class="main-title">💧 冷卻水塔水蒸發量計算與分析系統</div>', unsafe_allow_html=True)

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.markdown(f'<div class="metric-card"><h3>蒸發水量 (Merkel)</h3><h2>{results["E_merkel"]:.4f}</h2><p>m³/h</p></div>', unsafe_allow_html=True)
with col2:
    st.markdown(f'<div class="metric-card"><h3>總補給水量</h3><h2>{results["E_total"]:.4f}</h2><p>m³/h</p></div>', unsafe_allow_html=True)
with col3:
    st.markdown(f'<div class="metric-card"><h3>冷卻效率</h3><h2>{md["efficiency"]:.1f}%</h2><p></p></div>', unsafe_allow_html=True)
with col4:
    st.markdown(f'<div class="metric-card"><h3>氣水比 (L/G)</h3><h2>{md["LG_ratio"]:.3f}</h2><p></p></div>', unsafe_allow_html=True)
with col5:
    st.markdown(f'<div class="metric-card"><h3>每日補水量</h3><h2>{results["E_total"]*24:.1f}</h2><p>m³/日</p></div>', unsafe_allow_html=True)

st.markdown("")

# ===== 計算結果比較表 =====
st.markdown('<div class="section-header">📊 計算結果比較</div>', unsafe_allow_html=True)
col_a, col_b = st.columns(2)

with col_a:
    st.markdown("##### 簡易公式 vs 焓差法")
    comparison_data = {
        '計算方法': ['簡易經驗公式', 'Merkel 焓差法', '差異比率'],
        '蒸發水量 (m³/h)': [
            f"{results['E_simple']:.4f}", f"{results['E_merkel']:.4f}",
            f"{(results['E_merkel']/results['E_simple']*100 - 100):.1f}%" if results['E_simple'] > 0 else '-'
        ],
    }
    st.table(comparison_data)

with col_b:
    st.markdown("##### 水損失明細")
    loss_data = {
        '項目': ['蒸發損失', '飛濺損失', '排污損失', '**總補給水量**'],
        '流量 (m³/h)': [f"{results['E_merkel']:.4f}", f"{results['E_splash']:.4f}",
                        f"{results['E_blowdown']:.4f}", f"**{results['E_total']:.4f}**"],
        '佔比 (%)': [f"{results['evap_pct']:.1f}", f"{results['splash_pct']:.1f}",
                     f"{results['blowdown_pct']:.1f}", "100.0"],
    }
    st.table(loss_data)


# ===== 圖表（全部中文）=====
st.markdown('<div class="section-header">📈 圖表分析</div>', unsafe_allow_html=True)
chart_images = {}

fig_col1, fig_col2 = st.columns(2)

# 圖1：水損失佔比圓餅圖
with fig_col1:
    fig1, ax1 = plt.subplots(figsize=(6, 5))
    labels = ['蒸發損失', '飛濺損失', '排污損失']
    sizes = [results['evap_pct'], results['splash_pct'], results['blowdown_pct']]
    colors = ['#3498db', '#e74c3c', '#f39c12']
    explode = (0.05, 0.05, 0.05)
    wedges, texts, autotexts = ax1.pie(
        sizes, explode=explode, labels=labels, colors=colors,
        autopct='%1.1f%%', shadow=True, startangle=90,
        textprops={'fontsize': 10, 'fontproperties': font_prop}
    )
    for t in autotexts:
        t.set_fontproperties(font_prop)
    ax1.set_title('水損失分布比例', fontsize=14, fontweight='bold',
                  color='#1a5276', pad=15, fontproperties=font_prop)
    st.pyplot(fig1)
    chart_images['water_loss_pie'] = fig_to_bytes(fig1)
    plt.close(fig1)

# 圖2：冷卻曲線
with fig_col2:
    fig2, ax2 = plt.subplots(figsize=(6, 5))
    temps = md['temperatures']
    h_sat = md['enthalpies_sat']
    h_air = md['enthalpies_air']
    ax2.plot(temps, h_sat, 'r-o', label='飽和空氣焓值 (h_s)', linewidth=2, markersize=3)
    ax2.plot(temps, h_air, 'b-s', label='空氣焓值 (h_a)', linewidth=2, markersize=3)
    ax2.fill_between(temps, h_air, h_sat, alpha=0.15, color='green')
    ax2.set_xlabel('水溫 (°C)', fontsize=11, fontproperties=font_prop)
    ax2.set_ylabel('焓值 (kJ/kg)', fontsize=11, fontproperties=font_prop)
    ax2.set_title('冷卻曲線（焓值 vs 溫度）', fontsize=14, fontweight='bold',
                  color='#1a5276', fontproperties=font_prop)
    ax2.legend(fontsize=9, prop=font_prop)
    ax2.grid(True, alpha=0.3)
    st.pyplot(fig2)
    chart_images['cooling_curve'] = fig_to_bytes(fig2)
    plt.close(fig2)

fig_col3, fig_col4 = st.columns(2)

# 圖3：COC vs 補給水量
with fig_col3:
    fig3, ax3 = plt.subplots(figsize=(6, 5))
    coc_vals = np.arange(1.5, 10.5, 0.5)
    E_evap = results['E_merkel']
    totals = []
    blowdowns = []
    for c in coc_vals:
        bd = E_evap / (c - 1)
        blowdowns.append(bd)
        totals.append(E_evap + results['E_splash'] + bd)

    ax3.plot(coc_vals, totals, 'b-o', label='總補給水量', linewidth=2, markersize=4)
    ax3.plot(coc_vals, blowdowns, 'r--s', label='排污損失', linewidth=2, markersize=4)
    ax3.axhline(y=E_evap, color='green', linestyle=':', label=f'蒸發量={E_evap:.3f}', alpha=0.7)
    ax3.axvline(x=COC, color='gray', linestyle='--', alpha=0.5, label=f'目前 COC={COC:.1f}')
    ax3.set_xlabel('濃縮倍數 (COC)', fontsize=11, fontproperties=font_prop)
    ax3.set_ylabel('水量 (m³/h)', fontsize=11, fontproperties=font_prop)
    ax3.set_title('濃縮倍數 vs 補給水量', fontsize=14, fontweight='bold',
                  color='#1a5276', fontproperties=font_prop)
    ax3.legend(fontsize=9, prop=font_prop)
    ax3.grid(True, alpha=0.3)
    st.pyplot(fig3)
    chart_images['coc_trend'] = fig_to_bytes(fig3)
    plt.close(fig3)

# 圖4：溫度-焓值關係
with fig_col4:
    fig4, ax4 = plt.subplots(figsize=(6, 5))
    t_range = np.arange(10, 50, 0.5)
    h_sat_curve = [enthalpy_saturated(t) for t in t_range]
    ax4.plot(t_range, h_sat_curve, 'r-', label='飽和空氣焓值曲線', linewidth=2)
    ax4.axvspan(T_out, T_in, alpha=0.15, color='blue', label=f'工作範圍 ({T_out}-{T_in}°C)')
    ax4.set_xlabel('溫度 (°C)', fontsize=11, fontproperties=font_prop)
    ax4.set_ylabel('焓值 (kJ/kg 乾空氣)', fontsize=11, fontproperties=font_prop)
    ax4.set_title('溫度 - 焓值關係圖', fontsize=14, fontweight='bold',
                  color='#1a5276', fontproperties=font_prop)
    ax4.legend(fontsize=9, prop=font_prop)
    ax4.grid(True, alpha=0.3)
    st.pyplot(fig4)
    chart_images['temp_enthalpy'] = fig_to_bytes(fig4)
    plt.close(fig4)


# ===== 詳細數據表 =====
st.markdown('<div class="section-header">📋 詳細計算數據</div>', unsafe_allow_html=True)

with st.expander("🌡️ 環境空氣性質", expanded=False):
    st.table({
        '性質': ['溫度', '比濕度', '焓值', '相對濕度'],
        '入口': [f"{Tdb:.1f} °C", f"{results['W_inlet']:.5f} kg/kg",
                 f"{results['h_inlet']:.2f} kJ/kg", f"{results['RH']:.1f}%"],
        '出口': [f"{md['T_air_out']:.1f} °C", f"{md['W_air_out']:.5f} kg/kg",
                 f"{md['h_air_out']:.2f} kJ/kg", "-"],
    })

with st.expander("🔧 Merkel 法詳細參數", expanded=False):
    st.table({
        '參數': ['氣水比 (L/G)', 'Merkel 數', '逼近溫度', '冷卻範圍', '冷卻效率'],
        '數值': [f"{md['LG_ratio']:.3f}", f"{md['merkel_number']:.3f}",
                 f"{md['T_approach']:.1f} °C", f"{md['T_range']:.1f} °C",
                 f"{md['efficiency']:.1f}%"],
    })


# ===== 匯出功能 =====
st.markdown('<div class="section-header">📥 匯出報告</div>', unsafe_allow_html=True)
exp_col1, exp_col2, exp_col3 = st.columns(3)

with exp_col1:
    st.markdown("##### 📄 PDF 分析報告")
    st.caption("含封面、表格、圖表、圖表結語、結論建議")
    try:
        pdf_bytes = generate_pdf(results, params, chart_images, chart_conclusions)
        st.download_button("⬇️ 下載 PDF 報告", data=pdf_bytes,
            file_name=f"冷卻水塔報告_{calc_date.strftime('%Y%m%d')}.pdf",
            mime="application/pdf", use_container_width=True)
    except Exception as e:
        st.error(f"PDF 產生失敗: {e}")

with exp_col2:
    st.markdown("##### 📝 Word 報告文件")
    st.caption("含目錄、公式說明、圖表結語、附錄")
    try:
        docx_bytes = generate_docx(results, params, chart_images, chart_conclusions)
        st.download_button("⬇️ 下載 Word 文件", data=docx_bytes,
            file_name=f"冷卻水塔報告_{calc_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    except Exception as e:
        st.error(f"Word 產生失敗: {e}")

with exp_col3:
    st.markdown("##### 📊 Excel 計算表")
    st.caption("含逐步公式、敏感度分析、可編輯")
    try:
        xlsx_bytes = generate_xlsx(results, params)
        st.download_button("⬇️ 下載 Excel 檔案", data=xlsx_bytes,
            file_name=f"冷卻水塔計算表_{calc_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    except Exception as e:
        st.error(f"Excel 產生失敗: {e}")

st.markdown("---")
st.markdown(
    '<div style="text-align:center; color:#888; font-size:0.8rem;">'
    '冷卻水塔蒸發量計算系統 v1.0 ｜ 計算方法：簡易經驗公式 & Merkel 焓差法'
    '</div>', unsafe_allow_html=True
)
