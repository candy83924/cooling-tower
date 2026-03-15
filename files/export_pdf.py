"""
PDF 分析報告匯出模組（繁體中文版）
含圖表結語分析
"""
import io
import os
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def register_chinese_font():
    font_paths = [
        'C:/Windows/Fonts/msjh.ttc', 'C:/Windows/Fonts/mingliu.ttc',
        'C:/Windows/Fonts/kaiu.ttf', 'C:/Windows/Fonts/simsun.ttc',
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('ChineseFont', fp))
                return 'ChineseFont'
            except Exception:
                continue
    return 'Helvetica'


def get_styles(font_name):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CoverTitle', fontName=font_name, fontSize=28,
        leading=40, alignment=TA_CENTER, spaceAfter=20, textColor=HexColor('#1a5276')))
    styles.add(ParagraphStyle(name='SectionTitle', fontName=font_name, fontSize=16,
        leading=22, spaceBefore=16, spaceAfter=10, textColor=HexColor('#1a5276')))
    styles.add(ParagraphStyle(name='SubTitle', fontName=font_name, fontSize=12,
        leading=18, spaceBefore=8, spaceAfter=6, textColor=HexColor('#2c3e50'), bold=True))
    styles.add(ParagraphStyle(name='BodyCN', fontName=font_name, fontSize=10,
        leading=16, spaceAfter=6, alignment=TA_JUSTIFY))
    styles.add(ParagraphStyle(name='ChartNote', fontName=font_name, fontSize=9,
        leading=14, spaceAfter=8, alignment=TA_JUSTIFY, textColor=HexColor('#34495e'),
        leftIndent=10, rightIndent=10, backColor=HexColor('#f8f9fa'),
        borderWidth=0.5, borderColor=HexColor('#dee2e6'), borderPadding=6))
    return styles


def header_footer(canvas, doc, project_name, font_name):
    canvas.saveState()
    canvas.setFont(font_name, 8)
    canvas.setFillColor(HexColor('#1a5276'))
    canvas.drawString(2*cm, A4[1]-1.2*cm, project_name)
    canvas.setStrokeColor(HexColor('#1a5276'))
    canvas.setLineWidth(0.5)
    canvas.line(2*cm, A4[1]-1.4*cm, A4[0]-2*cm, A4[1]-1.4*cm)
    canvas.setFillColor(HexColor('#888888'))
    canvas.drawString(2*cm, 1.2*cm, f"產生日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    canvas.drawRightString(A4[0]-2*cm, 1.2*cm, f"第 {doc.page} 頁")
    canvas.restoreState()


def make_data_table(data_rows, col_widths, font_name):
    table = Table(data_rows, colWidths=col_widths)
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1a5276')),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('FONTNAME', (0, 0), (-1, -1), font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#cccccc')),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]
    for i in range(1, len(data_rows)):
        if i % 2 == 0:
            style_cmds.append(('BACKGROUND', (0, i), (-1, i), HexColor('#eaf2f8')))
    table.setStyle(TableStyle(style_cmds))
    return table


def generate_pdf(results, params, chart_images=None, chart_conclusions=None):
    font_name = register_chinese_font()
    styles = get_styles(font_name)
    if chart_conclusions is None:
        chart_conclusions = {}

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
        topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)

    story = []
    proj = params.get('project_name', '冷卻水塔分析報告')
    engineer = params.get('engineer', '-')
    date_str = params.get('date', datetime.now().strftime('%Y-%m-%d'))

    # ===== 封面 =====
    story.append(Spacer(1, 80))
    story.append(Paragraph("冷卻水塔", styles['CoverTitle']))
    story.append(Paragraph("水蒸發量分析報告", styles['CoverTitle']))
    story.append(Spacer(1, 10))
    line_table = Table([['', '']], colWidths=[A4[0] - 4*cm])
    line_table.setStyle(TableStyle([('LINEBELOW', (0, 0), (-1, 0), 2, HexColor('#1a5276'))]))
    story.append(line_table)
    story.append(Spacer(1, 30))
    cover_items = [['專案名稱', proj], ['計算人員', engineer], ['日期', date_str]]
    cover_table = Table(cover_items, colWidths=[4*cm, 10*cm])
    cover_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name), ('FONTSIZE', (0,0), (-1,-1), 12),
        ('TEXTCOLOR', (0,0), (0,-1), HexColor('#1a5276')),
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('ALIGN', (0,0), (0,-1), 'RIGHT'), ('ALIGN', (1,0), (1,-1), 'LEFT'),
        ('LEFTPADDING', (1,0), (1,-1), 15),
    ]))
    story.append(cover_table)
    story.append(PageBreak())

    # ===== 1. 輸入參數 =====
    story.append(Paragraph("1. 輸入參數", styles['SectionTitle']))
    tower_label = "逆流式" if params.get('tower_type') == 'counter' else "橫流式"
    season_label = "夏季 (C=1.0)" if params.get('season') == 'summer' else "冬季 (C=0.8)"
    param_data = [
        ['參數', '數值', '單位'],
        ['循環水量 (Q)', f"{params['Q']:.1f}", 'm³/h'],
        ['進水溫度 (T_in)', f"{params['T_in']:.1f}", '°C'],
        ['出水溫度 (T_out)', f"{params['T_out']:.1f}", '°C'],
        ['溫差 (ΔT)', f"{params['T_in'] - params['T_out']:.1f}", '°C'],
        ['乾球溫度 (Tdb)', f"{params['Tdb']:.1f}", '°C'],
        ['濕球溫度 (Twb)', f"{params['Twb']:.1f}", '°C'],
        ['空氣流量 (G)', f"{params['G']:.1f}", 'm³/h'],
        ['填料特性係數 (KaV/L)', f"{params['KaV_L']:.2f}", '-'],
        ['濃縮倍數 (COC)', f"{params['COC']:.1f}", '-'],
        ['水塔類型', tower_label, '-'],
        ['季節', season_label, '-'],
    ]
    story.append(make_data_table(param_data, [6*cm, 5*cm, 3*cm], font_name))
    story.append(Spacer(1, 15))

    # ===== 2. 計算結果 =====
    story.append(Paragraph("2. 計算結果", styles['SectionTitle']))
    md = results['merkel_details']
    result_data = [
        ['項目', '數值', '單位'],
        ['氣水比 (L/G)', f"{md['LG_ratio']:.3f}", '-'],
        ['Merkel 數', f"{md['merkel_number']:.3f}", '-'],
        ['蒸發量（簡易公式）', f"{results['E_simple']:.4f}", 'm³/h'],
        ['蒸發量（Merkel 法）', f"{results['E_merkel']:.4f}", 'm³/h'],
        ['飛濺損失', f"{results['E_splash']:.4f}", 'm³/h'],
        ['排污損失', f"{results['E_blowdown']:.4f}", 'm³/h'],
        ['總補給水量', f"{results['E_total']:.4f}", 'm³/h'],
        ['冷卻效率', f"{md['efficiency']:.1f}", '%'],
        ['逼近溫度', f"{md['T_approach']:.1f}", '°C'],
        ['冷卻範圍', f"{md['T_range']:.1f}", '°C'],
    ]
    story.append(make_data_table(result_data, [6*cm, 5*cm, 3*cm], font_name))
    story.append(Spacer(1, 15))

    # ===== 3. 水損失分布 =====
    story.append(Paragraph("3. 水損失分布", styles['SectionTitle']))
    dist_data = [
        ['損失類型', '流量 (m³/h)', '佔比 (%)'],
        ['蒸發損失', f"{results['E_merkel']:.4f}", f"{results['evap_pct']:.1f}"],
        ['飛濺損失', f"{results['E_splash']:.4f}", f"{results['splash_pct']:.1f}"],
        ['排污損失', f"{results['E_blowdown']:.4f}", f"{results['blowdown_pct']:.1f}"],
        ['合計', f"{results['E_total']:.4f}", "100.0"],
    ]
    story.append(make_data_table(dist_data, [5*cm, 5*cm, 4*cm], font_name))
    story.append(Spacer(1, 15))

    # ===== 4. 環境空氣性質 =====
    story.append(Paragraph("4. 環境空氣性質", styles['SectionTitle']))
    air_data = [
        ['性質', '入口', '出口', '單位'],
        ['溫度', f"{params['Tdb']:.1f}", f"{md['T_air_out']:.1f}", '°C'],
        ['比濕度', f"{results['W_inlet']:.5f}", f"{md['W_air_out']:.5f}", 'kg/kg'],
        ['焓值', f"{results['h_inlet']:.2f}", f"{md['h_air_out']:.2f}", 'kJ/kg'],
        ['相對濕度', f"{results['RH']:.1f}", '-', '%'],
    ]
    story.append(make_data_table(air_data, [4.5*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_name))
    story.append(Spacer(1, 15))

    # ===== 5. 圖表分析（含結語）=====
    if chart_images:
        story.append(PageBreak())
        story.append(Paragraph("5. 圖表分析", styles['SectionTitle']))

        chart_titles = {
            'water_loss_pie': '圖 5-1：水損失分布比例',
            'cooling_curve': '圖 5-2：冷卻曲線（焓值 vs 溫度）',
            'coc_trend': '圖 5-3：濃縮倍數 vs 補給水量',
            'temp_enthalpy': '圖 5-4：溫度 - 焓值關係圖',
        }

        for key, title in chart_titles.items():
            if key in chart_images and chart_images[key]:
                try:
                    story.append(Paragraph(title, styles['SubTitle']))
                    img_io = io.BytesIO(chart_images[key])
                    img = Image(img_io, width=15*cm, height=9.5*cm)
                    story.append(img)
                    story.append(Spacer(1, 6))
                    # 圖表結語
                    if key in chart_conclusions:
                        story.append(Paragraph(
                            f"<b>分析結語：</b>{chart_conclusions[key]}",
                            styles['ChartNote']
                        ))
                    story.append(Spacer(1, 12))
                except Exception:
                    pass

    # ===== 6. 結論與建議 =====
    story.append(PageBreak())
    story.append(Paragraph("6. 結論與建議", styles['SectionTitle']))

    daily_makeup = results['E_total'] * 24
    monthly_makeup = daily_makeup * 30
    conclusions = [
        f"總補給水量需求：{results['E_total']:.4f} m³/h（約 {daily_makeup:.1f} m³/日，{monthly_makeup:.0f} m³/月）。",
    ]
    if md['efficiency'] > 70:
        conclusions.append("冷卻效率良好（>70%），水塔運行狀態佳。")
    elif md['efficiency'] > 50:
        conclusions.append("冷卻效率中等（50-70%），建議檢查風量或填料狀況。")
    else:
        conclusions.append("冷卻效率偏低（<50%），請檢查填料、風扇性能及擋水器。")

    if params['COC'] < 3:
        conclusions.append(
            f"目前濃縮倍數為 {params['COC']:.1f}，若提高至 4-6 可減少排污水量"
            f"約 {(results['E_blowdown'] * 0.3):.3f} m³/h，達到節水效果。")

    evap_ratio = results['E_merkel'] / results['E_simple'] if results['E_simple'] > 0 else 1
    conclusions.append(
        f"Merkel 法與簡易公式比值：{evap_ratio:.2f}。"
        f"{'兩種方法結果一致。' if 0.7 < evap_ratio < 1.3 else '偏差較大，請確認輸入參數。'}")

    for c in conclusions:
        story.append(Paragraph(f"• {c}", styles['BodyCN']))
        story.append(Spacer(1, 4))

    def on_page(canvas, doc):
        header_footer(canvas, doc, proj, font_name)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    buffer.seek(0)
    return buffer.getvalue()
