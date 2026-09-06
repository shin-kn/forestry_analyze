"""日本の林業をテーマにしたサンプルプレゼンテーションを生成する。

デザインは sample.pptx と初版のちょうど中間を狙っている。

sample.pptx から採ったもの:
  - 白地。ベタ塗りの背景は使わず、淡いティントと細罫で面を作る
  - 見出しの上に英字のアイブロウ、下に全幅の細罫
  - グレースケール3段（INK / INK2 / MUTED）で情報の階層を作る
  - 数字と英字は Consolas、和文は Meiryo
  - 画像を使わず、図はすべて図形で組む

初版から残したもの:
  - 大きめの級数と、詰め込みすぎない余白
  - 丸バッジ・角丸カードといった柔らかいモチーフ
  - 森林をテーマにしたグリーンのアクセント

実行:  uv.exe run python build_deck.py
出力:  日本の林業_サンプル.pptx
"""

import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# --- パレット ---------------------------------------------------------------
INK = RGBColor(0x1A, 0x1A, 0x18)      # 本文の主色
INK2 = RGBColor(0x4A, 0x4A, 0x44)     # 副次テキスト
MUTED = RGBColor(0x8A, 0x8A, 0x80)    # キャプション
RULE = RGBColor(0xDD, 0xDD, 0xD6)     # 細罫
RULE2 = RGBColor(0xC8, 0xC8, 0xC0)    # やや濃い罫
BAND = RGBColor(0xF7, 0xF7, 0xF4)     # 中立の帯
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

GREEN = RGBColor(0x2C, 0x7A, 0x46)    # 主アクセント（森林側）
BLUE = RGBColor(0x2A, 0x78, 0xD6)     # 副アクセント（需要側）
AMBER = RGBColor(0xC8, 0x8A, 0x00)    # 注意・ボトルネック
TINT_G = RGBColor(0xE6, 0xF2, 0xEA)
TINT_B = RGBColor(0xE3, 0xEE, 0xFB)
TINT_A = RGBColor(0xFA, 0xF0, 0xDC)
TINT_N = RGBColor(0xEF, 0xEF, 0xEA)

JP = "Meiryo"
MONO = "Consolas"

W, H = 13.333, 7.5
MX = 0.72          # 左右マージン
CW = 11.88         # 本文幅
COL = 3.75         # 3カラムのカード幅
PITCH = 4.06       # 3カラムのピッチ
NOTE = "※ 本資料はサンプルです。数値はデモ用の概数で、特定の統計を出典としたものではありません。"


# --- 低レベルヘルパー --------------------------------------------------------
def style_run(run, size, bold=False, color=INK, font=JP, ea=None):
    """欧文(latin)と和文(ea)の両方にフォントを指定する。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rPr = run.font._rPr
    latin = rPr.find(qn("a:latin"))
    node = rPr.find(qn("a:ea"))
    if node is None:
        node = rPr.makeelement(qn("a:ea"), {})
        latin.addnext(node)
    node.set("typeface", ea or font)


def add_run(paragraph, text, size, bold, color, font=JP, ea=None):
    run = paragraph.add_run()
    run.text = text
    style_run(run, size, bold, color, font, ea)
    return run


def textbox(slide, x, y, w, h, blocks, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            font=JP, ea=None, spacing=1.4):
    """blocks = [(text, size, bold, color, space_after_pt), ...]  改行は \\n"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, (text, size, bold, color, gap) in enumerate(blocks):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(gap)
        p.line_spacing = spacing
        for j, line in enumerate(text.split("\n")):
            if j:
                p.add_line_break()
            add_run(p, line, size, bold, color, font, ea)
    return tb


def box(slide, x, y, w, h, fill, radius=0.05, kind=MSO_SHAPE.ROUNDED_RECTANGLE):
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    if kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    return shp


def rule(slide, x, y, w, color=RULE, width=1.0):
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


def tick(slide, x, y, color, w=0.05, h=0.24):
    """箇条書きの頭に置く小さな縦棒。"""
    return box(slide, x, y, w, h, color, kind=MSO_SHAPE.RECTANGLE)


def badge(slide, x, y, d, label, fill=GREEN, color=WHITE, size=12):
    """初版から引き継いだ丸バッジ。"""
    shp = box(slide, x, y, d, d, fill, kind=MSO_SHAPE.OVAL)
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_run(p, label, size, True, color, MONO, JP)
    return shp


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def frame(slide, eyebrow, title, footer=None):
    """共通の枠: アイブロウ / 見出し / 全幅の細罫 / 脚注。"""
    textbox(slide, MX, 0.42, 10.0, 0.28, [(eyebrow, 11, False, MUTED, 0)],
            font=MONO, ea=JP)
    textbox(slide, MX, 0.74, CW, 0.60, [(title, 28, True, INK, 0)])
    rule(slide, MX, 1.44, CW)
    if footer:
        textbox(slide, MX, 6.88, CW, 0.34, [(footer, 10.5, False, MUTED, 0)])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def style_axes(chart, accent):
    chart.font.size = Pt(11)
    chart.font.name = JP
    chart.font.color.rgb = MUTED
    chart.has_legend = False
    cat, val = chart.category_axis, chart.value_axis
    cat.has_major_gridlines = False
    cat.format.line.color.rgb = RULE
    cat.tick_labels.font.size = Pt(11.5)
    cat.tick_labels.font.color.rgb = INK
    val.has_major_gridlines = True
    val.major_gridlines.format.line.color.rgb = RULE
    val.major_gridlines.format.line.width = Pt(0.75)
    val.format.line.fill.background()
    val.tick_labels.font.size = Pt(10.5)
    val.tick_labels.font.color.rgb = MUTED
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(12)
    dl.font.bold = True
    dl.font.color.rgb = accent
    dl.font.name = MONO
    return plot


def chart_caption(chart, text):
    chart.has_title = True
    tf = chart.chart_title.text_frame
    tf.text = text
    style_run(tf.paragraphs[0].runs[0], 11, False, MUTED)


# --- スライド ----------------------------------------------------------------
def slide_title(prs):
    s = blank(prs)
    box(s, 0, 0, 0.18, H, GREEN, kind=MSO_SHAPE.RECTANGLE)
    textbox(s, 1.10, 2.14, 11.0, 0.90, [("日本の林業のいま", 40, True, INK, 0)])
    textbox(s, 1.13, 3.12, 11.0, 0.34,
            [("forestry in japan - resource, cost, and demand", 13, False, GREEN, 0)],
            font=MONO, ea=JP)
    rule(s, 1.13, 3.72, 4.20, RULE2, 1.2)
    textbox(s, 1.13, 3.96, 10.4, 1.00,
            [("資源はすでに育っている。\n"
              "課題は「伐る・植える・使う」を一本の線につなぎ直すこと。", 15, True, INK2, 0)])
    textbox(s, 1.13, 6.55, 10.4, 0.30,
            [("サンプルプレゼンテーション ／ 2026年9月 ／ 数値はデモ用の概数", 11, False, MUTED, 0)])
    notes(s, "PowerPoint生成のサンプル。戦後に植えた人工林が一斉に主伐期を迎えた、"
             "という一点だけ持ち帰ってもらえれば十分です。")
    return s


def slide_overview(prs):
    s = blank(prs)
    frame(s, "OVERVIEW", "全体像 — 資源は育った。滞っているのは「使う」側")

    steps = [("植える", "再造林", TINT_N), ("育てる", "間伐・保育", TINT_G),
             ("伐る", "主伐・搬出", TINT_G), ("運ぶ", "流通・加工", TINT_B),
             ("使う", "建築・製品", TINT_B)]
    bw, pitch, y = 2.20, 2.42, 1.95
    for i, (head, sub, fill) in enumerate(steps):
        x = MX + i * pitch
        box(s, x, y, bw, 1.00, fill)
        textbox(s, x + 0.22, y + 0.20, bw - 0.44, 0.60,
                [(head, 13.5, True, INK, 3), (sub, 11, False, INK2, 0)])
        if i:
            rule(s, x - 0.19, y + 0.50, 0.16, RULE2)
    textbox(s, MX + 2 * pitch, y + 1.10, 2.6, 0.28,
            [("↑ ここで詰まっている", 11, True, AMBER, 0)])

    cols = [
        (GREEN, "資源量", "人工林の過半が50年生を超え、主伐できる木は十分に育っている。"),
        (AMBER, "コスト", "立木価格が低く、搬出と再造林の費用を伐採収入で賄いにくい。"),
        (BLUE, "需要", "国産材を量で使い切る出口が細く、価格が上がりにくい。"),
    ]
    for i, (accent, head, body) in enumerate(cols):
        x = MX + i * PITCH
        tick(s, x, 3.62, accent)
        textbox(s, x + 0.22, 3.60, 3.3, 0.30, [(head, 13, True, accent, 0)])
        textbox(s, x + 0.22, 4.02, COL - 0.22, 0.80, [(body, 11, False, INK2, 0)])

    box(s, MX, 5.02, CW, 1.52, BAND)
    textbox(s, MX + 0.30, 5.24, 11.3, 0.30, [("設計の芯", 13, True, INK, 0)])
    textbox(s, MX + 0.30, 5.66, 11.3, 0.80,
            [("木が足りないのではない。育った木を伐り出し、跡地に植え直し、製品として使い切る——\n"
              "この循環のどこか一か所でも詰まれば、森林は資源のまま滞留する。",
              12.5, False, INK2, 0)])
    notes(s, "パイプラインの3番目で止まっている、という絵を先に見せる。"
             "以降のスライドはこの図のどこの話かを常に指し示す。")
    return s


def slide_numbers(prs):
    s = blank(prs)
    frame(s, "KEY NUMBERS", "数字で見る、日本の森", NOTE)
    cards = [
        ("2/3", GREEN, TINT_G, "国土に占める森林の割合",
         "森林面積はおよそ2,500万ha。\n先進国のなかでも上位の水準。"),
        ("40%", GREEN, TINT_G, "森林に占める人工林の割合",
         "面積にしておよそ1,000万ha。\nスギ・ヒノキが大半を占める。"),
        ("50+", AMBER, TINT_A, "人工林の過半が達した林齢（年）",
         "多くが主伐期に入り、\n使わなければ立ったまま滞留する。"),
    ]
    for i, (big, accent, fill, label, desc) in enumerate(cards):
        x = MX + i * PITCH
        box(s, x, 1.95, COL, 2.85, fill)
        textbox(s, x + 0.34, 2.26, 3.0, 0.95, [(big, 46, True, accent, 0)],
                font=MONO, ea=JP)
        textbox(s, x + 0.34, 3.30, COL - 0.60, 0.36, [(label, 13, True, INK, 0)])
        textbox(s, x + 0.34, 3.76, COL - 0.60, 0.80, [(desc, 11, False, INK2, 0)])

    rule(s, MX, 5.35, CW)
    textbox(s, MX, 5.58, CW, 0.90,
            [("一斉に植えたものは、一斉に高齢化する。"
              "量も面積もすでに積み上がっていて、いま偏っているのは「齢」の方。\n"
              "「植えて育てる」段階から「伐って使い、また植える」段階へ移る時期にある。",
              12.5, False, INK2, 0)])
    notes(s, "3つの数字だけ覚えてもらう。特に3つ目、林齢50年超が後半の主伐の話につながる。")
    return s


def slide_issues(prs):
    s = blank(prs)
    frame(s, "ISSUES", "林業が直面する3つの課題")
    rows = [
        ("01", "担い手が減っている",
         "林業就業者は長期的に減少し、高齢化も進む。伐採・搬出の技術を継承する相手がいない。"),
        ("02", "所有が小さく、分散している",
         "境界が不明な森林が多く、まとまった面積での施業に持ち込みにくい。集約化に手間がかかる。"),
        ("03", "採算が合いにくい",
         "立木価格が低迷し、伐採収入だけでは再造林のコストを回収しづらい。伐り逃げの誘因が残る。"),
    ]
    for i, (num, head, body) in enumerate(rows):
        y = 1.95 + i * 1.50
        badge(s, MX, y, 0.62, num)
        textbox(s, MX + 0.86, y + 0.02, 6.6, 0.34, [(head, 15, True, INK, 0)])
        textbox(s, MX + 0.86, y + 0.50, 6.60, 0.80, [(body, 11.5, False, INK2, 0)])
        if i < 2:
            rule(s, MX, y + 1.24, 7.46)

    box(s, 8.85, 1.95, COL, 4.10, TINT_G)
    textbox(s, 9.17, 2.32, 3.1, 1.60,
            [("木が足りない\nという話では\nない。", 16, True, GREEN, 0)])
    rule(s, 9.17, 4.10, 1.60, RULE2, 1.2)
    textbox(s, 9.17, 4.34, 3.10, 1.50,
            [("戦後に植えたスギ・ヒノキの多くが、いま主伐期を迎えている。\n\n"
              "止まっているのは資源ではなく、伐って使うまでの仕組みの方。",
              11, False, INK2, 0)])
    notes(s, "3つとも「木が足りない」話ではない点を強調する。"
             "資源はあるのに回らない、という構造の問題として提示する。")
    return s


def slide_age_chart(prs):
    s = blank(prs)
    frame(s, "AGE CLASS", "人工林は、すでに主伐期に入っている", NOTE)
    textbox(s, MX, 1.62, CW, 0.30,
            [("齢級構成は高齢側に大きく偏っている。使わなければ、資源は立ったまま滞留する。",
              12.5, False, INK2, 0)])
    data = CategoryChartData()
    data.categories = ["～30年生", "31～50年生", "51～70年生", "71年生～"]
    data.add_series("人工林面積", (90, 210, 560, 140))
    chart = s.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(MX), Inches(2.10), Inches(CW), Inches(4.45), data).chart
    plot = style_axes(chart, GREEN)
    plot.vary_by_categories = True
    plot.gap_width = 110
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    for i, pt in enumerate(chart.plots[0].series[0].points):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = AMBER if i == 2 else GREEN
    chart_caption(chart, "齢級別の人工林面積（万ha・サンプル値）")
    notes(s, "51〜70年生が突出しているところを指す。"
             "この山を計画的に崩しながら植え直すのが今後20年の仕事になる。")
    return s


def slide_rate_chart(prs):
    s = blank(prs)
    frame(s, "SELF-SUFFICIENCY", "木材自給率は回復してきた。ただし水準はまだ4割台", NOTE)
    textbox(s, MX, 1.62, CW, 0.30,
            [("2000年代前半を底に上昇が続く。輸入材から国産材への揺り戻しが起きている。",
              12.5, False, INK2, 0)])
    data = CategoryChartData()
    data.categories = ["2000", "2005", "2010", "2015", "2020", "2024"]
    data.add_series("木材自給率", (18.2, 20.3, 26.0, 33.3, 41.8, 43.0))
    chart = s.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(MX), Inches(2.10), Inches(CW), Inches(4.45), data).chart
    plot = style_axes(chart, GREEN)
    plot.data_labels.position = XL_LABEL_POSITION.ABOVE
    ser = chart.plots[0].series[0]
    ser.format.line.color.rgb = GREEN
    ser.format.line.width = Pt(2.2)
    ser.marker.format.fill.solid()
    ser.marker.format.fill.fore_color.rgb = GREEN
    ser.marker.format.line.color.rgb = WHITE
    val = chart.value_axis
    val.minimum_scale, val.maximum_scale = 0.0, 50.0
    chart_caption(chart, "木材自給率の推移（％・サンプル値）")
    notes(s, "上向きのトレンドを示しつつ、半分以上はまだ輸入材だという点も併せて触れる。")
    return s


def slide_actions(prs):
    s = blank(prs)
    frame(s, "ACTIONS", "これからの3つの打ち手")
    cards = [
        ("01", GREEN, TINT_G, "路網と機械化",
         "作業道を計画的に入れ、高性能林業機械で搬出する。\n"
         "コストが下がらなければ、伐り出す判断自体ができない。"),
        ("02", AMBER, TINT_A, "確実な再造林",
         "伐ったら植える。コンテナ苗と伐採・植栽の一貫作業で\n"
         "植栽コストを圧縮し、伐り逃げの誘因を消す。"),
        ("03", BLUE, TINT_B, "出口をつくる",
         "中高層木造やCLTなど、国産材を量で使い切る需要を\n"
         "先に育てておく。出口がなければ①②は費用増で終わる。"),
    ]
    for i, (num, accent, fill, head, body) in enumerate(cards):
        x = MX + i * PITCH
        box(s, x, 1.95, COL, 3.05, fill)
        badge(s, x + 0.34, 2.24, 0.54, num, fill=accent, size=11)
        textbox(s, x + 0.34, 3.02, COL - 0.60, 0.36, [(head, 15, True, INK, 0)])
        textbox(s, x + 0.34, 3.50, COL - 0.60, 1.30, [(body, 11.5, False, INK2, 0)])

    box(s, MX, 5.20, CW, 1.50, BAND)
    textbox(s, MX + 0.30, 5.38, 11.3, 0.30,
            [("まとめ — 「守る森」から「回す森」へ", 13, True, INK, 0)])
    points = [
        "資源はもうある。足りないのは、伐って植えて使うまでの循環の仕組み。",
        "3つの打ち手は独立ではない。出口がなければ、前の2つは費用増にしかならない。",
        "次の50年の森の姿は、いま植える木で決まる。",
    ]
    for i, text in enumerate(points):
        y = 5.74 + i * 0.32
        tick(s, MX + 0.30, y + 0.02, GREEN, h=0.20)
        textbox(s, MX + 0.52, y, 11.0, 0.28, [(text, 11.5, False, INK2, 0)])
    notes(s, "3つはセットで初めて回る、という順序で締める。"
             "特に03の需要側がないと、01と02はコスト増にしかならない。")
    return s


def patch_theme_font(prs):
    """テーマの和文フォントを Meiryo にして、グラフ等の既定も揃える。"""
    part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    xml = part.blob.decode("utf-8")
    xml = xml.replace('<a:ea typeface=""/>', f'<a:ea typeface="{JP}"/>')
    part._blob = xml.encode("utf-8")


def main():
    # Windows のコンソール既定は cp1252 で、日本語のパスを print できない
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    patch_theme_font(prs)

    slide_title(prs)
    slide_overview(prs)
    slide_numbers(prs)
    slide_issues(prs)
    slide_age_chart(prs)
    slide_rate_chart(prs)
    slide_actions(prs)

    out = Path(__file__).with_name("日本の林業_サンプル.pptx")
    prs.save(out)
    print(f"saved: {out} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
