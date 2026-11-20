"""Artex サービス説明資料（natural-slides の決まり × RAYVENカラー）を .pptx で出力する。
座標は 1920×1080 基準の px。1px = 6350 EMU、文字は px × 0.5 = pt。"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn
from lxml import etree
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ICON = lambda n: os.path.join(HERE, 'icons', f'icon-{n}.png')
LOGO = os.path.join(HERE, 'rayven-logo-crop.png')

# RAYVENカラー（A = 濃紺）
A = '141E46'; ST = ['DCE3EF', 'AFC2E6', '4669B0', '141E46']
TEXT = '3C3C3C'; SUB = '6B6B6B'; WEAK = '9A9A9A'
FRAME = 'DDDBD5'; HAIR = 'E6E4DE'; ARROW = 'A3A5A8'; MARK = 'F9DC72'; WHITE = 'FFFFFF'
LATIN, JP = 'Montserrat', 'Noto Sans JP'

px = lambda v: Emu(int(round(v * 6350)))
rgb = lambda h: RGBColor.from_string(h)

prs = Presentation()
prs.slide_width, prs.slide_height = px(1920), px(1080)
BLANK = prs.slide_layouts[6]


def shape(s, kind, x, y, w, h, fill=None, line=None, lw=0, radius=None, dash=False):
    sp = s.shapes.add_shape(kind, px(x), px(y), px(w), px(h))
    if fill: sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    else: sp.fill.background()
    if line:
        sp.line.color.rgb = rgb(line); sp.line.width = px(lw)
        if dash: sp.line.dash_style = MSO_LINE.DASH
    else: sp.line.fill.background()
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = min(0.5, radius / min(w, h))
    st = sp._element.find(qn('p:style'))
    if st is not None: sp._element.remove(st)
    return sp


def text(s, x, y, w, h, t, size, color=TEXT, bold=True, align='l', anchor='t', weight900=False, latin=LATIN):
    tb = s.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {'t': MSO_ANCHOR.TOP, 'm': MSO_ANCHOR.MIDDLE, 'b': MSO_ANCHOR.BOTTOM}[anchor]
    for i, ln in enumerate(t.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {'l': PP_ALIGN.LEFT, 'c': PP_ALIGN.CENTER, 'r': PP_ALIGN.RIGHT}[align]
        p.line_spacing = 1.25
        r = p.add_run(); r.text = ln
        f = r.font; f.size = Pt(size * 0.5); f.bold = bold; f.color.rgb = rgb(color); f.name = latin
        rpr = r._r.get_or_add_rPr()
        ea = etree.SubElement(rpr, qn('a:ea')); ea.set('typeface', JP)
    return tb


def arrow_line(s, x1, y1, x2, y2, color=ARROW, lw=5, head=True):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, px(x1), px(y1), px(x2), px(y2))
    st = c._element.find(qn('p:style'))
    if st is not None: c._element.remove(st)
    c.line.color.rgb = rgb(color); c.line.width = px(lw)
    if head:
        ln = c.line._get_or_add_ln()
        t = etree.SubElement(ln, qn('a:tailEnd')); t.set('type', 'triangle'); t.set('w', 'lg'); t.set('len', 'lg')
    return c


def chevron(s, x, y, size=44):
    shape(s, MSO_SHAPE.CHEVRON, x, y, size * 0.7, size, fill=ARROW)


def pill(s, cx, y, label, size=28, outline=False):
    n = sum(0.6 if ord(ch) < 128 else 1 for ch in label)
    w = n * size * 1.15 + 88; h = size + (40 if outline else 36)
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx - w / 2, y, w, h, fill=WHITE if outline else A,
          line=A if outline else None, lw=6, radius=h / 2)
    text(s, cx - w / 2, y, w, h, label, size, color=A if outline else WHITE, align='c', anchor='m')
    return h


def marker_line(s, cx, y, label, size):
    """リード文2行目：黄色の長方形を語の下半分に敷き、上に文字を置く"""
    n = sum(0.6 if ord(ch) < 128 else 1 for ch in label)
    w = n * size + 40
    shape(s, MSO_SHAPE.RECTANGLE, cx - w / 2, y + size * 0.62, w, size * 0.62, fill=MARK)
    text(s, 240, y, 1440, size * 1.3, label, size, align='c')


def frame(title, page, notes=''):
    s = prs.slides.add_slide(BLANK)
    shape(s, MSO_SHAPE.RECTANGLE, 120, 66, 8, 48, fill=A)
    text(s, 152, 58, 1200, 64, title, 44, anchor='m')
    shape(s, MSO_SHAPE.RECTANGLE, 120, 146, 1680, 2, fill=HAIR)
    s.shapes.add_picture(LOGO, px(1733), px(44), height=px(72))
    text(s, 1680, 1020, 120, 32, f'{page:02d}', 24, color=WEAK, bold=False, align='r')
    if notes: s.notes_slide.notes_text_frame.text = notes
    return s


def lead(s, line1, mark=None, y=180):
    text(s, 240, y, 1440, 52, line1, 38, align='c')
    if mark: marker_line(s, 960, y + 60, mark, 38)


def step_boxes(s, x0, y, w, h, gap, steps, fills, icon=170, name=40, desc=30):
    for i, (ic, nm, ds) in enumerate(steps):
        x = x0 + i * (w + gap)
        f = fills[i]; tc = WHITE if f in ('4669B0', '141E46') else A
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=f, radius=24)
        s.shapes.add_picture(ICON(ic), px(x + (w - icon) / 2), px(y + 56), px(icon), px(icon))
        text(s, x + 32, y + 56 + icon + 40, w - 64, name * 1.3, nm, name, color=tc, align='c')
        if ds: text(s, x + 32, y + 56 + icon + 40 + name * 1.3 + 24, w - 64, desc * 1.4, ds, desc, color=tc, bold=False, align='c')
        if i < len(steps) - 1: chevron(s, x + w + (gap - 31) / 2, y + h / 2 - 22)


# ---------- 1 表紙 ----------
s = prs.slides.add_slide(BLANK)
shape(s, MSO_SHAPE.RECTANGLE, 0, 0, 640, 1080, fill=A)
text(s, 80, 360, 520, 300, 'AI\nペネトレーション\nテスト', 60, color=WHITE)
s.shapes.add_picture(LOGO, px(1733), px(44), height=px(72))
text(s, 760, 330, 1040, 100, '攻撃者が実際に使うAIで', 76, weight900=True)
shape(s, MSO_SHAPE.RECTANGLE, 750, 470 + 48, 8 * 76 + 40, 48, fill=MARK)
text(s, 770, 450, 1040, 100, '本番同様の診断を', 76)
text(s, 770, 640, 1000, 50, '企業の情報システム部門向け', 32, bold=False)
text(s, 770, 700, 1000, 50, '株式会社RAYVEN', 32, bold=False)
s.notes_slide.notes_text_frame.text = '攻撃者が実際に使っているAIペンテストツールARTEXで診断するため、本番の攻撃と同様の結果が得られます。'

# ---------- 2 課題 ----------
s = frame('いま起きている攻撃', 2,
          'ARTEXはもともと正規のAIペンテストツールとして2026年に公開されました。現在は攻撃者にも使われ、米FBIも注意を呼びかけています（ITECS, 2026）。'
          'Unit 42は、公開APIやパスワード再設定画面を入口に、1日未満でクラウドに到達した事例を報告しています。')
lead(s, '正規のAIペンテストツールARTEXが、攻撃にも使われている', '入口からクラウド到達まで1日未満')
step_boxes(s, 240, 360, 440, 560, 60,
           [('door', '入口を見つける', '公開APIや再設定画面'),
            ('chain', '弱点をつなぐ', '既知の弱点を組み合わせる'),
            ('cloud', 'クラウドへ到達', '社内のクラウドに入られる')],
           [ST[0], ST[2], ST[3]])
text(s, 240, 950, 1440, 34, '出典：Palo Alto Networks Unit 42（2026）、Qualys Blog（2026年11月15日）', 24, color=WEAK, bold=False, align='c')

# ---------- 3 比較 ----------
s = frame('診断方法の比較', 3,
          '運用保守だけでペンテストをしていない場合、実際に侵入できるかは分かりません。人が手作業で行う従来のペンテストは本番同様の結果が出ますが、費用は範囲により数十万円から数百万円と幅があります。'
          'この診断は、攻撃者と同じARTEXでAIが試し、結果を専門家が確かめます。')
lead(s, '人が行うペンテストと同じく、実際に侵入できるかまで確かめる', '攻撃者と同じAIの手口で、30万円から')
cols = [('運用保守のみ', '8A94A6'), ('従来のペンテスト', '8A94A6'), ('この診断', A)]
LX, LW, CW, GAP, TY = 240, 360, 354, 6, 360
for j, (h, c) in enumerate(cols):
    x = LX + LW + GAP + j * (CW + GAP)
    shape(s, MSO_SHAPE.RECTANGLE, x, TY, CW, 90, fill=c)
    text(s, x, TY, CW, 90, h, 32, color=WHITE, align='c', anchor='m')
rows = [('実際に侵入\nできるか', [('×', 'ペンテストをしない'), ('○', '人が試す'), ('○', 'AIが試し人が確認')]),
        ('攻撃者の\nAIツール', [('×', ''), ('—', '人の手作業'), ('○', '同じARTEXで試す')]),
        ('費用', [('—', ''), ('数十万〜数百万円', '範囲により変わる'), ('30万円〜', '税抜')])]
for i, (lb, cells) in enumerate(rows):
    y = TY + 90 + GAP + i * (150 + GAP)
    shape(s, MSO_SHAPE.RECTANGLE, LX, y, LW, 150, fill='EEEEEC')
    text(s, LX + 32, y, LW - 64, 150, lb, 32, anchor='m')
    for j, (sym, why) in enumerate(cells):
        x = LX + LW + GAP + j * (CW + GAP); own = j == 2
        shape(s, MSO_SHAPE.RECTANGLE, x, y, CW, 150, fill='E8EBF3' if own else 'F5F5F3')
        big = len(sym) == 1
        text(s, x + 16, y + (14 if why else 37), CW - 32, 76, sym, 64 if big else 34, color=A if own else TEXT, align='c', anchor='m')
        if why: text(s, x + 16, y + 96, CW - 32, 44, why, 28, color=TEXT, bold=False, align='c')
text(s, 240, 960, 1440, 34, '従来のペンテストの費用：CREX GROUP、ITトレンドの解説記事（2026年11月確認）', 24, color=WEAK, bold=False, align='c')

# ---------- 4 診断の流れ ----------
s = frame('診断の流れ', 4,
          '攻撃者が実際に使うツールで、攻撃者と同じ手順をたどります。偵察から攻撃の組み立てまではAIが自動で進めます。AIが見つけた結果は、専門家が確かめたうえで対策とあわせて報告します。')
lead(s, '攻撃者が実際に使うARTEXで、同じ手順をたどる', '本番の攻撃と同様の結果が分かる')
W, G, X0, Y = 324, 48, 240, 380
pill(s, X0 + (3 * W + 2 * G) / 2, Y, 'AIが自動で実行')
pill(s, X0 + 3 * (W + G) + W / 2, Y, '専門家が確認')
step_boxes(s, X0, Y + 100, W, 500, G,
           [('search', '偵察', '入口を洗い出す'),
            ('warn', '弱点の発見', '入口ごとに探す'),
            ('chain', '攻撃の組み立て', '侵入できるか試す'),
            ('report', '確認と報告', '対策とあわせて報告')],
           ST, icon=150, name=36, desc=28)

# ---------- 5 安全性 ----------
s = frame('診断の安全性', 5,
          '診断は、事前に合意した時間帯だけ実施し、業務への影響を避けます。攻撃ツールとAIモデルはRAYVEN社内のGPUで動くローカルLLMで動かし、診断のデータを外部のAIサービスに送ることはありません。通信はすべて記録します。')
lead(s, '攻撃者と同じ道具を使っても、業務とデータは守る')
for i, (no, ic, hd, ln) in enumerate([('point 1', 'clock', '時間帯を決める', '合意した時間帯だけ実施する'),
                                      ('point 2', 'gpu', '社内のローカルLLM', '診断データを外部のAIに渡さない')]):
    x = 240 + i * 740
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 300, 700, 560, fill=WHITE, line=A if i else FRAME, lw=8, radius=24)
    pill(s, x + 350, 340, no, size=28, outline=True)
    s.shapes.add_picture(ICON(ic), px(x + 260), px(420), px(180), px(180))
    text(s, x + 40, 640, 620, 70, hd, 48, color=A, align='c')
    text(s, x + 40, 740, 620, 50, ln, 32, align='c')
text(s, 240, 920, 1440, 50, '通信はすべて記録し、何をしたかを後から追える', 30, align='c')

# ---------- 6 期間・費用・相談 ----------
s = frame('期間と費用', 6, '費用は診断する範囲に応じてお見積もりします。まずはメールで、診断したい対象をお知らせください。')
lead(s, '1回の診断を約1週間で終え、診断報告書を渡す')


def big_number(s, cx, pre, num, unit, cap, size=220, base=600):
    nw = len(num) * size * 0.62; us = size * 0.25
    pw = len(pre) * us + (16 if pre else 0)
    uw = sum(0.6 if ord(c) < 128 else 1 for c in unit) * us
    total = pw + nw + 16 + uw; x = cx - total / 2; mh = size / 3
    shape(s, MSO_SHAPE.RECTANGLE, x + pw - 10, base - mh - 5, nw + 16 + uw + 50, mh, fill=MARK)
    if pre: text(s, x, base - us * 1.4, pw, us * 1.3, pre, us, anchor='b')
    text(s, x + pw, base - size * 1.07, nw + 10, size * 1.07, num, size, color=A, anchor='b')
    text(s, x + pw + nw + 16, base - us * 1.4, uw + 20, us * 1.3, unit, us, color=A, anchor='b')
    text(s, cx - 340, base + 40, 680, 50, cap, 34, align='c')


big_number(s, 600, '約', '1', '週間', '診断期間（1回）')
shape(s, MSO_SHAPE.RECTANGLE, 957, 330, 6, 360, fill=FRAME)
big_number(s, 1320, '', '30', '万円〜', '費用（税抜）')
shape(s, MSO_SHAPE.RECTANGLE, 240, 740, 1440, 2, fill=HAIR)
text(s, 240, 780, 1440, 50, '診断したい範囲を、まずメールで相談する', 36, align='c')
mail = 'info@rayven.cloud'
mw = len(mail) * 80 * 0.62
shape(s, MSO_SHAPE.RECTANGLE, 960 - mw / 2 - 25, 912, mw + 50, 40, fill=MARK)
text(s, 240, 850, 1440, 110, mail, 80, color=A, align='c', anchor='m')

out = os.path.join(HERE, 'Artex_サービス説明.pptx')
prs.save(out)
print(out)
