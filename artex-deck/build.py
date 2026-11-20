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

# ---------- 3 診断の流れ ----------
s = frame('診断の流れ', 3,
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

# ---------- 4 期間と費用 ----------
s = frame('診断の期間と費用', 4, '費用は診断する範囲に応じてお見積もりします。30万円（税抜）は最小構成の目安です。')
lead(s, '1回の診断を約1週間で終え、診断報告書を渡す')


def big_number(s, cx, pre, num, unit, cap):
    nw = len(num) * 280 * 0.62
    pw = len(pre) * 64 + (16 if pre else 0)
    uw = sum(0.6 if ord(c) < 128 else 1 for c in unit) * 64
    total = pw + nw + 16 + uw; x = cx - total / 2; base = 800
    shape(s, MSO_SHAPE.RECTANGLE, x + pw - 10, base - 95, nw + 16 + uw + 50, 90, fill=MARK)
    if pre: text(s, x, base - 90, pw, 84, pre, 64, anchor='b')
    text(s, x + pw, base - 300, nw + 10, 300, num, 280, color=A, anchor='b')
    text(s, x + pw + nw + 16, base - 90, uw + 20, 84, unit, 64, color=A, anchor='b')
    text(s, cx - 340, base + 60, 680, 50, cap, 36, align='c')


big_number(s, 600, '約', '1', '週間', '診断期間（1回）')
shape(s, MSO_SHAPE.RECTANGLE, 957, 430, 6, 520, fill=FRAME)
big_number(s, 1320, '', '30', '万円〜', '費用（税抜）')

# ---------- 5 安全性 ----------
s = frame('診断環境の安全性', 5,
          '攻撃ツールARTEXとその中で使うAIモデルは、RAYVEN社内のGPUサーバーで動かします。診断のデータを外部のAIサービスに送ることはありません。'
          '通信はすべて記録し、診断は事前に合意した範囲だけを対象にします。')
lead(s, '診断はRAYVEN社内のGPUで、AIごと動かす', '診断のデータは外部のAIサービスに渡さない')
bx, by, bw, bh = 240, 370, 1000, 600
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh, fill=WHITE, line=A, lw=8, radius=24)
pill(s, bx + bw / 2, by + 40, 'RAYVEN社内のGPU', size=30)
for i, (ic, nm) in enumerate([('search', '攻撃ツール'), ('gpu', 'AIモデル'), ('log', '通信の記録')]):
    x = bx + 48 + i * 312
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, by + 150, 280, 400, fill=WHITE, line=FRAME, lw=6, radius=24)
    s.shapes.add_picture(ICON(ic), px(x + 60), px(by + 200), px(160), px(160))
    text(s, x + 20, by + 410, 240, 50, nm, 36, align='c')
arrow_line(s, bx + bw + 4, by + bh / 2, 1440, by + bh / 2, head=False)
shape(s, MSO_SHAPE.OVAL, 1290, by + bh / 2 - 50, 100, 100, fill=WHITE, line=A, lw=8)
text(s, 1290, by + bh / 2 - 50, 100, 100, '×', 60, color=A, align='c', anchor='m')
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1440, by + 130, 240, 340, fill=WHITE, line=FRAME, lw=6, radius=24, dash=True)
s.shapes.add_picture(ICON('cloud'), px(1490), px(by + 170), px(140), px(140))
text(s, 1460, by + 340, 200, 100, '外部の\nAIサービス', 30, color=SUB, align='c')

# ---------- 6 次のアクション ----------
s = frame('ご相談の流れ', 6, 'まずはメールで、診断したい対象をお知らせください。範囲とお見積もりを決めてから、約1週間で診断と報告を行います。')
lead(s, '診断したい範囲を、まず一緒に決める')
step_boxes(s, 240, 290, 440, 470, 60,
           [('mail', 'お問い合わせ', 'メールで受け付ける'),
            ('scope', '範囲の合意', '対象と見積もりを決める'),
            ('report', '診断と報告', '約1週間で実施し報告する')],
           [ST[0], ST[2], ST[3]], icon=150)
mail = 'info@rayven.cloud'
mw = len(mail) * 72 * 0.62
shape(s, MSO_SHAPE.RECTANGLE, 960 - mw / 2 - 25, 880, mw + 50, 40, fill=MARK)
text(s, 240, 820, 1440, 100, mail, 72, color=A, align='c', anchor='m')

out = os.path.join(HERE, 'Artex_サービス説明.pptx')
prs.save(out)
print(out)
