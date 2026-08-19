from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

GREEN_DARK  = RGBColor(0x1a, 0x3a, 0x2a)
GREEN_MID   = RGBColor(0x2d, 0x6a, 0x4f)
GREEN_LIGHT = RGBColor(0xe5, 0xef, 0xe9)
AMBER       = RGBColor(0xc8, 0x80, 0x0a)
AMBER_LIGHT = RGBColor(0xfd, 0xf4, 0xe3)
RED         = RGBColor(0xc0, 0x39, 0x2b)
GRAY        = RGBColor(0x5a, 0x6e, 0x64)
WHITE       = RGBColor(0xff, 0xff, 0xff)
YELLOW_BG   = RGBColor(0xff, 0xf9, 0xc3)

def set_cell_bg(cell, rgb):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def para(text, bold=False, color=None, size=10, italic=False, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return p

# TITLE
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run('Test kreacji — tura 2: wynik i rekomendacja')
r.bold = True; r.font.size = Pt(20); r.font.color.rgb = GREEN_DARK

para('Tatrzanski Park Narodowy  ·  Meta Ads  ·  28 lipca – 14 sierpnia 2026', color=GRAY, size=10)
para('Budzet: ~600 PLN / wariant  ·  Cel: Sprzedaz  ·  Odbiorcy: Zainteresowania', color=GRAY, size=10, space_after=14)

# CONTEXT NOTE
t = doc.add_table(rows=1, cols=1)
t.style = 'Table Grid'
cell = t.rows[0].cells[0]
set_cell_bg(cell, YELLOW_BG)
cp = cell.paragraphs[0]
r = cp.add_run('Dlaczego robimy ture 2?\n')
r.bold = True; r.font.size = Pt(10)
r2 = cp.add_run(
    'Tura 1 (maj 2026) zakonczyla sie remisem na CPA (30,65 vs 30,69 PLN) przy malej skali '
    '(13 zakupow na wariant). Wyniki byly kierunkowe, nie pewne. '
    'Ewa Holek-Krzysztof zaproponowala powtorzenie testu z budzetem 600 PLN / wariant, '
    'by osiagnac 20+ zakupow i wiarygodny wynik.'
)
r2.font.size = Pt(10)
doc.add_paragraph()

# SECTION: RESULTS TABLE
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(6)
r = p.add_run('WYNIKI TURY 2')
r.bold = True; r.font.size = Pt(9); r.font.color.rgb = GREEN_MID

headers = ['Metryka', '\U0001f33f Plener', '\U0001f3ac Studio (lacznie)', 'Lepszy wynik']
rows = [
    ('Wydano (PLN)',     '599,92',  '598,01',  '—'),
    ('Wyswietlenia',    '104 277', '∲48 900', '\U0001f33f Plener'),
    ('CPM (PLN)',        '5,75',    '8,23–9,77', '\U0001f33f Plener'),
    ('CTR link',         '1,01%',   '0,75–0,86%', '\U0001f33f Plener'),
    ('CPC link (PLN)',   '0,57',    '1,09–1,13', '\U0001f33f Plener'),
    ('Zakupy',           '22',      '13',      '\U0001f33f Plener'),
    ('⭐ CPA (PLN)', '27,27',   '∲46,00', '\U0001f33f Plener'),
    ('Ranking jakosci',  'Powyzej sredniej', '—', '\U0001f33f Plener'),
]

t1 = doc.add_table(rows=len(rows)+1, cols=4)
t1.style = 'Table Grid'
for i, h in enumerate(headers):
    cell = t1.rows[0].cells[i]
    set_cell_bg(cell, GREEN_DARK)
    r = cell.paragraphs[0].add_run(h)
    r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(10)

for ri, row in enumerate(rows):
    is_cpa = row[0].startswith('⭐')
    for ci, val in enumerate(row):
        cell = t1.rows[ri+1].cells[ci]
        if is_cpa: set_cell_bg(cell, GREEN_LIGHT)
        r = cell.paragraphs[0].add_run(val)
        r.font.size = Pt(10)
        if is_cpa: r.bold = True
        if ci == 1 and 'Plener' in row[3]: r.font.color.rgb = GREEN_MID; r.bold = True
        if ci == 2 and 'Studio' in row[3]: r.font.color.rgb = GREEN_MID; r.bold = True
        if is_cpa and ci == 2: r.font.color.rgb = RED

doc.add_paragraph()

# SECTION: COMPARISON
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run('POROWNANIE OBU TUR')
r.bold = True; r.font.size = Pt(9); r.font.color.rgb = GREEN_MID

comp_headers = ['', '\U0001f5d3 Tura 1 (maj)', '\U0001f5d3 Tura 2 (lip–sie)']
comp_rows = [
    ('Budzet / wariant', '~200–400 PLN', '~600 PLN'),
    ('Zakupy plener',    '13',               '22'),
    ('Zakupy studio',    '13',               '13'),
    ('CPA plener',       '30,65 PLN',        '27,27 PLN'),
    ('CPA studio',       '∲30,69 PLN',  '∲46,00 PLN'),
    ('Wynik',            '\U0001f91d Remis',  '\U0001f33f Plener wygrywa'),
]

t2 = doc.add_table(rows=len(comp_rows)+1, cols=3)
t2.style = 'Table Grid'
for i, h in enumerate(comp_headers):
    cell = t2.rows[0].cells[i]
    set_cell_bg(cell, GREEN_DARK)
    r = cell.paragraphs[0].add_run(h)
    r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(10)

for ri, row in enumerate(comp_rows):
    is_result = row[0] == 'Wynik'
    for ci, val in enumerate(row):
        cell = t2.rows[ri+1].cells[ci]
        if is_result: set_cell_bg(cell, GREEN_LIGHT)
        r = cell.paragraphs[0].add_run(val)
        r.font.size = Pt(10)
        if is_result: r.bold = True
        if ci == 0: r.font.color.rgb = GRAY

doc.add_paragraph()

# WINNER
wt = doc.add_table(rows=1, cols=1)
wt.style = 'Table Grid'
wc = wt.rows[0].cells[0]
set_cell_bg(wc, GREEN_LIGHT)
wp = wc.paragraphs[0]
r = wp.add_run('ZWYCIEZCA: Zdjecia plenerowe\n')
r.bold = True; r.font.size = Pt(14); r.font.color.rgb = GREEN_DARK
r2 = wp.add_run(
    'Przy rownym budzecie (~600 PLN / wariant) plener wygenerował 70% wiecej zakupow '
    '(22 vs 13) i CPA nizszy o 19 PLN (27,27 vs ~46 PLN). '
    'CPM plenerowy jest znacznie tanszy (5,75 vs ~9 PLN) — Meta faworyzuje te kreacje w aukcji. '
    'Ranking jakosci: powyzej sredniej. Wynik jednoznaczny.'
)
r2.font.size = Pt(10)
doc.add_paragraph()

# RECOMMENDATIONS
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run('REKOMENDACJE')
r.bold = True; r.font.size = Pt(9); r.font.color.rgb = GREEN_MID

for icon, title, text in [
    ('\U0001f680', 'Skaluj plener',
     'Plener zostaje jako glowna kreacja sprzedazowa. Mozna zwiekszyc budzet i przetestowac na lookalike 1-3%.'),
    ('\U0001f4e6', 'Studio odloz',
     'Studio nie uzasadnia dalszego testowania przy tym budzecie. Wysoki CPM i CPA wskazuja, ze Meta nie faworyzuje tej kreacji w aukcji.'),
    ('⏱', 'Czas testu: 14-17 dni minimum',
     '17 dni przy 600 PLN / wariant pozwolilo zebrac 22 zakupy dla plenerowego — to wystarczajaca skala. Krotszy czas wrocilby do sytuacji z tury 1 (za malo danych).'),
]:
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(6)
    r = p2.add_run(f'{icon}  {title}\n')
    r.bold = True; r.font.size = Pt(11); r.font.color.rgb = GREEN_DARK
    r2 = p2.add_run(text)
    r2.font.size = Pt(10)

doc.add_paragraph()
para('Tatrzanski Park Narodowy  ·  Meta Ads  ·  Kampania: 20260728_O zwierzetach – test AB  ·  Sierpien 2026',
     color=GRAY, size=9, italic=True)

doc.save('/home/user/claude-ads/raport-test-ab-tura2.docx')
print('done')
