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

GREEN  = RGBColor(0x1a, 0x3a, 0x2a)
GRAY   = RGBColor(0x5a, 0x6e, 0x64)
WHITE  = RGBColor(0xff, 0xff, 0xff)
AMBER  = RGBColor(0xb8, 0x6a, 0x00)
GREEN_L= RGBColor(0xe5, 0xef, 0xe9)
RED    = RGBColor(0xc0, 0x39, 0x2b)
GREEN_M= RGBColor(0x2d, 0x6a, 0x4f)

def set_cell_bg(cell, rgb):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def p(text, bold=False, color=None, size=10, italic=False, sa=4, sb=0):
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(sa)
    par.paragraph_format.space_before = Pt(sb)
    r = par.add_run(text)
    r.bold = bold; r.italic = italic; r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return par

# ── TITLE ──
p('Test kreacji — tura 2', bold=True, size=20, color=GREEN, sa=2)
p('Tatrzanski Park Narodowy  ·  Meta Ads  ·  28 lipca – 14 sierpnia 2026', color=GRAY, size=10, sa=2)
p('Budzet: ~600 PLN / wariant  ·  Cel: Sprzedaz  ·  Odbiorcy: Zainteresowania', color=GRAY, size=10, sa=16)

# ── CONTEXT ──
p('Dlaczego tura 2?', bold=True, size=11, color=GREEN, sa=4)
p('Tura 1 (maj 2026) zakonczyla sie remisem na CPA (30,65 vs 30,69 PLN) przy 13 zakupach na wariant — zbyt mala skala by wyciagnac wnioski. Na prosbę Ewy Holek-Krzysztof powtorzylismy test z budzetem 600 PLN / wariant i minimalnym czasem 17 dni.', sa=14)

# ── TABLE ──
p('Wyniki tury 2', bold=True, size=11, color=GREEN, sa=6)

headers = ['Metryka', 'Plener', 'Studio (lacznie)', 'Lepszy']
rows = [
    ('Wydano (PLN)',      '599,92',      '598,01',       '—'),
    ('Wyswietlenia',     '104 277',      '~48 900',      'Plener'),
    ('CPM (PLN)',         '5,75',         '8,23–9,77',   'Plener'),
    ('CTR link',          '1,01%',        '0,75–0,86%',  'Plener'),
    ('CPC link (PLN)',    '0,57',         '1,09–1,13',   'Plener'),
    ('Zakupy',            '22',           '13',           'Plener'),
    ('CPA (PLN) ⭐',     '27,27',        '~46,00',       'Plener'),
    ('Ranking jakosci',   'Powyzej sr.', '—',            'Plener'),
]

t = doc.add_table(rows=len(rows)+1, cols=4)
t.style = 'Table Grid'
for i, h in enumerate(headers):
    cell = t.rows[0].cells[i]
    set_cell_bg(cell, GREEN)
    r = cell.paragraphs[0].add_run(h)
    r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(10)

for ri, row in enumerate(rows):
    is_cpa = '⭐' in row[0]
    for ci, val in enumerate(row):
        cell = t.rows[ri+1].cells[ci]
        if is_cpa: set_cell_bg(cell, GREEN_L)
        r = cell.paragraphs[0].add_run(val)
        r.font.size = Pt(10)
        if is_cpa: r.bold = True
        if ci == 1 and row[3] == 'Plener': r.font.color.rgb = GREEN_M; r.bold = True
        if ci == 2 and row[3] == 'Plener': r.font.color.rgb = RED

doc.add_paragraph()

# ── COMPARISON ──
p('Porownanie obu tur', bold=True, size=11, color=GREEN, sa=6)

comp_h = ['', 'Tura 1 (maj)', 'Tura 2 (lip–sie)']
comp_r = [
    ('Budzet / wariant', '~200–400 PLN', '~600 PLN'),
    ('Zakupy plener',    '13',            '22'),
    ('Zakupy studio',    '13',            '13'),
    ('CPA plener',       '30,65 PLN',     '27,27 PLN'),
    ('CPA studio',       '~30,69 PLN',    '~46,00 PLN'),
    ('Wynik',            'Remis',         'Plener wygrywa'),
]

t2 = doc.add_table(rows=len(comp_r)+1, cols=3)
t2.style = 'Table Grid'
for i, h in enumerate(comp_h):
    cell = t2.rows[0].cells[i]
    set_cell_bg(cell, GREEN)
    r = cell.paragraphs[0].add_run(h)
    r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(10)

for ri, row in enumerate(comp_r):
    is_result = row[0] == 'Wynik'
    for ci, val in enumerate(row):
        cell = t2.rows[ri+1].cells[ci]
        if is_result: set_cell_bg(cell, GREEN_L)
        r = cell.paragraphs[0].add_run(val)
        r.font.size = Pt(10)
        if is_result: r.bold = True
        if ci == 0: r.font.color.rgb = GRAY

doc.add_paragraph()

# ── WINNER ──
p('Zwyciezca: Zdjecia plenerowe', bold=True, size=13, color=GREEN, sa=4)
p('Przy rownym budzecie (~600 PLN / wariant) plener wygenerował 70% wiecej zakupow (22 vs 13) i CPA nizszy o 19 PLN (27,27 vs ~46 PLN). CPM plenerowy jest znacznie tanszy (5,75 vs ~9 PLN). Ranking jakosci: powyzej sredniej. Wynik jednoznaczny.', sa=16)

# ── REKOMENDACJE ──
p('Rekomendacje', bold=True, size=11, color=GREEN, sa=6)
for title, text in [
    ('Skaluj plener', 'Plener zostaje jako glowna kreacja sprzedazowa. Mozna zwiekszyc budzet i przetestowac na lookalike 1–3%.'),
    ('Studio odloz', 'Studio nie uzasadnia dalszego testowania. Wysoki CPM i CPA wskazuja, ze Meta nie faworyzuje tej kreacji w aukcji.'),
    ('Czas testu: 14–17 dni minimum', '17 dni przy 600 PLN / wariant pozwolilo zebrac 22 zakupy — wystarczajaca skala. Krotszy czas wrocilby do sytuacji z tury 1.'),
]:
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(6)
    r = par.add_run(title + '\n')
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = GREEN
    r2 = par.add_run(text)
    r2.font.size = Pt(10)

# ── DIVIDER ──
doc.add_paragraph()
par = doc.add_paragraph()
par.paragraph_format.space_after = Pt(12)
r = par.add_run('─' * 60)
r.font.size = Pt(9); r.font.color.rgb = GRAY

# ── UWAGI (in amber) ──
p('UWAGI', bold=True, size=11, color=AMBER, sa=6)

for text in [
    'Studio w tym tescie bylo podzielone na dwa zestawy reklam z roznymi oknami atrybucji (klikniecia z 7 dni vs klikniecia z 7 dni lub wyswietlenia z 1 dnia). Lacznie wydano ~598 PLN i wygenerowano 13 zakupow. Dane sa spojne — nie ma rozjezdzania sie wynikow miedzy wariantami.',
    'Wyniki studio moga byc czesciowo zaburzone przez wyzszy CPM — Meta faworyzuje w aukcji kreacje o wyzszym CTR i nizszym koszcie, co само w sobie jest sygnałem, ze algorytm "woli" plener. Nie jest to wina samego zdjecia, ale i tak jest to informacja operacyjna: ta kreacja jest drozsza w emisji.',
    'Przy 22 zakupach dla plenerowego i 13 dla studyjnego mamy wystarczajaca skale, zeby ten wynik traktowac jako wiarygodny, nie przypadkowy. Poprzedni test (tura 1, 13 vs 13 zakupow) tego progu nie osiagal.',
]:
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(8)
    par.paragraph_format.left_indent = Cm(0.5)
    r = par.add_run('– ')
    r.font.size = Pt(10); r.font.color.rgb = AMBER; r.bold = True
    r2 = par.add_run(text)
    r2.font.size = Pt(10); r2.font.color.rgb = AMBER

doc.add_paragraph()
p('Tatrzanski Park Narodowy  ·  Meta Ads  ·  Sierpien 2026',
  color=GRAY, size=9, italic=True)

doc.save('/home/user/claude-ads/raport-test-ab-tura2.docx')
print('done')
