"""Jana Terma Rujukan JKKWPK (Word) dalam format korporat.

Guna: python3 build_tor.py  -> TOR_JKKWPK_Draf_Mesyuarat.docx
Kemudian: soffice --headless --convert-to pdf TOR_JKKWPK_Draf_Mesyuarat.docx
"""
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

GREEN = "0B4F33"      # hijau gelap korporat
GOLD = "B8913A"       # emas
LIGHT = "EEF3EF"      # latar sel lembut
GREY = "6B6B6B"
FONT = "Arial"
A = "assets/"

doc = Document()

# ---------- gaya asas ----------
st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
pf = st.paragraph_format
pf.space_after = Pt(6)
pf.space_before = Pt(0)
pf.line_spacing = 1.35

sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.left_margin = sec.right_margin = Cm(2.5)
sec.top_margin = Cm(2.0)
sec.bottom_margin = Cm(2.0)
sec.footer_distance = Cm(1.0)


def rgb(h):
    return RGBColor.from_string(h)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), fill)
    tcPr.append(sh)


def cell_borders(cell, color="FFFFFF", sz=0, sides=("top", "left", "bottom", "right")):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for s in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{s}")
        if s in sides and sz:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(sz))
            e.set(qn("w:color"), color)
        else:
            e.set(qn("w:val"), "nil")
        b.append(e)
    tcPr.append(b)


def cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        e = OxmlElement(f"w:{k}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        m.append(e)
    tcPr.append(m)


def table_borders(table, color="BFBFBF", sz=4, inside=True):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for s in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{s}")
        if s.startswith("inside") and not inside:
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(sz))
            e.set(qn("w:color"), color)
        b.append(e)
    tblPr.append(b)


def para_border(p, side="top", color=GREEN, sz=24, space=4):
    pPr = p._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    e = OxmlElement(f"w:{side}")
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(sz))
    e.set(qn("w:space"), str(space))
    e.set(qn("w:color"), color)
    pb.append(e)
    pPr.append(pb)


def keep_next(p):
    p.paragraph_format.keep_with_next = True


def add_field(run, instr):
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1); run._r.append(it); run._r.append(f2)


def runs(p, text, bold=False, size=None, color=None, italic=False):
    """Sokong **tebal** ringkas dalam teks."""
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        r = p.add_run(part)
        r.bold = bold or (i % 2 == 1)
        r.italic = italic
        if size:
            r.font.size = Pt(size)
        if color:
            r.font.color.rgb = rgb(color)
    return p


# ---------- elemen kandungan ----------
L1, L2, L3 = Cm(1.25), Cm(2.5), Cm(3.25)


def heading(num, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = L1
    p.paragraph_format.first_line_indent = -L1
    p.paragraph_format.tab_stops.add_tab_stop(L1)
    runs(p, f"{num}\t{title}", bold=True, size=12, color=GREEN)
    para_border(p, "bottom", GOLD, 6, 2)
    keep_next(p)


def subheading(num, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.left_indent = L2
    p.paragraph_format.first_line_indent = -(L2 - L1)
    p.paragraph_format.tab_stops.add_tab_stop(L2)
    runs(p, f"{num}\t{title}", bold=True, color=GREEN)
    keep_next(p)


def clause(num, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = L2
    p.paragraph_format.first_line_indent = -(L2 - L1)
    p.paragraph_format.tab_stops.add_tab_stop(L2)
    runs(p, f"{num}\t{text}")
    return p


def body(text, indent=L2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = indent
    runs(p, text)
    return p


def item(label, text, indent=L2, gap=Cm(0.75)):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = indent + gap
    p.paragraph_format.first_line_indent = -gap
    p.paragraph_format.tab_stops.add_tab_stop(indent + gap)
    p.paragraph_format.space_after = Pt(3)
    runs(p, f"{label}\t{text}")
    return p


def letters(items, indent=L2):
    for i, t in enumerate(items):
        item(f"{chr(97 + i)}.", t, indent)
    doc.paragraphs[-1].paragraph_format.space_after = Pt(8)


def decision(text):
    """Kotak 'Untuk Keputusan Mesyuarat'."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    c = t.cell(0, 0)
    c.width = Cm(13.5)
    shade(c, "FBF5E6")
    cell_borders(c, GOLD, 12, ("left",))
    cell_margins(c, 100, 100, 160, 120)
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    runs(p, "UNTUK KEPUTUSAN MESYUARAT:  ", bold=True, size=9.5, color="8A6A1F")
    runs(p, text, size=10)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def data_table(headers, rows, widths, indent_cm=2.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table_borders(t, "C9D3CC", 4)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        shade(c, GREEN)
        cell_margins(c)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        runs(p, h, bold=True, size=10, color="FFFFFF")
    # ulang baris tajuk jika merentas halaman
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]
            cell_margins(c)
            if ri % 2 == 1:
                shade(c, LIGHT)
            lines = val if isinstance(val, list) else [val]
            for li, line in enumerate(lines):
                p = c.paragraphs[0] if li == 0 else c.add_paragraph()
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                if isinstance(val, list):
                    p.paragraph_format.left_indent = Cm(0.4)
                    p.paragraph_format.first_line_indent = Cm(-0.4)
                    runs(p, "•  " + line, size=10)
                else:
                    runs(p, line, size=10)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


# =====================================================================
# MUKA DEPAN
# =====================================================================
def logo_row(files, height_cm, widths_cm):
    t = doc.add_table(rows=1, cols=len(files))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, f in enumerate(files):
        c = t.cell(0, i)
        c.width = Cm(widths_cm)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell_borders(c)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_picture(A + f, height=Cm(height_cm))
    return t


# jalur atas
band = doc.add_table(rows=1, cols=1)
band.alignment = WD_TABLE_ALIGNMENT.CENTER
bc = band.cell(0, 0)
bc.width = Cm(16)
shade(bc, GREEN)
cell_borders(bc)
cell_margins(bc, 60, 60, 0, 0)
bp = bc.paragraphs[0]
bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_after = Pt(0)
runs(bp, "MAJLIS AGAMA ISLAM NEGERI PULAU PINANG", bold=True, size=9, color="FFFFFF")

sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(18)

logo_row(["jata_pulau_pinang.png", "logo_mainpp.png"], 3.2, 4.5)

sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(26)

# kotak tajuk berbingkai emas
box = doc.add_table(rows=1, cols=1)
box.alignment = WD_TABLE_ALIGNMENT.CENTER
c = box.cell(0, 0)
c.width = Cm(15.5)
cell_borders(c, GOLD, 24)
cell_margins(c, 280, 280, 300, 300)
p = c.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
r = p.add_run("T E R M A   R U J U K A N")
r.font.size = Pt(26); r.font.color.rgb = rgb(GREEN); r.bold = True
p2 = c.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(4)
runs(p2, "JAWATANKUASA KEBAJIKAN, WANITA, KESIHATAN\nDAN PEMBANGUNAN KELUARGA", bold=True, size=13, color="333333")
p3 = c.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(0)
runs(p3, "(JKKWPK)", bold=True, size=13, color=GOLD)
p4 = c.add_paragraph()
p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
p4.paragraph_format.space_before = Pt(10)
p4.paragraph_format.space_after = Pt(0)
runs(p4, "MAJLIS AGAMA ISLAM NEGERI PULAU PINANG (MAINPP)", size=10.5, color=GREY)

sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(22)

# maklumat dokumen
info = doc.add_table(rows=4, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
table_borders(info, "C9D3CC", 4)
meta = [
    ("Status Dokumen", "Draf untuk Pertimbangan Mesyuarat JKKWPK"),
    ("Disediakan oleh", "Urus Setia JKKWPK, Bahagian Undang-Undang Keluarga Islam (BUUKI), JHEAIPP"),
    ("Untuk Diperakukan kepada", "Mesyuarat Majlis Agama Islam Negeri Pulau Pinang"),
    ("Tarikh Draf", "November 2026"),
]
for i, (k, v) in enumerate(meta):
    a, b = info.rows[i].cells
    a.width, b.width = Cm(5), Cm(10.5)
    shade(a, LIGHT)
    for cc in (a, b):
        cell_margins(cc, 70, 70, 120, 120)
        cc.paragraphs[0].paragraph_format.space_after = Pt(0)
        cc.paragraphs[0].paragraph_format.line_spacing = 1.1
    runs(a.paragraphs[0], k, bold=True, size=9.5, color=GREEN)
    runs(b.paragraphs[0], v, size=9.5)

sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(30)

lbl = doc.add_paragraph()
lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
lbl.paragraph_format.space_after = Pt(8)
runs(lbl, "AGENSI RAKAN STRATEGIK", bold=True, size=8.5, color=GREY)
logo_row(["logo_zakat_mainpp.png", "logo_lppkn.png", "logo_jkm.png", "logo_jkn.png"], 1.7, 4.0)

# =====================================================================
# KANDUNGAN — seksyen baharu (pengaki bermula di sini)
# =====================================================================
sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
sec2.footer.is_linked_to_previous = False
sec2.header.is_linked_to_previous = False
pg = OxmlElement("w:pgNumType"); pg.set(qn("w:start"), "1"); sec2._sectPr.append(pg)

# pengaki (halaman selain muka depan)
ft = sec2.footer
fp = ft.paragraphs[0]
fp.text = ""
para_border(fp, "top", GREEN, 36, 6)
from docx.enum.text import WD_TAB_ALIGNMENT
fp.paragraph_format.tab_stops.add_tab_stop(Cm(15.9), alignment=WD_TAB_ALIGNMENT.RIGHT)
from docx.shared import Twips
for pos in (Twips(4680), Twips(9360)):
    fp.paragraph_format.tab_stops.add_tab_stop(pos, alignment=WD_TAB_ALIGNMENT.CLEAR)
runs(fp, "TERMA RUJUKAN JKKWPK  |  MAINPP", size=8, color=GREY)
r = fp.add_run("\t"); r.font.size = Pt(8)
r = fp.add_run(); r.font.size = Pt(8); r.font.color.rgb = rgb(GREY)
add_field(r, "PAGE")
# pengepala kecil
hp = sec2.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
runs(hp, "DRAF UNTUK PERTIMBANGAN MESYUARAT JKKWPK", bold=True, size=7.5, color=GOLD)

# ---------- 1.0 ----------
heading("1.0", "MATLAMAT DAN TUJUAN PENUBUHAN")
clause("1.1", "Jawatankuasa Kebajikan, Wanita, Kesihatan dan Pembangunan Keluarga (JKKWPK) ditubuhkan sebagai jawatankuasa di bawah Majlis Agama Islam Negeri Pulau Pinang (MAINPP) bagi merancang, menyelaras, memantau dan menilai pelaksanaan program berkaitan kebajikan, wanita, kesihatan dan pembangunan keluarga Islam di Negeri Pulau Pinang.")
clause("1.2", "Penubuhan jawatankuasa ini bertujuan memperkukuh kerjasama strategik antara MAINPP dengan agensi-agensi berkaitan dalam merangka dan melaksanakan program yang memberi manfaat kepada masyarakat Islam, khususnya dalam aspek kesejahteraan keluarga, kebajikan sosial, kesihatan dan pembangunan komuniti.")
clause("1.3", "Jawatankuasa ini juga berperanan sebagai platform perbincangan, penyelarasan dan pemantauan bagi memastikan program yang dirancang dilaksanakan secara teratur, berkesan dan selaras dengan dasar serta hala tuju MAINPP.")

# ---------- 2.0 ----------
heading("2.0", "STRUKTUR KEAHLIAN")
clause("2.1", "Struktur keanggotaan JKKWPK adalah seperti berikut:")
data_table(
    ["Jawatan", "Penyandang"],
    [
        ["Pengerusi", "YBhg. Timbalan Yang Dipertua, Majlis Agama Islam Negeri Pulau Pinang"],
        ["Setiausaha", "Ketua Penolong Pengarah, Bahagian Undang-Undang Keluarga Islam (BUUKI), Jabatan Hal Ehwal Agama Islam Pulau Pinang (JHEAIPP)"],
        ["Ahli", "Seorang (1) Ahli Majlis Agama Islam Negeri Pulau Pinang yang dilantik oleh MAINPP"],
        ["Ahli", "Pengarah, Lembaga Penduduk dan Pembangunan Keluarga Negara (LPPKN) Negeri Pulau Pinang"],
        ["Ahli", "Pengarah, Jabatan Kebajikan Masyarakat (JKM) Negeri Pulau Pinang"],
        ["Ahli", "Pengarah, Jabatan Kesihatan Negeri Pulau Pinang"],
        ["Ahli", "Ketua Pegawai Eksekutif (CEO), Zakat MAINPP"],
        ["Urus Setia", "Bahagian Undang-Undang Keluarga Islam (BUUKI), JHEAIPP"],
    ],
    [3.5, 10.0],
)
clause("2.2", "Struktur keanggotaan semasa adalah seramai tujuh (7) orang termasuk Pengerusi dan Setiausaha. Walau bagaimanapun, bilangan serta komposisi keanggotaan boleh disemak dan diubah oleh MAINPP dari semasa ke semasa, termasuk mengikut keperluan pelantikan tahunan.")
clause("2.3", "Pelantikan Pengerusi dan ahli-ahli jawatankuasa adalah tertakluk kepada perakuan dan kelulusan MAINPP.")
clause("2.4", "Ahli yang mewakili agensi hendaklah terdiri daripada Pengarah agensi atau Ketua Jabatan yang berkenaan sebagaimana ditetapkan dalam surat pelantikan MAINPP.")
clause("2.5", "Sekiranya ahli yang mewakili agensi tidak dapat hadir, Ketua Jabatan berkenaan boleh menamakan seorang pegawai kanan sebagai wakil melalui surat penamaan yang dikemukakan kepada Urus Setia sebelum mesyuarat. Wakil yang dinamakan dikira dalam kuorum bagi mesyuarat tersebut.")
clause("2.6", "Setiausaha bertanggungjawab mengurus dan menyelaras urusan jawatankuasa, manakala Urus Setia yang terdiri daripada pegawai dan kakitangan BUUKI bertanggungjawab membantu Setiausaha dalam urusan pentadbiran, mesyuarat, dokumentasi, penyelarasan program dan pelaporan.")
clause("2.7", "Tiada jawatan Timbalan Pengerusi diwujudkan dalam struktur keahlian jawatankuasa ini.")

# ---------- 3.0 ----------
heading("3.0", "FUNGSI DAN BIDANG TUGAS JAWATANKUASA")
body("JKKWPK bertanggungjawab untuk:", L1)
for n, t in [
    ("3.1", "Membincangkan dan mencadangkan hala tuju serta perancangan strategik program kebajikan, wanita, kesihatan dan pembangunan keluarga Islam di Negeri Pulau Pinang."),
    ("3.2", "Mengenal pasti keperluan, isu dan cabaran masyarakat Islam dalam aspek kebajikan, kesihatan, wanita dan institusi kekeluargaan."),
    ("3.3", "Meneliti dan mempertimbangkan cadangan program yang dikemukakan oleh agensi-agensi berkaitan sebelum diangkat untuk pertimbangan dan kelulusan MAINPP."),
    ("3.4", "Menggalakkan kerjasama dan perkongsian kepakaran antara MAINPP, jabatan kerajaan, agensi serta pihak berkaitan dalam pelaksanaan program."),
    ("3.5", "Memastikan program yang dirancang mempunyai objektif, kumpulan sasaran, kaedah pelaksanaan dan petunjuk pencapaian yang jelas."),
    ("3.6", "Memantau kemajuan pelaksanaan program serta menilai keberkesanan program berdasarkan laporan dan maklum balas agensi pelaksana."),
    ("3.7", "Mengemukakan syor penambahbaikan bagi meningkatkan keberkesanan penggunaan peruntukan dan impak program kepada masyarakat."),
    ("3.8", "Membincangkan isu-isu berbangkit yang berkaitan dengan bidang tugas jawatankuasa dan mengemukakan cadangan penyelesaian kepada MAINPP."),
]:
    clause(n, t)

# ---------- 4.0 ----------
heading("4.0", "KEKERAPAN MESYUARAT DAN KUORUM")
for n, t in [
    ("4.1", "JKKWPK hendaklah bermesyuarat sekurang-kurangnya dua (2) kali setahun atau mengikut keperluan semasa."),
    ("4.2", "Mesyuarat khas boleh diadakan atas arahan Pengerusi sekiranya terdapat perkara penting yang memerlukan pertimbangan dan tindakan segera."),
    ("4.3", "Notis mesyuarat, agenda serta dokumen berkaitan hendaklah diedarkan oleh Urus Setia kepada ahli-ahli jawatankuasa sekurang-kurangnya tujuh (7) hari sebelum mesyuarat, kecuali bagi mesyuarat khas."),
    ("4.4", "Kuorum mesyuarat hendaklah sekurang-kurangnya separuh daripada jumlah keseluruhan ahli yang dilantik secara rasmi, dengan bilangan dibundarkan ke atas, dan mesti merangkumi Pengerusi atau Setiausaha."),
    ("4.5", "Berdasarkan struktur keanggotaan semasa seramai tujuh (7) orang, kuorum mesyuarat adalah sekurang-kurangnya empat (4) orang."),
    ("4.6", "Sekiranya kuorum tidak mencukupi, mesyuarat boleh diteruskan sebagai sesi perbincangan. Walau bagaimanapun, sebarang keputusan atau syor rasmi hendaklah dibawa untuk pertimbangan dalam mesyuarat yang mempunyai kuorum atau diputuskan secara edaran dengan persetujuan Pengerusi."),
    ("4.7", "Minit mesyuarat hendaklah disediakan oleh Urus Setia dan diedarkan kepada ahli-ahli dalam tempoh empat belas (14) hari selepas mesyuarat untuk makluman serta tindakan susulan."),
]:
    clause(n, t)

# ---------- 5.0 ----------
heading("5.0", "BIDANG KUASA DAN PELAPORAN")
for n, t in [
    ("5.1", "JKKWPK berperanan sebagai jawatankuasa kerja yang membincangkan, meneliti, menyelaras dan mengemukakan syor berkaitan perancangan serta pelaksanaan program di bawah bidang kuasanya."),
    ("5.2", "Sebarang cadangan program, peruntukan kewangan dan komitmen perbelanjaan hendaklah tertakluk kepada pertimbangan serta kelulusan pihak MAINPP mengikut bidang kuasa dan tatacara kewangan yang berkuat kuasa."),
    ("5.3", "JKKWPK tidak mempunyai kuasa autonomi untuk meluluskan perbelanjaan atau membuat komitmen kewangan selain daripada kuasa yang diberikan secara rasmi oleh MAINPP."),
    ("5.4", "Keputusan dan syor jawatankuasa hendaklah diangkat kepada Mesyuarat Penuh MAINPP untuk makluman, pertimbangan atau kelulusan, mengikut kesesuaian perkara."),
    ("5.5", "Urus Setia bertanggungjawab menyediakan laporan keseluruhan pelaksanaan program, status perbelanjaan dan pencapaian jawatankuasa untuk dikemukakan kepada MAINPP sekurang-kurangnya sekali setahun."),
]:
    clause(n, t)

# ---------- 6.0 ----------
heading("6.0", "TEMPOH LANTIKAN")
for n, t in [
    ("6.1", "Tempoh pelantikan Pengerusi dan ahli-ahli JKKWPK adalah selama dua (2) tahun atau mengikut tempoh yang ditetapkan oleh MAINPP."),
    ("6.2", "Pelantikan semula ahli jawatankuasa adalah tertakluk kepada pertimbangan dan kelulusan MAINPP."),
    ("6.3", "Bagi ahli yang dilantik berdasarkan jawatan dalam agensi, kesinambungan keahlian adalah tertakluk kepada kelulusan MAINPP sekiranya berlaku pertukaran atau perubahan penyandang jawatan."),
    ("6.4", "Pelantikan Ahli Majlis sebagai ahli jawatankuasa adalah tertakluk kepada tempoh pelantikan keanggotaan MAINPP, melainkan diputuskan sebaliknya oleh MAINPP."),
    ("6.5", "MAINPP berhak meminda, menamatkan atau memperbaharui pelantikan ahli jawatankuasa mengikut keperluan dan ketetapan semasa."),
]:
    clause(n, t)

# ---------- 7.0 ----------
heading("7.0", "PENGURUSAN PERUNTUKAN, KEWANGAN DAN PELAKSANAAN PROGRAM")
subheading("7.1", "Peruntukan Tahunan")
for n, t in [
    ("7.1.1", "Peruntukan kewangan bagi pelaksanaan program JKKWPK adalah dipohon setiap tahun melalui MAINPP."),
    ("7.1.2", "Setiausaha JKKWPK bertanggungjawab menyediakan kertas cadangan permohonan peruntukan tahunan secara keseluruhan bagi JKKWPK untuk dikemukakan kepada MAINPP bagi pertimbangan dan kelulusan."),
    ("7.1.3", "Permohonan peruntukan hendaklah mengambil kira perancangan program tahunan, keperluan kumpulan sasaran, cadangan agensi, anggaran kos dan keutamaan semasa."),
    ("7.1.4", "Sumber peruntukan adalah daripada Objek Lanjut 734J, Akaun Zakat MAINPP, tertakluk kepada kelulusan, peraturan dan ketetapan kewangan yang berkuat kuasa."),
    ("7.1.5", "Peruntukan JKKWPK adalah di bawah agihan **asnaf fisabilillah**. Setiap program yang dibiayai hendaklah selaras dengan hukum syarak serta ketetapan agihan zakat MAINPP bagi asnaf fisabilillah, dan kertas cadangan hendaklah menyatakan justifikasi program tersebut di bawah asnaf fisabilillah."),
    ("7.1.6", "Agensi pelaksana dan Urus Setia hendaklah memastikan tiada pertindihan manfaat atau pembiayaan dengan program sedia ada di bawah Zakat MAINPP atau agensi lain bagi kumpulan sasaran yang sama."),
]:
    clause(n, t)

subheading("7.2", "Pengurusan Kewangan")
for n, t in [
    ("7.2.1", "Pengurusan kewangan dan pentadbiran peruntukan JKKWPK adalah di bawah tanggungjawab BUUKI selaku Urus Setia jawatankuasa."),
    ("7.2.2", "BUUKI bertanggungjawab mengurus permohonan Local Order (LO) kepada MAINPP, tuntutan bayaran, penyediaan dokumen kewangan serta urusan pembayaran berkaitan program yang diluluskan."),
    ("7.2.3", "Semua urusan perbelanjaan hendaklah mematuhi tatacara kewangan, peraturan perolehan dan prosedur pembayaran yang ditetapkan oleh MAINPP."),
    ("7.2.4", "Setiap perbelanjaan hendaklah disokong dengan dokumen yang lengkap dan dikemukakan untuk semakan serta pengesahan mengikut prosedur yang ditetapkan."),
]:
    clause(n, t)

subheading("7.3", "Penyediaan Kertas Cadangan dan Pelaksanaan Program oleh Agensi")
for n, t in [
    ("7.3.1", "Setiap agensi yang terlibat bertanggungjawab menyediakan kertas cadangan bagi program yang berada di bawah portfolio dan bidang tanggungjawab masing-masing."),
    ("7.3.2", "Kertas cadangan hendaklah mengandungi sekurang-kurangnya objektif, rasional, kumpulan sasaran, justifikasi di bawah asnaf fisabilillah, tarikh, tempat, kaedah pelaksanaan, anggaran perbelanjaan dan hasil yang dijangka."),
    ("7.3.3", "Agensi yang mengemukakan kertas cadangan bertanggungjawab melaksanakan program masing-masing setelah mendapat kelulusan pihak berkuasa yang berkaitan."),
    ("7.3.4", "Agensi pelaksana hendaklah memastikan program dilaksanakan mengikut kertas cadangan dan peruntukan yang telah diluluskan."),
    ("7.3.5", "Agensi pelaksana hendaklah mengemukakan laporan pelaksanaan, pencapaian program dan dokumen berkaitan kepada BUUKI dalam tempoh tiga puluh (30) hari selepas program selesai, bagi tujuan pemantauan serta penyediaan laporan keseluruhan JKKWPK."),
]:
    clause(n, t)

subheading("7.4", "Peranan BUUKI sebagai Penyelaras dan Pemantau")
clause("7.4.1", "BUUKI bertindak sebagai penyelaras dan pemantau keseluruhan program di bawah JKKWPK.")
clause("7.4.2", "Tanggungjawab BUUKI merangkumi:")
letters([
    "Menyelaras perancangan dan takwim program tahunan JKKWPK.",
    "Menyelaras komunikasi dan kerjasama antara agensi pelaksana.",
    "Memantau status kemajuan dan pelaksanaan program.",
    "Mengurus permohonan LO, tuntutan dan dokumentasi kewangan.",
    "Memastikan laporan pelaksanaan program diterima daripada agensi pelaksana.",
    "Menyediakan laporan keseluruhan pelaksanaan program dan perbelanjaan untuk dikemukakan kepada MAINPP.",
])
clause("7.4.3", "Bagi program yang berada di bawah portfolio BUUKI/JHEAIPP, BUUKI bertanggungjawab menyediakan kertas cadangan dan melaksanakan program tersebut sebagai agensi pelaksana.")
clause("7.4.4", "Laporan pelaksanaan program di bawah portfolio BUUKI/JHEAIPP hendaklah dibentangkan dalam mesyuarat JKKWPK untuk semakan dan pengesahan ahli jawatankuasa.")

subheading("7.5", "Pembahagian Portfolio Program")
body("Pembahagian portfolio program mengikut bidang kepakaran dan tanggungjawab agensi adalah seperti berikut:")
data_table(
    ["Agensi", "Bidang Portfolio"],
    [
        ["Lembaga Penduduk dan Pembangunan Keluarga Negara (LPPKN)",
         ["Pendidikan dan pembangunan keluarga.", "Keibubapaan dan kesejahteraan rumah tangga.", "Pendidikan kekeluargaan dan hubungan suami isteri.", "Program berkaitan pembangunan institusi keluarga."]],
        ["Jabatan Kebajikan Masyarakat (JKM)",
         ["Program kebajikan masyarakat.", "Bantuan dan sokongan kepada golongan memerlukan.", "Kebajikan warga emas, orang kurang upaya (OKU) dan golongan rentan.", "Program pemerkasaan keluarga dan komuniti yang memerlukan sokongan kebajikan."]],
        ["Jabatan Kesihatan Negeri Pulau Pinang (JKN)",
         ["Kesedaran dan pendidikan kesihatan.", "Kesihatan keluarga dan komuniti.", "Kesihatan mental dan kesejahteraan emosi.", "Program pencegahan penyakit serta intervensi kesihatan."]],
        ["Jabatan Hal Ehwal Agama Islam Pulau Pinang (JHEAIPP) melalui BUUKI",
         ["Pembangunan keluarga Islam.", "Pendidikan dan pengukuhan institusi kekeluargaan Islam.", "Program praperkahwinan dan pascaperkahwinan Islam.", "Program pengukuhan hubungan kekeluargaan berteraskan ajaran Islam."]],
        ["Zakat MAINPP",
         ["Sokongan dan kerjasama dalam aspek kebajikan serta kesejahteraan golongan asnaf.", "Cadangan program yang bersesuaian dengan fungsi dan bidang kuasa Zakat MAINPP.", "Kerjasama dalam mengenal pasti keperluan kumpulan sasaran yang layak menerima manfaat."]],
    ],
    [4.8, 8.7],
)
clause("7.5.1", "Pembahagian portfolio ini merupakan panduan penyelarasan bidang tanggungjawab dan tidak bermaksud peruntukan kewangan dibahagikan secara sama rata atau secara tetap kepada setiap agensi.")
clause("7.5.2", "Peruntukan bagi setiap program adalah berdasarkan kertas cadangan, keperluan, keutamaan, kemampuan kewangan dan kelulusan MAINPP.")
clause("7.5.3", "Portfolio pembangunan dan pemerkasaan wanita hendaklah diletakkan mengikut salah satu pilihan berikut:")
item("Pilihan A", "Di bawah Jabatan Kebajikan Masyarakat (JKM);", L2, Cm(2.1))
item("Pilihan B", "Di bawah Lembaga Penduduk dan Pembangunan Keluarga Negara (LPPKN); atau", L2, Cm(2.1))
item("Pilihan C", "Dilaksanakan dengan penglibatan pertubuhan bukan kerajaan (NGO) wanita melalui panel atau kumpulan kerja khas di bawah perenggan 8.0.", L2, Cm(2.1))
decision("Mesyuarat diminta memilih satu (1) daripada Pilihan A, B atau C.")

subheading("7.6", "Elaun Mesyuarat")
clause("7.6.1", "Kadar elaun mesyuarat bagi ahli JKKWPK yang dilantik secara rasmi adalah seperti berikut:")
data_table(
    ["Jawatan", "Kadar Elaun (setiap mesyuarat yang dihadiri)"],
    [["Pengerusi", "RM200.00"], ["Ahli jawatankuasa, termasuk Setiausaha", "RM150.00"]],
    [6.5, 7.0],
)
for n, t in [
    ("7.6.2", "Kadar elaun tersebut adalah berdasarkan ketetapan dalam Minit Mesyuarat Majlis Agama Islam Negeri Pulau Pinang Bil. 1 Tahun 2026 (Kali Ke-275)."),
    ("7.6.3", "Elaun mesyuarat dibayar daripada peruntukan JKKWPK di bawah Objek Lanjut 734J, Akaun Zakat MAINPP (asnaf fisabilillah), tertakluk kepada ketetapan agihan zakat dan tatacara kewangan yang berkuat kuasa."),
    ("7.6.4", "Pembayaran elaun adalah berdasarkan kehadiran yang disahkan oleh Urus Setia serta tertakluk kepada kelulusan MAINPP, peruntukan yang tersedia dan tatacara kewangan yang berkuat kuasa."),
    ("7.6.5", "Selain Setiausaha, pegawai Urus Setia, jemputan khas, wakil panel atau pihak luar tidak layak menerima elaun di bawah ketetapan ini melainkan mendapat kelulusan berasingan daripada MAINPP."),
    ("7.6.6", "Kelayakan elaun bagi wakil yang hadir melalui surat penamaan di bawah perenggan 2.5 adalah tertakluk kepada keputusan mesyuarat."),
]:
    clause(n, t)
decision("Mesyuarat diminta memutuskan sama ada wakil yang hadir melalui surat penamaan layak atau tidak layak menerima elaun mesyuarat.")

# ---------- 8.0 ----------
heading("8.0", "PANEL ATAU KUMPULAN KERJA KHAS")
for n, t in [
    ("8.1", "JKKWPK boleh mencadangkan penubuhan panel atau kumpulan kerja khas bagi membantu pelaksanaan sesuatu program atau menangani isu khusus yang memerlukan kepakaran dan penglibatan pihak luar."),
    ("8.2", "Keanggotaan panel atau kumpulan kerja khas boleh melibatkan wakil pertubuhan bukan kerajaan (NGO), pakar bidang, badan profesional atau pihak berkepentingan yang berkaitan."),
    ("8.3", "Pelantikan panel atau kumpulan kerja khas adalah tertakluk kepada pertimbangan dan kelulusan MAINPP."),
    ("8.4", "Panel atau kumpulan kerja khas berperanan memberikan pandangan, kepakaran, cadangan dan sokongan berdasarkan skop tugasan yang ditetapkan."),
    ("8.5", "Keanggotaan panel atau kumpulan kerja khas tidak secara automatik menjadikan individu atau organisasi berkenaan sebagai ahli tetap JKKWPK."),
    ("8.6", "Ahli panel atau kumpulan kerja khas tidak termasuk dalam pengiraan kuorum mesyuarat JKKWPK dan tidak layak menerima elaun mesyuarat ahli jawatankuasa melainkan diluluskan secara berasingan oleh MAINPP."),
]:
    clause(n, t)

# ---------- 9.0 ----------
heading("9.0", "PERANAN DAN TANGGUNGJAWAB")
roles = [
    ("9.1", "Pengerusi", [
        "Mempengerusikan mesyuarat JKKWPK.",
        "Menentukan hala tuju dan keutamaan perbincangan jawatankuasa.",
        "Memberikan panduan serta arahan berkaitan pelaksanaan fungsi jawatankuasa.",
        "Memastikan perbincangan dan keputusan jawatankuasa selaras dengan dasar serta ketetapan MAINPP.",
        "Mengemukakan syor atau perkara penting kepada MAINPP untuk pertimbangan selanjutnya.",
    ]),
    ("9.2", "Setiausaha", [
        "Bertanggungjawab terhadap penyelarasan keseluruhan urusan JKKWPK.",
        "Menyediakan kertas cadangan permohonan peruntukan tahunan keseluruhan JKKWPK untuk dikemukakan kepada MAINPP.",
        "Menyelaras perancangan dan agenda mesyuarat bersama Urus Setia.",
        "Memastikan keputusan mesyuarat diambil tindakan oleh pihak yang berkaitan.",
        "Meneliti dan menyelaras cadangan program yang dikemukakan oleh agensi.",
        "Melaporkan perkembangan dan isu pelaksanaan program kepada Pengerusi serta MAINPP.",
    ]),
    ("9.3", "Ahli Majlis", [
        "Memberikan pandangan, nasihat dan pertimbangan berkaitan dasar, hala tuju serta kepentingan MAINPP dalam perbincangan JKKWPK.",
        "Membantu memastikan cadangan dan syor jawatankuasa selaras dengan dasar, ketetapan serta kepentingan MAINPP.",
        "Memberikan pandangan terhadap cadangan program dan perancangan strategik sebelum dikemukakan kepada MAINPP untuk pertimbangan selanjutnya.",
        "Membantu memperkukuh hubungan penyelarasan antara JKKWPK dengan MAINPP.",
    ]),

    ("9.4", "Ahli Jawatankuasa", [
        "Menghadiri mesyuarat dan memberikan pandangan berdasarkan bidang kepakaran serta tanggungjawab agensi masing-masing.",
        "Mengenal pasti keperluan dan isu berkaitan bidang portfolio masing-masing.",
        "Mengemukakan cadangan program berserta kertas cadangan dan anggaran perbelanjaan.",
        "Memastikan program di bawah tanggungjawab agensi dilaksanakan setelah mendapat kelulusan.",
        "Mengemukakan laporan pelaksanaan dan pencapaian program kepada BUUKI.",
        "Memberikan kerjasama dalam penyelarasan dan pemantauan program JKKWPK.",
    ]),
    ("9.5", "Urus Setia (BUUKI)", [
        "Membantu Setiausaha dalam urusan pentadbiran dan pengurusan JKKWPK.",
        "Mengurus surat-menyurat, notis mesyuarat, agenda, minit mesyuarat dan dokumentasi.",
        "Menyelaras takwim serta pelaksanaan program bersama agensi berkaitan.",
        "Mengurus permohonan LO, tuntutan bayaran dan dokumentasi kewangan kepada MAINPP.",
        "Memantau pelaksanaan program dan mendapatkan laporan daripada agensi pelaksana.",
        "Menyediakan laporan keseluruhan program dan status perbelanjaan JKKWPK.",
        "Menjadi penghubung antara MAINPP dengan agensi-agensi yang terlibat.",
    ]),]
for n, title, items in roles:
    subheading(n, title)
    letters(items)

# ---------- 10.0 ----------
heading("10.0", "KONFLIK KEPENTINGAN DAN KERAHSIAAN MAKLUMAT")
for n, t in [
    ("10.1", "Ahli jawatankuasa hendaklah mengisytiharkan sebarang kepentingan, sama ada secara langsung atau tidak langsung, berhubung sesuatu perkara yang dibincangkan dalam mesyuarat."),
    ("10.2", "Ahli yang mempunyai kepentingan tidak boleh terlibat dalam membuat keputusan atau syor berhubung perkara tersebut, dan pengisytiharan itu hendaklah direkodkan dalam minit mesyuarat."),
    ("10.3", "Semua maklumat peribadi peserta program, penerima manfaat dan golongan asnaf yang diperoleh dalam urusan JKKWPK hendaklah dirahsiakan dan digunakan hanya bagi tujuan pelaksanaan, pemantauan dan pelaporan program."),
]:
    clause(n, t)

# ---------- 11.0 ----------
heading("11.0", "PEMANTAUAN DAN PENILAIAN PRESTASI")
clause("11.1", "Setiap program yang dilaksanakan hendaklah mempunyai objektif dan hasil yang boleh dinilai berdasarkan kumpulan sasaran serta keperluan program.")
clause("11.2", "Agensi pelaksana bertanggungjawab menyediakan laporan pelaksanaan program yang merangkumi maklumat peserta, aktiviti, pencapaian, perbelanjaan dan cadangan penambahbaikan.")
clause("11.3", "BUUKI bertanggungjawab menyelaras dan memantau laporan yang dikemukakan oleh agensi bagi menyediakan penilaian keseluruhan pelaksanaan program JKKWPK.")
clause("11.4", "Penilaian prestasi hendaklah mengambil kira:")
letters([
    "Tahap pencapaian objektif program.",
    "Bilangan dan kesesuaian kumpulan sasaran.",
    "Keberkesanan pelaksanaan program.",
    "Impak program kepada masyarakat.",
    "Pematuhan terhadap peruntukan yang diluluskan.",
    "Keberkesanan penggunaan sumber kewangan.",
    "Cadangan penambahbaikan bagi pelaksanaan program akan datang.",
])
clause("11.5", "Laporan pemantauan dan penilaian keseluruhan hendaklah dikemukakan kepada MAINPP untuk makluman dan pertimbangan selanjutnya.")

# ---------- 12.0 ----------
heading("12.0", "TARIKH KUAT KUASA DAN PINDAAN")
clause("12.1", "Terma Rujukan ini hendaklah dibentangkan dalam Mesyuarat JKKWPK untuk perakuan, dan seterusnya dikemukakan kepada MAINPP untuk kelulusan. Terma Rujukan ini berkuat kuasa mulai tarikh diluluskan oleh MAINPP.")
clause("12.2", "Sebarang pindaan terhadap Terma Rujukan ini hendaklah dibincangkan dalam Mesyuarat JKKWPK dan dikemukakan untuk pertimbangan dan kelulusan MAINPP.")
clause("12.3", "MAINPP berhak menyemak semula struktur, keahlian, fungsi dan bidang kuasa JKKWPK dari semasa ke semasa mengikut keperluan dan hala tuju organisasi.")

# ---------- 13.0 ----------
heading("13.0", "PENUTUP")
body("Terma Rujukan ini disediakan sebagai panduan rasmi bagi memastikan JKKWPK berfungsi secara teratur, sistematik dan berkesan dalam melaksanakan tanggungjawab yang telah ditetapkan.", L1)
body("Kerjasama strategik antara MAINPP, JHEAIPP, BUUKI dan agensi-agensi berkaitan diharapkan dapat memperkukuh pelaksanaan program kebajikan, wanita, kesihatan dan pembangunan keluarga Islam serta memberikan manfaat yang menyeluruh kepada masyarakat Islam di Negeri Pulau Pinang.", L1)

# ---------- pengesahan ----------
sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(16)
sg = doc.add_table(rows=1, cols=2)
sg.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (who, role) in enumerate([
    ("Disediakan oleh:", "Setiausaha\nJawatankuasa Kebajikan, Wanita, Kesihatan\ndan Pembangunan Keluarga (JKKWPK)"),
    ("Diperakukan oleh:", "Pengerusi\nJawatankuasa Kebajikan, Wanita, Kesihatan\ndan Pembangunan Keluarga (JKKWPK)"),
]):
    c = sg.cell(0, i)
    c.width = Cm(7.75)
    cell_borders(c)
    p = c.paragraphs[0]
    p.paragraph_format.keep_with_next = True
    runs(p, who, bold=True, size=10, color=GREEN)
    p = c.add_paragraph(); p.paragraph_format.space_before = Pt(36)
    p.paragraph_format.keep_with_next = True
    runs(p, "……………………………………………", size=10, color=GREY)
    p = c.add_paragraph(); p.paragraph_format.line_spacing = 1.15
    runs(p, role, size=9.5)
    p = c.add_paragraph()
    runs(p, "Tarikh: ……………………", size=9.5)

doc.save("TOR_JKKWPK_Draf_Mesyuarat.docx")
print("ok")
