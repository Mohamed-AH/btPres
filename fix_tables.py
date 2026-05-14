"""Fix COO tables on slides 3 & 4: proper row-span, generous heights, safe margins."""
from pptx import Presentation
from lxml import etree
import openpyxl

PPTX = '/home/user/btPres/Proposal of LC Systems Upgrades v3.pptx'
XLSX = '/home/user/btPres/COO.xlsx'

COLOR_HEADER_BG   = "003D7A"
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_ROW_ALT     = "E8F4EB"
COLOR_BORDER      = "003D7A"
COLOR_TEXT        = "1A1A2E"

# ── Load Excel data ──────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX)
ws = wb['Sheet1']
headers = None
data_rows = []
for i, row in enumerate(ws.iter_rows(values_only=True), 1):
    if i == 1:
        headers = list(row)
    elif any(c is not None for c in row):
        data_rows.append(list(row))

# Column F (index 5) is merged F2:F18; text lives in first row only
coo_text = data_rows[0][5] or ''
# Normalise: remove excessive whitespace but keep all words
coo_text = ' '.join(coo_text.split())

print(f"Headers: {headers}")
print(f"Data rows: {len(data_rows)}")
print(f"COO text ({len(coo_text)} chars): {coo_text[:80]}...")

# ── Helpers ──────────────────────────────────────────────────────────────────
def esc(t):
    return (str(t) if t is not None else '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('\xa0',' ').strip()

def border_xml(color=COLOR_BORDER, w=12700):
    return f'<a:ln w="{w}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>'

def header_cell(text, col_idx, col_widths_emu):
    txt = esc(text)
    return (
        '<a:tc>'
        '<a:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="en-US" sz="900" b="1" dirty="0">'
        '<a:latin typeface="Calibri"/>'
        f'<a:solidFill><a:srgbClr val="{COLOR_HEADER_TEXT}"/></a:solidFill>'
        f'</a:rPr><a:t>{txt}</a:t></a:r>'
        '</a:p></a:txBody>'
        '<a:tcPr marL="76200" marR="76200" marT="38100" marB="38100">'
        f'<a:solidFill><a:srgbClr val="{COLOR_HEADER_BG}"/></a:solidFill>'
        f'<a:lnL>{border_xml()}</a:lnL><a:lnR>{border_xml()}</a:lnR>'
        f'<a:lnT>{border_xml()}</a:lnT><a:lnB>{border_xml()}</a:lnB>'
        '</a:tcPr></a:tc>'
    )

def data_cell(text, col_idx, is_alt, row_span=1, is_vmerge=False):
    """Create a data cell. row_span>1 => vertical merge start. is_vmerge => continuation."""
    if is_vmerge:
        return (
            '<a:tc vMerge="1">'
            '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></a:txBody>'
            '<a:tcPr/></a:tc>'
        )

    txt = esc(text)
    bg = COLOR_ROW_ALT if is_alt else "FFFFFF"
    algn = "ctr" if col_idx in [0] else "l"
    # Use smaller font for long upgrade-criteria text
    sz = "800"
    if col_idx == 2 and len(txt) > 80:
        sz = "700"
    # COO column: centered
    if col_idx == 5:
        algn = "l"
        sz = "700"

    span_attr = f' rowSpan="{row_span}"' if row_span > 1 else ''

    return (
        f'<a:tc{span_attr}>'
        f'<a:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="{algn}"/>'
        f'<a:r><a:rPr lang="en-US" sz="{sz}" dirty="0">'
        '<a:latin typeface="Calibri"/>'
        f'<a:solidFill><a:srgbClr val="{COLOR_TEXT}"/></a:solidFill>'
        f'</a:rPr><a:t>{txt}</a:t></a:r>'
        '</a:p></a:txBody>'
        '<a:tcPr marL="76200" marR="76200" marT="38100" marB="38100">'
        f'<a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>'
        f'<a:lnL>{border_xml()}</a:lnL><a:lnR>{border_xml()}</a:lnR>'
        f'<a:lnT>{border_xml()}</a:lnT><a:lnB>{border_xml()}</a:lnB>'
        '</a:tcPr></a:tc>'
    )

def build_table_xml(rows_data, hdrs, table_w_emu, header_h_emu, row_h_emu, coo_span, has_coo):
    """Build full table XML."""
    # Column proportions: S/N 5%, System 17%, Upgrade Criteria 40%, Brand 12%, Origin 11%, COO 15%
    ratios = [0.05, 0.17, 0.40, 0.12, 0.11, 0.15]
    col_ws = [int(table_w_emu * r) for r in ratios]
    col_ws[-1] = table_w_emu - sum(col_ws[:-1])

    grid = ''.join(f'<a:gridCol w="{w}"/>' for w in col_ws)

    # Header row
    hdr_cells = ''.join(header_cell(h, ci, col_ws) for ci, h in enumerate(hdrs))
    hdr_row = f'<a:tr h="{header_h_emu}">{hdr_cells}</a:tr>'

    # Data rows
    data_xml = ''
    n = len(rows_data)
    for ri, row in enumerate(rows_data):
        is_alt = (ri % 2 == 1)
        cells = ''
        for ci in range(6):
            val = row[ci] if ci < len(row) else None
            if ci == 5:  # COO column
                if has_coo and ri == 0:
                    cells += data_cell(coo_text, ci, False, row_span=n)
                elif has_coo and ri > 0:
                    cells += data_cell('', ci, False, is_vmerge=True)
                else:
                    cells += data_cell('', ci, is_alt)
            else:
                cells += data_cell(val, ci, is_alt)
        data_xml += f'<a:tr h="{row_h_emu}">{cells}</a:tr>'

    return (
        '<a:tbl>'
        '<a:tblPr firstRow="1" bandRow="1">'
        '<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>'
        '</a:tblPr>'
        f'<a:tblGrid>{grid}</a:tblGrid>'
        f'{hdr_row}{data_xml}'
        '</a:tbl>'
    )

def insert_table(slide, label, rows_data, hdrs, has_coo):
    slide_w = 9144000  # 10"
    slide_h = 5143500  # 5.625"

    # Safe positioning with clear margins from all edges
    left  = 914400        # 1.0" from left (after sidebar)
    top   = 685800        # 0.75" from top
    right_margin = 274320 # 0.3" from right edge
    bottom_margin= 182880 # 0.2" from bottom

    table_w = slide_w - left - right_margin  # 8.948"
    max_h   = slide_h - top - bottom_margin  # 4.675"

    n = len(rows_data)
    header_h = 502920  # 0.55"
    row_h = (max_h - header_h) // n

    total_h = header_h + n * row_h

    print(f"  Slide {label}: table {table_w/914400:.3f}\" wide, {total_h/914400:.3f}\" tall")
    print(f"    left={left/914400:.3f}\" top={top/914400:.3f}\" right={(left+table_w)/914400:.3f}\" bottom={(top+total_h)/914400:.3f}\"")
    print(f"    {n} data rows, row_h={row_h/914400:.4f}\"")

    tbl_xml = build_table_xml(rows_data, hdrs, table_w, header_h, row_h, n, has_coo)

    gf_xml = (
        '<p:graphicFrame '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:nvGraphicFramePr>'
        f'<p:cNvPr id="8001" name="COOTable_{label}"/>'
        '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
        '<p:nvPr/>'
        '</p:nvGraphicFramePr>'
        '<p:xfrm>'
        f'<a:off x="{left}" y="{top}"/>'
        f'<a:ext cx="{table_w}" cy="{total_h}"/>'
        '</p:xfrm>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'{tbl_xml}'
        '</a:graphicData>'
        '</a:graphic>'
        '</p:graphicFrame>'
    )

    slide.shapes._spTree.append(etree.fromstring(gf_xml))

# ── Remove old tables then insert new ones ───────────────────────────────────
prs = Presentation(PPTX)

for slide_idx, label, rows, has_coo in [
    (2, 'Slide3', data_rows[:9],  True),
    (3, 'Slide4', data_rows[9:],  False),
]:
    slide = prs.slides[slide_idx]
    sp_tree = slide.shapes._spTree

    # Remove any existing graphicFrame (table) elements
    removed = 0
    for elem in list(sp_tree):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'graphicFrame':
            sp_tree.remove(elem)
            removed += 1
    print(f"Slide {slide_idx+1}: removed {removed} table(s)")

    insert_table(slide, label, rows, headers, has_coo)
    print(f"Slide {slide_idx+1}: new table inserted")

prs.save(PPTX)
print("\nSaved.")
