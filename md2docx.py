#!/usr/bin/env python3
"""md2docx.py — Markdown -> DOCX con python-docx.

Existe porque pandoc aborta con "Illegal instruction" en este VPS (5-ago-2026).
Cubre lo que usan los entregables de consultoría: encabezados, parrafos, listas, tablas
de pipes, bloques de codigo, imagenes y enfasis en linea (**negrita**, `codigo`).

Uso:  python3 md2docx.py entrada.md salida.docx ["Titulo del documento"]
"""
import re, sys, os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

INLINE = re.compile(r'(\*\*.+?\*\*|`[^`]+`|\*[^*]+\*)')

def add_runs(par, text):
    """Reparte el texto en runs aplicando negrita, cursiva y monoespaciado."""
    for tok in INLINE.split(text):
        if not tok:
            continue
        if tok.startswith('**') and tok.endswith('**') and len(tok) > 4:
            par.add_run(tok[2:-2]).bold = True
        elif tok.startswith('`') and tok.endswith('`') and len(tok) > 2:
            r = par.add_run(tok[1:-1]); r.font.name = 'Consolas'; r.font.size = Pt(9)
        elif tok.startswith('*') and tok.endswith('*') and len(tok) > 2:
            par.add_run(tok[1:-1]).italic = True
        else:
            par.add_run(tok)

def split_row(line):
    """Parte una fila de tabla respetando los \\| escapados."""
    body = line.strip().strip('|')
    parts = re.split(r'(?<!\\)\|', body)
    return [p.strip().replace('\\|', '|') for p in parts]

def is_sep(line):
    return bool(re.match(r'^\s*\|?[\s:\-|]+\|[\s:\-|]*$', line)) and '-' in line

def convert(md_path, docx_path, title=None):
    lines = open(md_path, encoding='utf-8').read().split('\n')
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10.5)
    if title:
        doc.add_heading(title, 0)

    i, n = 0, len(lines)
    while i < n:
        ln = lines[i]
        s = ln.strip()

        if not s or s == '---':
            i += 1; continue

        # bloque de codigo
        if s.startswith('```'):
            i += 1; buf = []
            while i < n and not lines[i].strip().startswith('```'):
                buf.append(lines[i]); i += 1
            i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(8)
            r = p.add_run('\n'.join(buf))
            r.font.name = 'Consolas'; r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
            continue

        # imagen
        m = re.match(r'^!\[(.*?)\]\((.+?)\)(\{.*\})?$', s)
        if m:
            cap, path = m.group(1), m.group(2)
            if os.path.exists(path):
                doc.add_picture(path, width=Inches(6.3))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if cap:
                    c = doc.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = c.add_run(cap); r.italic = True; r.font.size = Pt(9)
                    r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
            else:
                p = doc.add_paragraph()
                r = p.add_run('[FALTA LA CAPTURA: %s]' % (cap or path))
                r.italic = True; r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            i += 1; continue

        # encabezados
        m = re.match(r'^(#{1,4})\s+(.*)$', s)
        if m:
            doc.add_heading(m.group(2).strip(), min(len(m.group(1)), 4))
            i += 1; continue

        # tabla de pipes: fila + separador
        if s.startswith('|') and i + 1 < n and is_sep(lines[i + 1]):
            head = split_row(s); i += 2
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                rows.append(split_row(lines[i])); i += 1
            t = doc.add_table(rows=1, cols=len(head))
            t.style = 'Light Grid Accent 1'
            for c, txt in enumerate(head):
                cell = t.rows[0].cells[c]; cell.text = ''
                add_runs(cell.paragraphs[0], txt)
                for r_ in cell.paragraphs[0].runs: r_.bold = True
            for row in rows:
                cells = t.add_row().cells
                for c in range(len(head)):
                    cells[c].text = ''
                    add_runs(cells[c].paragraphs[0], row[c] if c < len(row) else '')
            doc.add_paragraph()
            continue

        # listas
        m = re.match(r'^(\s*)([-*]|\d+\.)\s+(.*)$', ln)
        if m:
            depth = len(m.group(1)) // 2
            style = 'List Number' if m.group(2)[0].isdigit() else 'List Bullet'
            if depth: style += ' %d' % min(depth + 1, 3)
            p = doc.add_paragraph(style=style)
            add_runs(p, m.group(3))
            i += 1; continue

        # parrafo: junta lineas hasta el siguiente bloque
        buf = []
        while i < n and lines[i].strip() and not re.match(r'^(#{1,4}\s|\||```|!\[|\s*([-*]|\d+\.)\s)', lines[i]) and lines[i].strip() != '---':
            buf.append(lines[i].strip()); i += 1
        if buf:
            add_runs(doc.add_paragraph(), ' '.join(buf))

    doc.save(docx_path)
    return docx_path

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    out = convert(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    print('OK %s (%d bytes)' % (out, os.path.getsize(out)))
