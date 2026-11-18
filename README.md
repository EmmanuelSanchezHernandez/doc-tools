# doc-tools

**EN:** Two small, dependency-light generators: a JSON spec → styled XLSX builder, and a Markdown → DOCX
converter (headings, lists, pipe tables, code blocks, images, inline emphasis). The converter exists because
`pandoc` crashed with *Illegal instruction* on my server.

---

Dos herramientas para generar entregables sin abrir Office, con **datos separados del código**.

## `xlsx_builder.py` — JSON → Excel con estilo

```bash
python3 xlsx_builder.py spec.json salida.xlsx
```

```json
{
  "sheets": [{
    "name": "Hoja1",
    "freeze_pane": "A2",
    "col_widths": {"A": 15, "B": 30},
    "headers": ["Columna A", "Columna B"],
    "header_color": "1F3864",
    "rows": [{"values": ["x", "y"], "row_color": "EBF3FB", "bold": false}]
  }]
}
```

Hojas múltiples, encabezados con color, filas con color opcional, anchos de columna y panel congelado. Usa
`openpyxl`.

## `md2docx.py` — Markdown → Word

```bash
python3 md2docx.py entrada.md salida.docx "Título del documento"
```

Cubre lo que usan los entregables de consultoría: encabezados, párrafos, listas, tablas de pipes, bloques de
código, imágenes y énfasis en línea (`**negrita**`, `` `código` ``). Usa `python-docx`.

**Por qué existe:** `pandoc` abortaba con *Illegal instruction* en mi VPS (5-ago-2026), así que escribí un
conversor propio que solo depende de Python.

## Dependencias

```bash
pip install openpyxl python-docx
```

## Límites

Es una herramienta pequeña, no un reemplazo de pandoc: no soporta notas al pie, HTML embebido ni plantillas de
marca.

## Licencia

MIT — ver [LICENSE](LICENSE).
