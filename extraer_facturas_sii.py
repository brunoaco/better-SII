#!/usr/bin/env python3
"""
Extractor masivo de datos desde PDFs de "Documento Electrónico Recibido" (SII Chile).

Uso:
    python extraer_facturas_sii.py <carpeta_con_pdfs> [archivo_salida.xlsx]

Ejemplo:
    python extraer_facturas_sii.py "C:\\Users\\Bruno\\Downloads\\SII_Recibidos" reporte.xlsx

Requiere:
    pip install pdfplumber openpyxl pandas
"""

import re
import sys
import glob
import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

ITEM_LINE_RE = re.compile(
    r'^(?P<codigo>\S+)\s+(?P<desc>.+?)\s+(?P<cantidad>\d[\d.,]*(?:\s*[A-Za-z]{1,3})?)\s+'
    r'(?P<precio>[\d.,]+)\s+(?P<valor>[\d.,]+)\s*$'
)
STOP_RE = re.compile(r'^(MONTO|I\.V\.A\.|IMPUESTO ADICIONAL|TOTAL|Timbre|Verifique|Referencias)')

TIPOS_DOC = [
    "GUIA DE DESPACHO ELECTRONICA",
    "NOTA DE CREDITO ELECTRONICA",
    "NOTA DE DEBITO ELECTRONICA",
    "LIQUIDACION FACTURA ELECTRONICA",
    "FACTURA DE COMPRA ELECTRONICA",
    "FACTURA ELECTRONICA",
    "EXENTA ELECTRONICA",
]


def parse_pdf(path):
    """Extrae encabezado + items de un PDF de documento recibido del SII."""
    with pdfplumber.open(path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    lines = text.split("\n")

    header = {"archivo": os.path.basename(path)}

    # Emisor: primer "R.U.T.:" tras el título; la razón social es la línea siguiente
    ruts = [(i, l) for i, l in enumerate(lines) if l.startswith("R.U.T.:")]
    if ruts:
        i0, l0 = ruts[0]
        header["emisor_rut"] = l0.replace("R.U.T.:", "").strip()
        header["emisor_razon_social"] = lines[i0 + 1].strip() if i0 + 1 < len(lines) else ""

    # Giro del emisor (best-effort: puede truncarse si el giro real ocupa 2 líneas)
    for l in lines:
        if l.strip().startswith("Giro:"):
            raw = l.split("Giro:", 1)[1].strip()
            raw = re.split(r'\s+(?:FACTURA|NOTA DE|GUIA DE|LIQUIDACION)\b', raw)[0].strip()
            header["emisor_giro"] = raw
            break

    # Receptor
    m = re.search(r'SEÑOR\(ES\):\s*(.+)', text)
    if m:
        header["receptor_razon_social"] = m.group(1).strip()
    if len(ruts) > 1:
        header["receptor_rut"] = ruts[1][1].replace("R.U.T.:", "").split(" S.I.I.")[0].strip()
    m = re.search(r'GIRO:\s*(.+)', text)
    if m:
        header["receptor_giro"] = m.group(1).split(" S.I.I.")[0].strip()

    # Folio, tipo documento, fecha
    m = re.search(r'Nº\s*(\d+)', text)
    if m:
        header["folio"] = m.group(1)

    m = re.search(r'Fecha Emision:?\s*([\wé\s]+?\d{4})', text)
    if m:
        header["fecha_emision"] = m.group(1).strip()

    flat = text.replace(" ", "").replace("\n", "")
    for tipo in TIPOS_DOC:
        if tipo.replace(" ", "") in flat:
            header["tipo_documento"] = (
                "FACTURA EXENTA/NO AFECTA ELECTRONICA" if tipo == "EXENTA ELECTRONICA" else tipo
            )
            break

    # Totales
    for label, pat in [
        ("monto_neto", r'MONTO NETO\$\s*([\d.,]+)'),
        ("monto_exento", r'MONTO EXENTO\$\s*([\d.,]+)'),
        ("iva", r'I\.V\.A\.[^$]*\$\s*([\d.,]+)'),
        ("impuesto_adicional", r'IMPUESTO ADICIONAL\$\s*([\d.,]+)'),
        ("total", r'TOTAL\$\s*([\d.,]+)'),
    ]:
        m = re.search(pat, text)
        if m:
            header[label] = m.group(1)

    # Items del detalle (una entrada puede envolver varias líneas de descripción;
    # Cantidad/Precio/Valor siempre quedan en la primera línea del ítem, el resto
    # son continuaciones de texto que hay que reagrupar)
    items = []
    tabla_inicio = next(
        (i for i, l in enumerate(lines) if l.strip().startswith("Codigo") and "Descripcion" in l),
        None,
    )
    if tabla_inicio is not None:
        current = None
        for l in lines[tabla_inicio + 1:]:
            s = l.strip()
            if STOP_RE.match(s):
                break
            m = ITEM_LINE_RE.match(s)
            if m:
                if current:
                    items.append(current)
                current = {
                    "codigo": m.group("codigo"),
                    "cantidad": m.group("cantidad"),
                    "precio": m.group("precio"),
                    "valor": m.group("valor"),
                    "_desc_lines": [m.group("desc")],
                }
            elif current and s and s != "Adic.*":
                current["_desc_lines"].append(s)
        if current:
            items.append(current)

    for it in items:
        it["descripcion"] = "\n".join(it.pop("_desc_lines"))
        it["archivo"] = header["archivo"]
        it["folio"] = header.get("folio", "")

    if not items:
        header["_advertencia"] = "No se detectaron líneas de detalle — revisar manualmente"

    return header, items


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial", size=10)


def style_sheet(ws, df):
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        max_len = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str)])
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 45)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
    ws.freeze_panes = "A2"


def main():
    if len(sys.argv) < 2:
        print("Uso: python extraer_facturas_sii.py <carpeta_con_pdfs> [salida.xlsx]")
        sys.exit(1)

    carpeta = sys.argv[1]
    salida = sys.argv[2] if len(sys.argv) > 2 else "extraccion_facturas_sii.xlsx"

    pdfs = sorted(glob.glob(os.path.join(carpeta, "*.pdf")))
    if not pdfs:
        print(f"No se encontraron PDFs en: {carpeta}")
        sys.exit(1)

    print(f"Procesando {len(pdfs)} archivo(s)...")

    headers, all_items = [], []
    errores = []
    for i, path in enumerate(pdfs, 1):
        try:
            h, items = parse_pdf(path)
            headers.append(h)
            all_items.extend(items)
            print(f"  [{i}/{len(pdfs)}] OK  {os.path.basename(path)}")
        except Exception as e:
            errores.append((os.path.basename(path), str(e)))
            print(f"  [{i}/{len(pdfs)}] ERROR  {os.path.basename(path)}: {e}")

    df_docs = pd.DataFrame(headers)
    df_items = pd.DataFrame(all_items)

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_docs.to_excel(writer, sheet_name="Documentos", index=False)
        df_items.to_excel(writer, sheet_name="Detalle", index=False)
        if errores:
            pd.DataFrame(errores, columns=["archivo", "error"]).to_excel(
                writer, sheet_name="Errores", index=False
            )

    wb = load_workbook(salida)
    style_sheet(wb["Documentos"], df_docs)
    style_sheet(wb["Detalle"], df_items)
    wb.save(salida)

    print(f"\nListo. {len(headers)} documento(s), {len(all_items)} línea(s) de detalle.")
    print(f"Errores: {len(errores)}")
    print(f"Archivo generado: {os.path.abspath(salida)}")


if __name__ == "__main__":
    main()
