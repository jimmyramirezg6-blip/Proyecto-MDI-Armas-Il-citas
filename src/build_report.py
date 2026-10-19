from __future__ import annotations

import datetime as dt
import html
import json
import math
import statistics
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPORTS_DIR = ROOT / "reports"
NOTEBOOKS_DIR = ROOT / "notebooks"

HISTORIC_FILE = ROOT / "mdi_armasilicitas_pm_historico_2017_2025.xlsx"
Y2026_FILE = ROOT / "mdi_armasilicitas_pm_2026_enero_septiembre.xlsx"

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EXCEL_EPOCH = dt.datetime(1899, 12, 30)

DISPLAY_MAP = {
    "LOS R�OS": "Los Rios",
    "MANAB�": "Manabi",
    "STO DGO DE LOS TS�CHILAS": "Sto Dgo de los Tsachilas",
    "SANTO DOMINGO DE LOS COLORADOS": "Santo Domingo de los Colorados",
    "�REAS DE ACCESO P�BLICO": "Areas de acceso publico",
    "�REA PRIVADA": "Area privada",
    "VIVIENDA/ALOJAMIENTO": "Vivienda/alojamiento",
    "CASA/VILLA": "Casa/Villa",
    "�REAS P�BLICAS": "Areas publicas",
    "V�A P�BLICA": "Via publica",
    "UNIDADES DE REACCI�N Y EMERGENCIA": "Unidades de reaccion y emergencia",
    "�REAS DEDICADAS AL COMERCIO": "Areas dedicadas al comercio",
    "CENTROS DE DIVERSI�N": "Centros de diversion",
    "ENTIDADES P�BLICAS": "Entidades publicas",
    "CENTRO DE REHABILITACI�N SOCIAL (CRS)": "Centro de rehabilitacion social (CRS)",
    "LUGARES QUE PRESTAN SERVICIOS": "Lugares que prestan servicios",
    "SEMIAUTOM�TICA": "Semiautomatica",
    "AUTOM�TICA": "Automatica",
    "TIRO A TIRO": "Tiro a tiro",
    "SE DESCONOCE": "Se desconoce",
    "NO_APLICA": "No aplica",
    "DELITOS CONTRA LA SEGURIDAD PUBLICA": "Delitos contra la seguridad publica",
    "DELITOS CONTRA EL DERECHO A LA PROPIEDAD": "Delitos contra el derecho a la propiedad",
    "DELITOS POR LA PRODUCCION O TRAFICO ILICITO DE SUSTANCIAS CATALOGADAS SUJETAS A FISCALIZACION": "Delitos por produccion o trafico ilicito de sustancias catalogadas sujetas a fiscalizacion",
    "DELITOS CONTRA LA INVIOLABILIDAD DE LA VIDA": "Delitos contra la inviolabilidad de la vida",
    "DELITOS CONTRA PERSONAS Y BIENES PROTEGIDOS POR EL DERECHO INTERNACIONAL HUMANITARIO": "Delitos contra personas y bienes protegidos por el derecho internacional humanitario",
    "DELITOS CONTRA LA LIBERTAD PERSONAL": "Delitos contra la libertad personal",
    "DELITOS CONTRA LA INTEGRIDAD PERSONAL": "Delitos contra la integridad personal",
    "TENENCIA Y PORTE DE ARMAS": "Tenencia y porte de armas",
    "ARMA DE FUEGO MUNICIONES Y EXPLOSIVOS NO AUTORIZADOS": "Arma de fuego, municiones y explosivos no autorizados",
    "ROBO A PERSONAS": "Robo a personas",
    "TRAFICO ILICITO DE SUSTANCIAS CATALOGADAS SUJETAS A FISCALIZACION": "Trafico ilicito de sustancias catalogadas sujetas a fiscalizacion",
    "PATRULLAJES Y OPERATIVOS": "Patrullajes y operativos",
    "NOVEDADES": "Novedades",
    "OPERATIVO": "Operativo",
}


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
    return n


def cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    value = cell.find("a:v", NS)
    inline = cell.find("a:is", NS)
    if cell_type == "s" and value is not None:
        return shared_strings[int(value.text)] if value.text and value.text.isdigit() else value.text
    if cell_type == "inlineStr" and inline is not None:
        return "".join(t.text or "" for t in inline.iterfind(".//a:t", NS))
    return value.text if value is not None else None


def normalize(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() in {"SIN_DATO", "SIN INFORMACION", "SIN INFORMACIÓN", "N/A", "NA"}:
        return None
    return text


def pretty_label(value):
    if value is None:
        return ""
    value = DISPLAY_MAP.get(str(value), str(value))
    value = value.replace("_", " ")
    value = " ".join(value.split())
    if value.isupper():
        return value.title()
    return value


def fmt_int(value):
    return f"{int(value):,}".replace(",", ".")


def pct(value, digits=1):
    return f"{value:.{digits}f}%"


def excel_date(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        num = float(text.replace(",", "."))
        if num > 0:
            return (EXCEL_EPOCH + dt.timedelta(days=num)).date()
    except ValueError:
        pass
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def load_shared_strings(zf: zipfile.ZipFile):
    shared_strings = []
    if "xl/sharedStrings.xml" not in zf.namelist():
        return shared_strings
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    for item in root.findall("a:si", NS):
        shared_strings.append("".join(t.text or "" for t in item.iterfind(".//a:t", NS)))
    return shared_strings


def pick_main_sheet(zf: zipfile.ZipFile, shared_strings):
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    chosen = None
    for sheet in workbook.findall(".//a:sheets/a:sheet", NS):
        rel_id = sheet.attrib[f"{{{NS['r']}}}id"]
        target = "xl/" + relmap[rel_id]
        xml = ET.fromstring(zf.read(target))
        rows = xml.findall(".//a:sheetData/a:row", NS)
        populated_rows = [row for row in rows if row.findall("a:c", NS)]
        if not populated_rows:
            continue
        if chosen is None or len(rows) > chosen[0]:
            chosen = (len(rows), sheet.attrib["name"], xml, rows)
    return chosen


def extract_records(path: Path):
    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        chosen = pick_main_sheet(zf, shared_strings)
        if chosen is None:
            raise RuntimeError(f"No data sheet found in {path.name}")
        _, sheet_name, xml, rows = chosen

        header_row = next(row for row in rows if row.findall("a:c", NS))
        header_cells = header_row.findall("a:c", NS)
        min_col = min(col_to_num("".join(ch for ch in cell.attrib["r"] if ch.isalpha())) for cell in header_cells)
        headers = [cell_value(cell, shared_strings) for cell in header_cells]
        index = {name: i for i, name in enumerate(headers) if name is not None}

        records = []
        for row in rows[rows.index(header_row) + 1 :]:
            values = [None] * len(headers)
            for cell in row.findall("a:c", NS):
                ref = cell.attrib.get("r")
                if not ref:
                    continue
                col_idx = col_to_num("".join(ch for ch in ref if ch.isalpha())) - min_col
                if 0 <= col_idx < len(values):
                    values[col_idx] = cell_value(cell, shared_strings)
            records.append(values)

    return {
        "path": path,
        "sheet_name": sheet_name,
        "headers": headers,
        "index": index,
        "records": records,
    }


def build_summary(dataset):
    records = dataset["records"]
    idx = dataset["index"]
    total = len(records)

    dates = []
    year_counter = Counter()
    month_counter = Counter()
    province_counter = Counter()
    canton_counter = Counter()
    tipo_delito_counter = Counter()
    delito_counter = Counter()
    tipo_lugar_counter = Counter()
    clase_arma_counter = Counter()
    zona_counter = Counter()
    geo_count = 0
    heatmap = defaultdict(Counter)

    for row in records:
        if "fecha_evento" in idx:
            date = excel_date(row[idx["fecha_evento"]])
            if date:
                dates.append(date)
                year_counter[date.year] += 1
                month_counter[(date.year, date.month)] += 1

        lat = normalize(row[idx["latitud"]]) if "latitud" in idx else None
        lon = normalize(row[idx["longitud"]]) if "longitud" in idx else None
        if lat and lon:
            geo_count += 1

        province = normalize(row[idx["nombre_provincia"]]) if "nombre_provincia" in idx else None
        canton = normalize(row[idx["nombre_canton"]]) if "nombre_canton" in idx else None
        tipo_delito = normalize(row[idx["tipo_delito"]]) if "tipo_delito" in idx else None
        delito = normalize(row[idx["delito"]]) if "delito" in idx else None
        tipo_lugar = normalize(row[idx["tipo_lugar"]]) if "tipo_lugar" in idx else None
        clase_arma = normalize(row[idx["clase_arma"]]) if "clase_arma" in idx else None
        zona = normalize(row[idx["zona"]]) if "zona" in idx else None

        if province:
            province_counter[province] += 1
        if canton:
            canton_counter[canton] += 1
        if tipo_delito:
            tipo_delito_counter[tipo_delito] += 1
        if delito:
            delito_counter[delito] += 1
        if tipo_lugar:
            tipo_lugar_counter[tipo_lugar] += 1
        if clase_arma:
            clase_arma_counter[clase_arma] += 1
        if zona:
            zona_counter[zona] += 1
        if tipo_lugar and clase_arma:
            heatmap[pretty_label(tipo_lugar)][pretty_label(clase_arma)] += 1

    return {
        "total": total,
        "dates": dates,
        "year_counter": year_counter,
        "month_counter": month_counter,
        "province_counter": province_counter,
        "canton_counter": canton_counter,
        "tipo_delito_counter": tipo_delito_counter,
        "delito_counter": delito_counter,
        "tipo_lugar_counter": tipo_lugar_counter,
        "clase_arma_counter": clase_arma_counter,
        "zona_counter": zona_counter,
        "geo_count": geo_count,
        "heatmap": heatmap,
        "date_min": min(dates) if dates else None,
        "date_max": max(dates) if dates else None,
    }


def top_items(counter, n=10):
    return [(pretty_label(label), count) for label, count in counter.most_common(n)]


def color_scale(value, maximum):
    if maximum <= 0:
        return "#edf2f7"
    ratio = max(0.0, min(1.0, value / maximum))
    # Very light blue to dark blue
    light = (237, 242, 247)
    dark = (26, 82, 118)
    rgb = tuple(int(light[i] + (dark[i] - light[i]) * ratio) for i in range(3))
    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def escape(text):
    return html.escape(str(text))


def svg_wrap(content, width, height):
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">{content}</svg>'
    )


def horizontal_bar_chart(title, items, width=1000, bar_height=30, left_margin=280, right_margin=90, top_margin=64, bottom_margin=44, color="#1f77b4"):
    if not items:
        return svg_wrap(f'<text x="20" y="30">{escape(title)} - sin datos</text>', width, 80)
    max_value = max(value for _, value in items)
    chart_h = top_margin + bottom_margin + bar_height * len(items)
    usable_w = width - left_margin - right_margin

    parts = [
        f'<rect x="0" y="0" width="{width}" height="{chart_h}" rx="16" fill="#ffffff" stroke="#d8e1e8"/>',
        f'<text x="24" y="34" font-size="22" font-weight="700" fill="#0f172a">{escape(title)}</text>',
    ]
    for i, (label, value) in enumerate(items):
        y = top_margin + i * bar_height
        bar_w = 0 if max_value == 0 else usable_w * (value / max_value)
        parts.append(f'<text x="24" y="{y + 18}" font-size="13" fill="#1f2937">{escape(label)}</text>')
        parts.append(f'<rect x="{left_margin}" y="{y}" width="{bar_w:.2f}" height="18" rx="8" fill="{color}"/>')
        parts.append(f'<text x="{left_margin + bar_w + 8:.2f}" y="{y + 14}" font-size="12" fill="#0f172a">{fmt_int(value)}</text>')
    return svg_wrap("".join(parts), width, chart_h)


def line_chart(title, labels, values, width=1000, height=360, left_margin=70, right_margin=30, top_margin=50, bottom_margin=60, color="#2563eb"):
    if not values:
        return svg_wrap(f'<text x="20" y="30">{escape(title)} - sin datos</text>', width, 80)
    max_value = max(values)
    min_value = min(values)
    if max_value == min_value:
        max_value += 1
    chart_w = width - left_margin - right_margin
    chart_h = height - top_margin - bottom_margin
    step = chart_w / max(1, len(values) - 1)
    points = []
    area_points = [f"{left_margin},{top_margin + chart_h}"]
    for i, value in enumerate(values):
        x = left_margin + i * step
        y = top_margin + chart_h - ((value - min_value) / (max_value - min_value)) * chart_h
        points.append((x, y))
        area_points.append(f"{x:.2f},{y:.2f}")
    area_points.append(f"{left_margin + (len(values) - 1) * step:.2f},{top_margin + chart_h}")

    grid_lines = []
    for frac in [0.0, 0.25, 0.5, 0.75, 1.0]:
        y = top_margin + chart_h - chart_h * frac
        val = min_value + (max_value - min_value) * frac
        grid_lines.append(f'<line x1="{left_margin}" y1="{y:.2f}" x2="{width - right_margin}" y2="{y:.2f}" stroke="#e5e7eb" stroke-width="1"/>')
        grid_lines.append(f'<text x="12" y="{y + 4:.2f}" font-size="11" fill="#64748b">{fmt_int(round(val))}</text>')

    x_labels = []
    for i, label in enumerate(labels):
        x = left_margin + i * step
        x_labels.append(f'<text x="{x:.2f}" y="{height - 22}" font-size="11" text-anchor="middle" fill="#334155">{escape(label)}</text>')

    points_svg = " ".join(f"{x:.2f},{y:.2f}" for x, y in points)
    circles = "".join(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="4" fill="{color}" />' for x, y in points)
    chart = [
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff" stroke="#d8e1e8"/>',
        f'<text x="24" y="34" font-size="22" font-weight="700" fill="#0f172a">{escape(title)}</text>',
        "".join(grid_lines),
        f'<polygon points="{" ".join(area_points)}" fill="rgba(37, 99, 235, 0.12)"/>',
        f'<polyline points="{points_svg}" fill="none" stroke="{color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>',
        circles,
        "".join(x_labels),
    ]
    return svg_wrap("".join(chart), width, height)


def heatmap_svg(title, heatmap, row_labels, col_labels, width=1000, cell_w=120, cell_h=40, left_margin=300, top_margin=90):
    rows = list(row_labels)
    cols = list(col_labels)
    if not rows or not cols:
        return svg_wrap(f'<text x="20" y="30">{escape(title)} - sin datos</text>', width, 80)

    max_value = 0
    for r in rows:
        for c in cols:
            max_value = max(max_value, heatmap.get(r, {}).get(c, 0))
    height = top_margin + cell_h * len(rows) + 70
    parts = [
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff" stroke="#d8e1e8"/>',
        f'<text x="24" y="34" font-size="22" font-weight="700" fill="#0f172a">{escape(title)}</text>',
        f'<text x="{left_margin}" y="{top_margin - 16}" font-size="13" fill="#475569">Clase de arma</text>',
        f'<text x="24" y="{top_margin - 16}" font-size="13" fill="#475569">Tipo de lugar</text>',
    ]

    for j, col in enumerate(cols):
        x = left_margin + j * cell_w
        parts.append(f'<text x="{x + cell_w / 2:.2f}" y="{top_margin - 2}" font-size="11" text-anchor="middle" fill="#334155">{escape(col)}</text>')

    for i, row in enumerate(rows):
        y = top_margin + i * cell_h
        parts.append(f'<text x="24" y="{y + 25}" font-size="12" fill="#1f2937">{escape(row)}</text>')
        for j, col in enumerate(cols):
            value = heatmap.get(row, {}).get(col, 0)
            x = left_margin + j * cell_w
            fill = color_scale(value, max_value)
            parts.append(f'<rect x="{x}" y="{y}" width="{cell_w - 4}" height="{cell_h - 4}" rx="8" fill="{fill}"/>')
            text_color = "#0f172a" if value < max_value * 0.55 else "#ffffff"
            parts.append(f'<text x="{x + (cell_w - 4) / 2:.2f}" y="{y + 24}" font-size="12" text-anchor="middle" fill="{text_color}">{value}</text>')

    return svg_wrap("".join(parts), width, height)


def build_month_labels(counter):
    month_names = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    items = sorted(counter.items())
    labels = [f"{month_names[m - 1]} {y}" for (y, m), _ in items]
    values = [count for _, count in items]
    return labels, values


def monthly_2026_chart(dataset):
    month_names = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep"]
    items = [(f"{month_names[m - 1]}", dataset["month_counter"][(2026, m)]) for m in range(1, 10)]
    return line_chart("Evolucion mensual 2026", [x for x, _ in items], [y for _, y in items], color="#d97706")


def historical_year_chart(dataset):
    years = sorted(dataset["year_counter"])
    labels = [str(y) for y in years]
    values = [dataset["year_counter"][y] for y in years]
    return line_chart("Evolucion anual historica (2017-2025)", labels, values, color="#2563eb")


def metrics_summary(hist, y2026):
    total_hist = hist["total"]
    total_2026 = y2026["total"]
    avg_month_hist = total_hist / 108
    avg_month_2026 = total_2026 / 9
    growth = (avg_month_2026 / avg_month_hist - 1) * 100

    guayas_hist = hist["province_counter"].get("GUAYAS", 0)
    guayas_2026 = y2026["province_counter"].get("GUAYAS", 0)
    security_hist = hist["tipo_delito_counter"].get("DELITOS CONTRA LA SEGURIDAD PUBLICA", 0)
    security_2026 = y2026["tipo_delito_counter"].get("DELITOS CONTRA LA SEGURIDAD PUBLICA", 0)
    public_place_hist = hist["tipo_lugar_counter"].get("�REAS DE ACCESO P�BLICO", 0)
    public_place_2026 = y2026["tipo_lugar_counter"].get("�REAS DE ACCESO P�BLICO", 0)

    return [
        {
            "label": "Total historico",
            "value": fmt_int(total_hist),
            "note": "2017-2025",
        },
        {
            "label": "Total 2026",
            "value": fmt_int(total_2026),
            "note": "Ene-Sep",
        },
        {
            "label": "Crecimiento mensual 2026",
            "value": f"{growth:+.1f}%",
            "note": f"vs promedio historico ({fmt_int(round(avg_month_hist))}/mes)",
        },
        {
            "label": "Guayas en el historico",
            "value": pct(guayas_hist / total_hist * 100),
            "note": f"{fmt_int(guayas_hist)} casos",
        },
        {
            "label": "Seguridad publica",
            "value": pct(security_hist / total_hist * 100),
            "note": f"{fmt_int(security_hist)} casos historicos",
        },
        {
            "label": "Acceso publico",
            "value": pct(public_place_hist / total_hist * 100),
            "note": f"{fmt_int(public_place_hist)} casos historicos",
        },
        {
            "label": "Cobertura geo historico",
            "value": pct(hist["geo_count"] / total_hist * 100),
            "note": f"{fmt_int(hist['geo_count'])}/{fmt_int(total_hist)} con lat/lon",
        },
    ]


def card_html(metric):
    return f"""
    <div class=\"card\">
      <div class=\"card-label\">{escape(metric['label'])}</div>
      <div class=\"card-value\">{escape(metric['value'])}</div>
      <div class=\"card-note\">{escape(metric['note'])}</div>
    </div>
    """


def make_report(hist, y2026):
    cards = "".join(card_html(m) for m in metrics_summary(hist, y2026))

    province_chart = horizontal_bar_chart("Top 10 provincias (historico)", top_items(hist["province_counter"], 10))
    canton_chart = horizontal_bar_chart("Top 10 cantones (historico)", top_items(hist["canton_counter"], 10))
    delito_chart = horizontal_bar_chart("Top 10 tipos de delito (historico)", top_items(hist["tipo_delito_counter"], 10))
    lugar_chart = horizontal_bar_chart("Top 10 tipos de lugar (historico)", top_items(hist["tipo_lugar_counter"], 10))
    year_chart = historical_year_chart(hist)
    month_chart = monthly_2026_chart(y2026)

    top_rows = [pretty_label(label) for label, _ in hist["tipo_lugar_counter"].most_common(5)]
    top_cols = [pretty_label(label) for label, _ in hist["clase_arma_counter"].most_common(5)]
    heatmap = heatmap_svg("Heatmap: tipo de lugar vs clase de arma", hist["heatmap"], top_rows, top_cols)

    historico_total = hist["total"]
    y2026_total = y2026["total"]
    avg_month_hist = historico_total / 108
    avg_month_2026 = y2026_total / 9
    growth = (avg_month_2026 / avg_month_hist - 1) * 100
    guayas_share = hist["province_counter"].get("GUAYAS", 0) / historico_total * 100
    security_share = hist["tipo_delito_counter"].get("DELITOS CONTRA LA SEGURIDAD PUBLICA", 0) / historico_total * 100
    public_share = hist["tipo_lugar_counter"].get("�REAS DE ACCESO P�BLICO", 0) / historico_total * 100
    geo_cov = hist["geo_count"] / historico_total * 100

    summary_text = f"""
    <p><strong>Hallazgo 1.</strong> El problema esta muy concentrado: Guayas representa {pct(guayas_share)} del historico y Guayaquil lidera el ranking cantonal.</p>
    <p><strong>Hallazgo 2.</strong> El volumen no se enfria: 2026 promedia {fmt_int(round(avg_month_2026))} casos por mes, un {growth:+.1f}% frente al promedio historico mensual.</p>
    <p><strong>Hallazgo 3.</strong> La categoria dominante es <em>Delitos contra la seguridad publica</em>, con {pct(security_share)} del historico.</p>
    <p><strong>Hallazgo 4.</strong> El contexto operativo esta concentrado en <em>Areas de acceso publico</em> ({pct(public_share)} del historico), con una cobertura geografica de {pct(geo_cov)} en el historico.</p>
    """

    html_doc = f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MDI Armas Ilícitas - Reporte final</title>
  <style>
    body {{
      margin: 0;
      font-family: Inter, Segoe UI, Arial, sans-serif;
      background: #f5f7fb;
      color: #0f172a;
    }}
    .wrap {{ max-width: 1200px; margin: 0 auto; padding: 32px 20px 60px; }}
    .hero {{ background: linear-gradient(135deg, #0f172a, #1d4ed8); color: white; border-radius: 22px; padding: 28px 30px; box-shadow: 0 18px 50px rgba(15,23,42,.16); }}
    .hero h1 {{ margin: 0 0 10px; font-size: 34px; }}
    .hero p {{ margin: 0; max-width: 900px; line-height: 1.55; color: rgba(255,255,255,.9); }}
    .section {{ margin-top: 24px; background: white; border-radius: 20px; padding: 22px; box-shadow: 0 10px 30px rgba(15,23,42,.07); }}
    .section h2 {{ margin: 0 0 16px; font-size: 24px; }}
    .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; }}
    .card {{ background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 18px; padding: 16px; }}
    .card-label {{ font-size: 12px; letter-spacing: .02em; text-transform: uppercase; color: #64748b; margin-bottom: 10px; }}
    .card-value {{ font-size: 28px; font-weight: 800; color: #0f172a; margin-bottom: 6px; }}
    .card-note {{ font-size: 13px; color: #475569; line-height: 1.4; }}
    .grid-2 {{ display: grid; grid-template-columns: 1fr; gap: 18px; }}
    .chart {{ overflow-x: auto; }}
    .note {{ color: #475569; line-height: 1.6; }}
    .bullets li {{ margin-bottom: 10px; line-height: 1.5; }}
    .meta-table {{ width: 100%; border-collapse: collapse; }}
    .meta-table td {{ border-bottom: 1px solid #e2e8f0; padding: 12px 8px; }}
    .meta-table td:first-child {{ font-weight: 700; color: #334155; width: 240px; }}
    .footer {{ margin-top: 18px; color: #64748b; font-size: 13px; }}
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>MDI Armas Ilícitas - Reporte final</h1>
      <p>Análisis ejecutivo de registros de armas ilícitas y delitos asociados en Ecuador. El foco es identificar concentración geográfica, tendencia temporal, tipología del hecho y escenarios operativos más frecuentes.</p>
    </section>

    <section class="section">
      <h2>5 métricas clave</h2>
      <div class="cards">{cards}</div>
    </section>

    <section class="section">
      <h2>Hallazgos contundentes</h2>
      <div class="note">{summary_text}</div>
      <ul class="bullets">
        <li><strong>Guayas domina el mapa</strong>: concentra la mayor parte de los registros y concentra el problema en una sola provincia.</li>
        <li><strong>2026 viene acelerado</strong>: el promedio mensual supera ampliamente la media historica.</li>
        <li><strong>La lectura operativa es clara</strong>: el fenomeno ocurre en escenarios de acceso publico y no de forma dispersa.</li>
        <li><strong>La categoria dominante es estable</strong>: seguridad publica sigue marcando el peso principal del conjunto.</li>
      </ul>
    </section>

    <section class="section"><h2>Evolucion temporal</h2><div class="chart">{year_chart}</div></section>
    <section class="section"><h2>Evolucion mensual 2026</h2><div class="chart">{month_chart}</div></section>
    <section class="section"><h2>Geografia</h2><div class="chart">{province_chart}</div></section>
    <section class="section"><h2>Detalle territorial</h2><div class="chart">{canton_chart}</div></section>
    <section class="section"><h2>Tipologia del hecho</h2><div class="chart">{delito_chart}</div></section>
    <section class="section"><h2>Escenario operativo</h2><div class="chart">{lugar_chart}</div></section>
    <section class="section"><h2>Interaccion tipo de lugar vs clase de arma</h2><div class="chart">{heatmap}</div></section>

    <section class="section">
      <h2>Metadatos de calidad</h2>
      <table class="meta-table">
        <tr><td>Historico</td><td>{escape(str(hist['date_min']))} a {escape(str(hist['date_max']))}</td></tr>
        <tr><td>2026</td><td>{escape(str(y2026['date_min']))} a {escape(str(y2026['date_max']))}</td></tr>
        <tr><td>Cobertura geográfica historica</td><td>{pct(geo_cov)}</td></tr>
        <tr><td>Casos con coordenadas</td><td>{fmt_int(hist['geo_count'])} de {fmt_int(hist['total'])}</td></tr>
      </table>
    </section>

    <div class="footer">Generado automaticamente por <code>src/build_report.py</code>.</div>
  </div>
</body>
</html>"""

    return html_doc


def build_notebook(hist, y2026):
    year_chart = historical_year_chart(hist)
    month_chart = monthly_2026_chart(y2026)
    province_chart = horizontal_bar_chart("Top 10 provincias (historico)", top_items(hist["province_counter"], 10))
    canton_chart = horizontal_bar_chart("Top 10 cantones (historico)", top_items(hist["canton_counter"], 10))
    delito_chart = horizontal_bar_chart("Top 10 tipos de delito (historico)", top_items(hist["tipo_delito_counter"], 10))
    lugar_chart = horizontal_bar_chart("Top 10 tipos de lugar (historico)", top_items(hist["tipo_lugar_counter"], 10))
    top_rows = [pretty_label(label) for label, _ in hist["tipo_lugar_counter"].most_common(5)]
    top_cols = [pretty_label(label) for label, _ in hist["clase_arma_counter"].most_common(5)]
    heatmap = heatmap_svg("Heatmap: tipo de lugar vs clase de arma", hist["heatmap"], top_rows, top_cols)

    cells = [
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "# MDI Armas Ilícitas\n",
                "\n",
                "Notebook ejecutivo con gráficos clave del proyecto. Los gráficos están embebidos para consulta rapida.\n",
            ],
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "## Hallazgos clave\n",
                "\n",
                f"- Total historico: {fmt_int(hist['total'])}\n",
                f"- Total 2026 (ene-sep): {fmt_int(y2026['total'])}\n",
                f"- Guayas concentra {pct(hist['province_counter'].get('GUAYAS', 0) / hist['total'] * 100)} del historico\n",
                f"- Delitos contra la seguridad publica concentran {pct(hist['tipo_delito_counter'].get('DELITOS CONTRA LA SEGURIDAD PUBLICA', 0) / hist['total'] * 100)}\n",
                f"- Areas de acceso publico concentran {pct(hist['tipo_lugar_counter'].get('�REAS DE ACCESO P�BLICO', 0) / hist['total'] * 100)}\n",
            ],
        },
    ]

    def code_cell(title, svg):
        return {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [
                {
                    "output_type": "display_data",
                    "metadata": {},
                    "data": {"text/html": svg},
                }
            ],
            "source": [f"# {title}\n", "# Gráfico embebido generado por src/build_report.py\n"],
        }

    cells.extend(
        [
            code_cell("Evolucion anual historica", year_chart),
            code_cell("Evolucion mensual 2026", month_chart),
            code_cell("Top provincias", province_chart),
            code_cell("Top cantones", canton_chart),
            code_cell("Top tipos de delito", delito_chart),
            code_cell("Top tipos de lugar", lugar_chart),
            code_cell("Heatmap tipo de lugar vs clase de arma", heatmap),
        ]
    )

    cells.append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "## Conclusión\n",
                "\n",
                "El problema esta concentrado geograficamente, domina en Guayas/Guayaquil y mantiene una señal fuerte en 2026.\n",
            ],
        }
    )

    notebook = {
        "cells": cells,
        "metadata": {
            "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
            "language_info": {"name": "python", "version": "3.14"},
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }
    return json.dumps(notebook, ensure_ascii=False, indent=2)


def main():
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    NOTEBOOKS_DIR.mkdir(parents=True, exist_ok=True)

    historic = extract_records(HISTORIC_FILE)
    y2026 = extract_records(Y2026_FILE)
    hist_summary = build_summary(historic)
    y2026_summary = build_summary(y2026)

    report_html = make_report(hist_summary, y2026_summary)
    notebook_json = build_notebook(hist_summary, y2026_summary)

    (REPORTS_DIR / "reporte_final.html").write_text(report_html, encoding="utf-8")
    (NOTEBOOKS_DIR / "01_analisis_armas_ilicitas.ipynb").write_text(notebook_json, encoding="utf-8")

    print("Generated:")
    print(f"- {REPORTS_DIR / 'reporte_final.html'}")
    print(f"- {NOTEBOOKS_DIR / '01_analisis_armas_ilicitas.ipynb'}")


if __name__ == "__main__":
    main()
