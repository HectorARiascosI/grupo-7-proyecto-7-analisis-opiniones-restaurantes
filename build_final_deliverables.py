"""Build the Grupo 7 report, presentation, and static HTML companion."""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import re
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import markdown
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor as DocxColor
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as PptxInches, Pt as PptxPt
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import HRFlowable, Paragraph, Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
COLORS = {
    "navy": "17324D",
    "teal": "148A87",
    "coral": "D76A4A",
    "gold": "E3B44B",
    "ink": "243447",
    "muted": "617487",
    "paper": "F5F7FA",
    "white": "FFFFFF",
    "line": "D9E1E8",
    "blue": "4B83B6",
}


def _read_business_lookup(path: Path) -> tuple[dict[str, dict[str, str]], dict[str, str]]:
    if not path.is_file():
        raise FileNotFoundError(f"No se encontró la dimensión de negocios: {path}")
    records: dict[str, dict[str, str]] = {}
    city_labels: dict[str, str] = {}
    with path.open(encoding="utf-8-sig", newline="") as source:
        for row in csv.DictReader(source):
            business_id = row.get("business_id", "").strip()
            if not business_id:
                continue
            raw_city = row.get("city", "").strip()
            city = raw_city or "Sin ciudad"
            city_key = raw_city.casefold()
            price = row.get("price_range", "").strip()
            if not price or price.casefold() in {"nan", "none", "null"}:
                price = "Sin dato"
            else:
                try:
                    price = str(int(float(price)))
                except ValueError:
                    price = "Sin dato"
            records[business_id] = {
                "name": row.get("business_name", "").strip() or business_id,
                "city": city,
                "city_key": city_key,
                "price": price,
            }
            if city_key:
                city_labels.setdefault(city_key, city)
    if not records:
        raise ValueError(f"La dimensión de negocios está vacía: {path}")
    return records, city_labels


def _zero_stats() -> dict[str, float]:
    return {
        "n": 0,
        "stars_sum": 0.0,
        "stars_sq": 0.0,
        "positive": 0,
        "negative": 0,
        "neutral": 0,
        "useful": 0.0,
        "complaint_reviews": 0,
    }


def _merge_group_stats(
    target: dict[str, dict[str, float]], grouped: pd.DataFrame
) -> None:
    for key, row in grouped.iterrows():
        label = "Sin etiqueta" if pd.isna(key) or not str(key).strip() else str(key)
        current = target.setdefault(label, _zero_stats())
        for field in current:
            current[field] += float(row[field])


def _summary_row(key: str, stats: dict[str, float]) -> dict[str, Any]:
    n = int(stats["n"])
    average = stats["stars_sum"] / n if n else 0.0
    variance = max(0.0, stats["stars_sq"] / n - average * average) if n else 0.0
    return {
        "label": key,
        "reviews": n,
        "average_stars": round(average, 4),
        "positive_share": round(stats["positive"] / n, 6) if n else 0.0,
        "negative_share": round(stats["negative"] / n, 6) if n else 0.0,
        "neutral_share": round(stats["neutral"] / n, 6) if n else 0.0,
        "useful_votes": int(stats["useful"]),
        "complaint_reviews": int(stats["complaint_reviews"]),
        "population_stddev": round(math.sqrt(variance), 4),
    }


def _read_profile_summary(path: Path) -> list[dict[str, Any]]:
    totals: dict[str, dict[str, float]] = {}
    with path.open(encoding="utf-8-sig", newline="") as source:
        for row in csv.DictReader(source):
            segment = row.get("activity_segment", "").strip() or "Sin segmento"
            values = totals.setdefault(
                segment,
                {"users": 0, "review_count": 0.0, "average_stars": 0.0, "fans": 0.0},
            )
            values["users"] += 1
            for field in ("review_count", "average_stars", "fans"):
                try:
                    values[field] += float(row.get(field) or 0)
                except ValueError:
                    continue
    return [
        {
            "segment": segment,
            "users": int(values["users"]),
            "average_review_count": round(values["review_count"] / values["users"], 2),
            "average_stars": round(values["average_stars"] / values["users"], 2),
            "average_fans": round(values["fans"] / values["users"], 2),
        }
        for segment, values in totals.items()
    ]


def _read_small_csv(path: Path, limit: int | None = None) -> list[dict[str, str]]:
    if not path.is_file():
        raise FileNotFoundError(f"No se encontró el archivo requerido: {path}")
    with path.open(encoding="utf-8-sig", newline="") as source:
        rows = list(csv.DictReader(source))
    return rows if limit is None else rows[:limit]


def build_summary(data_dir: Path, chunk_size: int = 200_000) -> dict[str, Any]:
    """Aggregate the compact fact table without retaining review text."""
    business, city_labels = _read_business_lookup(data_dir / "dim_business.csv")
    business_city_keys = {
        business_id: record["city_key"] for business_id, record in business.items()
    }
    fact_path = data_dir / "fact_review.csv"
    if not fact_path.is_file():
        raise FileNotFoundError(f"No se encontró la tabla compacta: {fact_path}")

    category_stats: dict[str, dict[str, float]] = {}
    city_stats: dict[str, dict[str, float]] = {}
    price_stats: dict[str, dict[str, float]] = {}
    month_stats: dict[str, dict[str, float]] = {}
    business_stats: dict[str, list[float]] = {}
    sentiment_counts = {"Positive": 0, "Neutral": 0, "Negative": 0}
    total_reviews = 0
    stars_sum = 0.0
    useful_sum = 0.0

    columns = [
        "business_id",
        "review_day",
        "stars",
        "useful",
        "sentiment_label",
        "category",
        "complaint_terms",
    ]
    chunks = pd.read_csv(
        fact_path,
        usecols=columns,
        dtype={
            "business_id": "string",
            "review_day": "string",
            "sentiment_label": "string",
            "category": "string",
            "complaint_terms": "string",
        },
        chunksize=chunk_size,
        encoding="utf-8-sig",
        low_memory=False,
    )
    for chunk_number, chunk in enumerate(chunks, 1):
        chunk["stars"] = pd.to_numeric(chunk["stars"], errors="coerce")
        chunk = chunk.dropna(subset=["stars"])
        if chunk.empty:
            continue
        chunk["useful"] = pd.to_numeric(chunk["useful"], errors="coerce").fillna(0)
        chunk["stars_sq"] = chunk["stars"] * chunk["stars"]
        chunk["positive"] = chunk["sentiment_label"].eq("Positive").astype("int8")
        chunk["negative"] = chunk["sentiment_label"].eq("Negative").astype("int8")
        chunk["neutral"] = chunk["sentiment_label"].eq("Neutral").astype("int8")
        chunk["has_complaint"] = (
            chunk["complaint_terms"].fillna("").str.strip().ne("").astype("int8")
        )
        chunk["category_key"] = chunk["category"].fillna("").str.strip()
        chunk.loc[chunk["category_key"].eq(""), "category_key"] = "Sin etiqueta"
        chunk["city"] = chunk["business_id"].map(
            {business_id: record["city"] for business_id, record in business.items()}
        ).fillna("Sin ciudad")
        chunk["city_key"] = chunk["business_id"].map(business_city_keys).fillna("")
        chunk["price"] = chunk["business_id"].map(
            {business_id: record["price"] for business_id, record in business.items()}
        ).fillna("Sin dato")
        chunk["month"] = chunk["review_day"].fillna("").str.slice(0, 7)
        chunk.loc[chunk["month"].eq(""), "month"] = "Sin fecha"

        aggregate_columns = {
            "n": ("stars", "count"),
            "stars_sum": ("stars", "sum"),
            "stars_sq": ("stars_sq", "sum"),
            "positive": ("positive", "sum"),
            "negative": ("negative", "sum"),
            "neutral": ("neutral", "sum"),
            "useful": ("useful", "sum"),
            "complaint_reviews": ("has_complaint", "sum"),
        }
        for key_column, target in (
            ("category_key", category_stats),
            ("city_key", city_stats),
            ("price", price_stats),
            ("month", month_stats),
        ):
            grouped = chunk.groupby(key_column, dropna=False).agg(**aggregate_columns)
            _merge_group_stats(target, grouped)

        grouped_businesses = chunk.groupby("business_id").agg(
            n=("stars", "count"), stars_sum=("stars", "sum")
        )
        for business_id, row in grouped_businesses.iterrows():
            totals = business_stats.setdefault(str(business_id), [0.0, 0.0])
            totals[0] += float(row["n"])
            totals[1] += float(row["stars_sum"])

        total_reviews += len(chunk)
        stars_sum += float(chunk["stars"].sum())
        useful_sum += float(chunk["useful"].sum())
        for label in sentiment_counts:
            sentiment_counts[label] += int(chunk["sentiment_label"].eq(label).sum())
        print(f"Agregadas {total_reviews:,} reseñas...", flush=True)

    if not total_reviews:
        raise ValueError("La tabla compacta no contiene reseñas con estrellas.")

    category_rows = [
        _summary_row(label, stats) for label, stats in category_stats.items()
    ]
    city_rows = [
        _summary_row(city_labels.get(key, "Sin ciudad"), stats)
        for key, stats in city_stats.items()
    ]
    price_rows = [
        _summary_row(label, stats) for label, stats in price_stats.items()
    ]
    month_rows = [
        {
            "month": label,
            "reviews": int(stats["n"]),
            "average_stars": round(stats["stars_sum"] / stats["n"], 4),
            "positive_share": round(stats["positive"] / stats["n"], 6),
        }
        for label, stats in month_stats.items()
    ]
    month_rows.sort(key=lambda row: row["month"])
    volume_rows = []
    rating_rows = []
    for business_id, (count, stars_total) in business_stats.items():
        record = business.get(business_id, {"name": business_id, "city": "Sin ciudad"})
        entry = {
            "name": record["name"],
            "city": record["city"],
            "reviews": int(count),
            "average_stars": round(stars_total / count, 4),
        }
        volume_rows.append(entry)
        if count >= 100:
            rating_rows.append(entry)
    volume_rows.sort(key=lambda row: (-row["reviews"], row["name"].casefold()))
    rating_rows.sort(
        key=lambda row: (-row["average_stars"], -row["reviews"], row["name"].casefold())
    )
    city_rows.sort(key=lambda row: (-row["reviews"], row["label"]))
    price_order = {"1": 1, "2": 2, "3": 3, "4": 4, "Sin dato": 5}
    price_rows.sort(key=lambda row: price_order.get(row["label"], 6))
    category_by_volume = sorted(
        category_rows, key=lambda row: (-row["reviews"], row["label"])
    )
    category_by_positive = sorted(
        (row for row in category_rows if row["reviews"] >= 100),
        key=lambda row: (-row["positive_share"], -row["reviews"], row["label"]),
    )
    category_by_deviation = sorted(
        (row for row in category_rows if row["reviews"] >= 100),
        key=lambda row: (-row["population_stddev"], row["label"]),
    )
    complaint_rows = _read_small_csv(data_dir / "complaint_term_counts.csv")
    keyword_rows = _read_small_csv(data_dir / "review_keyword_counts.csv", 20)
    profiles = _read_profile_summary(data_dir / "dim_user.csv")

    return {
        "meta": {
            "reviews": total_reviews,
            "average_stars": round(stars_sum / total_reviews, 3),
            "useful_votes": int(useful_sum),
            "businesses": len(business),
            "cities": sum(1 for key in city_stats if key in city_labels),
            "sentiment": sentiment_counts,
            "positive_share": round(sentiment_counts["Positive"] / total_reviews, 6),
            "negative_share": round(sentiment_counts["Negative"] / total_reviews, 6),
            "neutral_share": round(sentiment_counts["Neutral"] / total_reviews, 6),
            "minimum_reviews_for_rating_top10": 100,
        },
        "price": price_rows,
        "cities_top": city_rows[:10],
        "monthly": month_rows,
        "category_volume": category_by_volume[:12],
        "category_positive_min100": category_by_positive[:12],
        "category_deviation_min100": category_by_deviation[:10],
        "complaints": complaint_rows,
        "keywords": keyword_rows,
        "top_volume": volume_rows[:10],
        "top_rating_min100": rating_rows[:10],
        "profiles": profiles,
    }


def build_dashboard(summary: dict[str, Any], output: Path) -> None:
    payload = json.dumps(summary, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("<", "\\u003c")
    page = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Dashboard estático agregado del análisis Yelp del Grupo 7.">
<title>Grupo 7 | Opiniones de restaurantes</title>
<style>
:root{--navy:#17324d;--teal:#148a87;--coral:#d76a4a;--gold:#e3b44b;--ink:#243447;--muted:#617487;--paper:#f5f7fa;--line:#d9e1e8;--white:#fff}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:16px/1.5 "Segoe UI",Arial,sans-serif}
.wrap{max-width:1180px;margin:auto;padding:28px 20px 60px}.hero{background:linear-gradient(125deg,#17324d,#245a70);color:#fff;padding:34px;border-radius:20px}
.eyebrow{text-transform:uppercase;letter-spacing:.14em;font-size:.78rem;color:#b8d7dc;font-weight:700}.hero h1{font-size:clamp(2rem,4vw,3rem);line-height:1.08;margin:.45rem 0}.hero p{max-width:760px;color:#e0eaef;margin:.6rem 0 0}
h2{color:var(--navy);margin:34px 0 14px;font-size:1.45rem}.sub{color:var(--muted);margin:-8px 0 16px}
.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}.card,.panel{background:#fff;border:1px solid var(--line);border-radius:14px;padding:18px;box-shadow:0 4px 14px #18344d0a}
.kpi-label{color:var(--muted);font-size:.9rem}.kpi-value{font-size:1.8rem;font-weight:750;color:var(--navy);margin-top:5px}.kpi-note{font-size:.8rem;color:var(--muted)}
.two{display:grid;grid-template-columns:1fr 1fr;gap:14px}.panel h3{margin:0 0 14px;color:var(--navy);font-size:1.05rem}
.bar-list{display:grid;gap:10px}.bar-row{display:grid;grid-template-columns:minmax(120px,1.3fr) 2fr auto;gap:10px;align-items:center;font-size:.88rem}.bar-label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.track{height:12px;background:#edf2f6;border-radius:20px;overflow:hidden}.fill{height:100%;background:var(--teal);border-radius:20px}.bar-value{font-variant-numeric:tabular-nums;color:var(--muted);text-align:right}
.line-wrap{width:100%;overflow:hidden}.line-wrap svg{width:100%;height:auto}.axis{stroke:#dbe3e9;stroke-width:1}.trend{fill:none;stroke:var(--teal);stroke-width:3}.dot{fill:var(--coral)}
table{width:100%;border-collapse:collapse;font-size:.9rem}th,td{text-align:left;padding:9px 8px;border-bottom:1px solid var(--line)}th{color:var(--muted);font-weight:650}td.num{text-align:right;font-variant-numeric:tabular-nums}
select{max-width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:8px;color:var(--ink);background:#fff;margin-bottom:12px}
.tag{display:inline-block;background:#e8f4f3;color:#0f6663;padding:3px 8px;border-radius:999px;font-size:.78rem;font-weight:650}
.notice{border-left:4px solid var(--gold);background:#fff9e9;padding:14px 16px;border-radius:8px;color:#5d4b21}.footer{margin-top:30px;color:var(--muted);font-size:.85rem}
.empty{color:var(--muted);font-style:italic}
@media(max-width:850px){.grid{grid-template-columns:repeat(2,1fr)}.two{grid-template-columns:1fr}}
@media(max-width:520px){.wrap{padding:14px}.grid{grid-template-columns:1fr 1fr}.bar-row{grid-template-columns:95px 1fr auto;gap:6px;font-size:.8rem}.hero{padding:24px}}
@media print{body{background:#fff}.wrap{max-width:none;padding:0}.hero,.card,.panel{box-shadow:none;break-inside:avoid}.hero{background:#17324d!important;color:#fff!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}h2{break-after:avoid}}
</style>
</head>
<body>
<main class="wrap">
<header class="hero">
  <div class="eyebrow">Grupo 7 · Proyecto final · Ejercicio 7</div>
  <h1>Opiniones de clientes en restaurantes</h1>
  <p>Exploración agregada de reseñas Yelp: calificaciones, sentimiento, etiquetas de negocio, precio, tiempo y perfiles. Esta página es un complemento estático; no es una publicación en vivo del archivo PBIX.</p>
</header>
<section class="grid" id="kpis" aria-label="Indicadores principales"></section>
<section>
  <h2>Tendencia mensual</h2>
  <p class="sub">Promedio diario de estrellas agregado por mes. Las fechas sin reseñas no aportan observaciones.</p>
  <div class="panel">
    <label for="year-filter">Filtrar año:</label>
    <select id="year-filter"></select>
    <div class="line-wrap" id="trend"></div>
  </div>
</section>
<section class="two">
  <div class="panel"><h3>Calificación por rango de precio</h3><div id="prices" class="bar-list"></div></div>
  <div class="panel"><h3>Ciudades con más reseñas</h3><div id="cities" class="bar-list"></div></div>
</section>
<section>
  <h2>Sentimiento por etiqueta Yelp</h2>
  <p class="sub">Se muestran grupos con al menos 100 reseñas para reducir rankings dominados por muestras muy pequeñas.</p>
  <div class="two">
    <div class="panel"><h3>Mayor proporción positiva (mínimo 100)</h3><div id="positive-categories" class="bar-list"></div></div>
    <div class="panel"><h3>Más reseñas por etiqueta</h3><div id="category-volume" class="bar-list"></div></div>
  </div>
</section>
<section class="two">
  <div class="panel"><h3>Top restaurantes por calificación (mínimo 100 reseñas)</h3><div id="top-rated"></div></div>
  <div class="panel"><h3>Top restaurantes por volumen de reseñas</h3><div id="top-volume"></div></div>
</section>
<section class="two">
  <div class="panel"><h3>Términos de queja</h3><div id="complaints" class="bar-list"></div><p class="kpi-note">Una reseña puede contener varios términos; estos conteos no son sumables como reseñas únicas.</p></div>
  <div class="panel"><h3>Palabras frecuentes</h3><div id="keywords"></div></div>
</section>
<section class="two">
  <div class="panel"><h3>Perfil por actividad del usuario</h3><div id="profiles"></div><p class="kpi-note">Los segmentos se basan en cuartiles de cantidad de reseñas; no son perfiles demográficos.</p></div>
  <div class="panel"><h3>Dispersión por etiqueta Yelp</h3><div id="deviation" class="bar-list"></div><p class="kpi-note">Desviación estándar poblacional de estrellas; leer junto con el tamaño del grupo.</p></div>
</section>
<section>
  <h2>Alcance y limitaciones</h2>
  <div class="notice">
    <strong>Interpretación de categoría:</strong> el ETL toma la primera etiqueta Yelp no genérica de un negocio que incluye “Restaurants”. No siempre es una cocina: el conjunto contiene etiquetas secundarias como servicios, comercio o mascotas. Interpreta estas comparaciones como grupos de etiquetas Yelp y no como categorías gastronómicas.
  </div>
  <div class="notice" style="margin-top:12px;border-left-color:#d76a4a;background:#fff2ee;color:#693f32">
    <strong>Agrupación de ciudades:</strong> los CSV contienen variantes de nombres por mayúsculas y espacios. Al normalizarlos con espacios externos eliminados y comparación sin distinguir mayúsculas, la agregación local reproduce los 848 grupos de Power BI, incluidas las 687.553 reseñas de Philadelphia y las 4.724.471 reseñas totales.
  </div>
  <p>Las reseñas son autoseleccionadas y no representan a todos los clientes. VADER clasifica texto en inglés mediante reglas léxicas. La asociación entre precio y estrellas no implica causalidad. El rango de precio vacío se conserva como dato faltante.</p>
  <p class="kpi-note">Fuente: Yelp Open Dataset. La página contiene únicamente agregados recalculados desde los CSV locales; no incluye textos de reseñas, identificadores de usuarios ni los JSON originales. Es una vista estática independiente del PBIX.</p>
</section>
<footer class="footer">Grupo 7 · Análisis de opiniones de restaurantes · Dashboard HTML estático generado desde las tablas compactas del proyecto.</footer>
</main>
<script id="dashboard-data" type="application/json">__DATA__</script>
<script>
const D=JSON.parse(document.getElementById('dashboard-data').textContent);
const num=new Intl.NumberFormat('es-CO');
const pct=v=>(v*100).toLocaleString('es-CO',{maximumFractionDigits:1})+'%';
const stars=v=>Number(v).toLocaleString('es-CO',{minimumFractionDigits:2,maximumFractionDigits:2});
function kpi(label,value,note){const e=document.createElement('article');e.className='card';e.innerHTML='<div class="kpi-label"></div><div class="kpi-value"></div><div class="kpi-note"></div>';e.children[0].textContent=label;e.children[1].textContent=value;e.children[2].textContent=note;document.getElementById('kpis').append(e)}
function bars(id,items,label,value,format){const root=document.getElementById(id);root.replaceChildren();if(!items.length){root.textContent='Sin datos';return}const max=Math.max(...items.map(x=>Number(x[value])||0),1);for(const item of items){const row=document.createElement('div');row.className='bar-row';const l=document.createElement('span');l.className='bar-label';l.textContent=label(item);const track=document.createElement('div');track.className='track';const fill=document.createElement('div');fill.className='fill';fill.style.width=Math.max(1,(Number(item[value])/max)*100)+'%';track.append(fill);const v=document.createElement('span');v.className='bar-value';v.textContent=format(item[value]);row.append(l,track,v);root.append(row)}}
function table(id,headers,rows){const root=document.getElementById(id);root.replaceChildren();const t=document.createElement('table'),head=document.createElement('thead'),hr=document.createElement('tr');headers.forEach(h=>{const th=document.createElement('th');th.textContent=h;hr.append(th)});head.append(hr);t.append(head);const body=document.createElement('tbody');rows.forEach(values=>{const tr=document.createElement('tr');values.forEach((v,i)=>{const td=document.createElement('td');td.textContent=v;if(i>0)td.className='num';tr.append(td)});body.append(tr)});t.append(body);root.append(t)}
function trend(year){let rows=D.monthly;if(year!=='all')rows=rows.filter(x=>x.month.startsWith(year));const root=document.getElementById('trend');if(!rows.length){root.textContent='Sin datos';return}const w=980,h=260,p=34,vals=rows.map(x=>Number(x.average_stars)),min=Math.min(...vals)-.08,max=Math.max(...vals)+.08;const points=vals.map((v,i)=>{const x=p+i*(w-2*p)/Math.max(1,vals.length-1),y=h-p-(v-min)*(h-2*p)/(max-min||1);return [x,y]});const path=points.map((a,i)=>(i?'L':'M')+a[0].toFixed(1)+' '+a[1].toFixed(1)).join(' ');root.innerHTML='<svg viewBox="0 0 '+w+' '+h+'" role="img" aria-label="Tendencia mensual de calificación"><line class="axis" x1="'+p+'" y1="'+(h-p)+'" x2="'+(w-p)+'" y2="'+(h-p)+'"/><line class="axis" x1="'+p+'" y1="'+p+'" x2="'+p+'" y2="'+(h-p)+'"/><path class="trend" d="'+path+'"/><circle class="dot" cx="'+points[0][0]+'" cy="'+points[0][1]+'" r="4"/><circle class="dot" cx="'+points.at(-1)[0]+'" cy="'+points.at(-1)[1]+'" r="4"/><text x="'+p+'" y="'+(h-5)+'" fill="#617487">'+rows[0].month+'</text><text x="'+(w-p-70)+'" y="'+(h-5)+'" fill="#617487">'+rows.at(-1).month+'</text><text x="'+p+'" y="20" fill="#17324d">Estrellas (1–5)</text></svg>'}
kpi('Reseñas',num.format(D.meta.reviews),'Estrellas válidas en la tabla compacta');
kpi('Promedio general',stars(D.meta.average_stars)+' / 5','Calificación media de la muestra');
kpi('Reseñas positivas',pct(D.meta.positive_share),'Clasificación de sentimiento VADER');
kpi('Negocios',num.format(D.meta.businesses),num.format(D.meta.cities)+' ciudades con actividad');
const years=[...new Set(D.monthly.map(x=>x.month.slice(0,4)))];const yf=document.getElementById('year-filter');const all=document.createElement('option');all.value='all';all.textContent='Todos';yf.append(all);years.forEach(y=>{const o=document.createElement('option');o.value=y;o.textContent=y;yf.append(o)});yf.addEventListener('change',()=>trend(yf.value));trend('all');
bars('prices',D.price,'label', 'average_stars',v=>stars(v)+' estrellas');
bars('cities',D.cities_top,'label','reviews',v=>num.format(v));
bars('positive-categories',D.category_positive_min100,'label','positive_share',pct);
bars('category-volume',D.category_volume,'label','reviews',v=>num.format(v));
bars('complaints',D.complaints,x=>x.complaint_term,'frequency',v=>num.format(v));
bars('deviation',D.category_deviation_min100,'label','population_stddev',v=>Number(v).toFixed(2));
table('top-rated',['Restaurante','Ciudad','Reseñas','Promedio'],D.top_rating_min100.map(x=>[x.name,x.city,num.format(x.reviews),stars(x.average_stars)]));
table('top-volume',['Restaurante','Ciudad','Reseñas','Promedio'],D.top_volume.map(x=>[x.name,x.city,num.format(x.reviews),stars(x.average_stars)]));
table('keywords',['Palabra','Frecuencia'],D.keywords.map(x=>[x.keyword,num.format(Number(x.frequency))]));
table('profiles',['Segmento','Usuarios','Reseñas prom.','Estrellas prom.','Fans prom.'],D.profiles.map(x=>[x.segment,num.format(x.users),stars(x.average_review_count),stars(x.average_stars),stars(x.average_fans)]));
</script>
</body>
</html>"""
    page = page.replace("__DATA__", payload)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(page, encoding="utf-8")


def build_report_html(markdown_path: Path, output: Path) -> str:
    source = markdown_path.read_text(encoding="utf-8")
    body = markdown.markdown(
        source,
        extensions=["tables", "fenced_code", "sane_lists", "toc"],
        output_format="html5",
    )
    title = html.escape(source.splitlines()[0].lstrip("# ").strip())
    page = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title><style>
@page{{size:A4;margin:18mm 17mm}}body{{font:11pt/1.55 Arial,sans-serif;color:#243447;max-width:900px;margin:36px auto;padding:0 24px}}
h1,h2,h3{{color:#17324d;line-height:1.25;break-after:avoid}}h1{{font-size:25pt;border-bottom:3px solid #148a87;padding-bottom:10px}}h2{{font-size:17pt;margin-top:25px}}h3{{font-size:13pt;margin-top:20px}}
p,li{{orphans:3;widows:3}}table{{border-collapse:collapse;width:100%;font-size:9pt;margin:14px 0;break-inside:avoid}}th,td{{border:1px solid #cbd6df;padding:6px 7px;vertical-align:top}}th{{background:#e8f0f5;color:#17324d}}tr:nth-child(even) td{{background:#f7f9fb}}
pre{{white-space:pre-wrap;background:#f3f6f8;border-left:3px solid #148a87;padding:10px;font-size:8.5pt;break-inside:avoid}}code{{font-family:Consolas,monospace;font-size:.93em}}blockquote{{border-left:3px solid #e3b44b;padding-left:12px;color:#53687b}}a{{color:#146a88;word-break:break-word}}
@media print{{body{{max-width:none;margin:0;padding:0}}}}
</style></head><body>{body}</body></html>"""
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(page, encoding="utf-8")
    return title


class _MarkdownDocxParser(HTMLParser):
    def __init__(self, document: Document) -> None:
        super().__init__(convert_charrefs=True)
        self.document = document
        self.paragraph = None
        self.bold = 0
        self.italic = 0
        self.inline_code = 0
        self.in_pre = False
        self.pre_text: list[str] = []
        self.in_table = False
        self.table_rows: list[tuple[list[str], bool]] = []
        self.current_row: list[str] = []
        self.current_cell: list[str] | None = None
        self.current_header_cell = False
        self.list_stack: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "table":
            self.in_table = True
            self.table_rows = []
        elif self.in_table and tag == "tr":
            self.current_row = []
        elif self.in_table and tag in {"td", "th"}:
            self.current_cell = []
            self.current_header_cell = tag == "th"
        elif tag in {"ul", "ol"}:
            self.list_stack.append("List Bullet" if tag == "ul" else "List Number")
        elif tag == "pre":
            self.in_pre = True
            self.pre_text = []
        elif tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            level = int(tag[1])
            self.paragraph = self.document.add_paragraph(style=f"Heading {min(level, 3)}")
            if level == 1:
                self.paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif tag == "p":
            self.paragraph = self.document.add_paragraph()
        elif tag == "li":
            style = self.list_stack[-1] if self.list_stack else "List Bullet"
            self.paragraph = self.document.add_paragraph(style=style)
        elif tag in {"strong", "b"}:
            self.bold += 1
        elif tag in {"em", "i"}:
            self.italic += 1
        elif tag == "code":
            self.inline_code += 1
        elif tag == "br" and self.paragraph is not None:
            self.paragraph.add_run().add_break()

    def handle_endtag(self, tag: str) -> None:
        if tag in {"p", "h1", "h2", "h3", "h4", "h5", "h6", "li"}:
            self.paragraph = None
        elif tag in {"ul", "ol"} and self.list_stack:
            self.list_stack.pop()
        elif tag in {"strong", "b"}:
            self.bold = max(0, self.bold - 1)
        elif tag in {"em", "i"}:
            self.italic = max(0, self.italic - 1)
        elif tag == "code":
            self.inline_code = max(0, self.inline_code - 1)
        elif tag == "pre":
            paragraph = self.document.add_paragraph()
            paragraph.paragraph_format.left_indent = Inches(0.2)
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(6)
            run = paragraph.add_run("".join(self.pre_text).rstrip())
            run.font.name = "Consolas"
            run.font.size = Pt(8.5)
            run.font.color.rgb = DocxColor.from_string("31475E")
            self.in_pre = False
        elif self.in_table and tag in {"td", "th"} and self.current_cell is not None:
            self.current_row.append("".join(self.current_cell).strip())
            self.current_cell = None
        elif self.in_table and tag == "tr":
            if self.current_row:
                self.table_rows.append((self.current_row, self.current_header_cell))
        elif tag == "table":
            self._write_table()
            self.in_table = False

    def handle_data(self, data: str) -> None:
        if self.in_pre:
            self.pre_text.append(data)
        elif self.in_table and self.current_cell is not None:
            self.current_cell.append(data)
        elif self.paragraph is not None:
            run = self.paragraph.add_run(data)
            run.bold = self.bold > 0
            run.italic = self.italic > 0
            if self.inline_code:
                run.font.name = "Consolas"
                run.font.size = Pt(9)

    def _write_table(self) -> None:
        if not self.table_rows:
            return
        column_count = max(len(row) for row, _ in self.table_rows)
        table = self.document.add_table(rows=0, cols=column_count)
        table.style = "Light Shading Accent 1"
        for row_values, header in self.table_rows:
            cells = table.add_row().cells
            for index in range(column_count):
                cell = cells[index]
                cell.text = row_values[index] if index < len(row_values) else ""
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                        if header:
                            run.bold = True


def build_docx(markdown_path: Path, output: Path) -> None:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    normal = document.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(10)
    normal.font.color.rgb = DocxColor.from_string(COLORS["ink"])
    for style_name, size, color in (
        ("Heading 1", 19, COLORS["navy"]),
        ("Heading 2", 15, COLORS["teal"]),
        ("Heading 3", 12, COLORS["navy"]),
    ):
        style = document.styles[style_name]
        style.font.name = "Aptos Display"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = DocxColor.from_string(color)
    header = section.header.paragraphs[0]
    header.text = "GRUPO 7  |  PROYECTO FINAL  |  EJERCICIO 7"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = DocxColor.from_string(COLORS["muted"])
    footer = section.footer.paragraphs[0]
    footer.text = "Análisis de opiniones de restaurantes · Yelp Open Dataset"
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in footer.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = DocxColor.from_string(COLORS["muted"])

    content = markdown.markdown(
        markdown_path.read_text(encoding="utf-8"),
        extensions=["tables", "fenced_code", "sane_lists"],
        output_format="html5",
    )
    parser = _MarkdownDocxParser(document)
    parser.feed(content)
    document.core_properties.title = "Informe técnico — Grupo 7, Proyecto 7"
    document.core_properties.subject = "Análisis de opiniones de clientes en restaurantes con Yelp"
    document.core_properties.author = "Grupo 7"
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)


def build_pdf(
    markdown_path: Path,
    output: Path,
    title: str = "Informe técnico — Grupo 7, Proyecto 7",
) -> None:
    text = markdown_path.read_text(encoding="utf-8")
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="ReportTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=27,
            textColor=colors.HexColor("#17324D"),
            alignment=TA_CENTER,
            spaceAfter=16,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportH2",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=18,
            textColor=colors.HexColor("#148A87"),
            spaceBefore=14,
            spaceAfter=7,
            keepWithNext=True,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportH3",
            parent=styles["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#17324D"),
            spaceBefore=10,
            spaceAfter=5,
            keepWithNext=True,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=12,
            textColor=colors.HexColor("#243447"),
            spaceAfter=5,
            allowWidows=0,
            allowOrphans=0,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportCode",
            parent=styles["Code"],
            fontName="Courier",
            fontSize=7,
            leading=9,
            backColor=colors.HexColor("#F3F6F8"),
            borderColor=colors.HexColor("#148A87"),
            borderWidth=0.5,
            borderPadding=6,
            leftIndent=6,
            rightIndent=6,
            spaceBefore=4,
            spaceAfter=7,
        )
    )

    def inline_markup(value: str) -> str:
        converted = html.escape(value, quote=False)
        converted = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", converted)
        converted = re.sub(r"`(.+?)`", r'<font name="Courier" size="7.5">\1</font>', converted)
        converted = re.sub(r"\*(.+?)\*", r"<i>\1</i>", converted)
        converted = re.sub(
            r"\[(.+?)\]\((https?://[^ )]+)\)",
            r'<link href="\2" color="#146A88">\1</link>',
            converted,
        )
        return converted

    def parse_table(rows: list[str]) -> Table:
        parsed: list[list[Paragraph]] = []
        for line in rows:
            cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
            if cells and all(re.fullmatch(r":?-{3,}:?", cell) for cell in cells):
                continue
            parsed.append([Paragraph(inline_markup(cell), styles["ReportBody"]) for cell in cells])
        table = Table(parsed, repeatRows=1, hAlign="LEFT", splitByRow=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8F0F5")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17324D")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 7),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CBD6DF")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FB")]),
                ]
            )
        )
        return table

    flowables: list[Any] = []
    lines = text.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            index += 1
            continue
        if stripped.startswith("```"):
            index += 1
            block: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith("```"):
                block.append(lines[index])
                index += 1
            flowables.append(Preformatted("\n".join(block), styles["ReportCode"]))
            index += 1
            continue
        if stripped.startswith("|"):
            table_lines = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                table_lines.append(lines[index])
                index += 1
            flowables.append(parse_table(table_lines))
            flowables.append(Spacer(1, 5))
            continue
        heading = re.match(r"^(#{1,3})\s+(.*)$", stripped)
        if heading:
            level = len(heading.group(1))
            title = inline_markup(heading.group(2))
            if level == 1:
                flowables.append(Paragraph(title, styles["ReportTitle"]))
                flowables.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#148A87"), spaceAfter=10))
            else:
                style = styles["ReportH2"] if level == 2 else styles["ReportH3"]
                flowables.append(Paragraph(title, style))
            index += 1
            continue
        if stripped.startswith("- ") or re.match(r"^\d+\.\s", stripped):
            items = []
            while index < len(lines) and (
                lines[index].strip().startswith("- ")
                or re.match(r"^\d+\.\s", lines[index].strip())
            ):
                current = re.sub(r"^(?:-\s+|\d+\.\s+)", "", lines[index].strip())
                items.append(Paragraph("• " + inline_markup(current), styles["ReportBody"]))
                index += 1
            flowables.extend(items)
            flowables.append(Spacer(1, 2))
            continue
        if stripped.startswith(">"):
            quote = []
            while index < len(lines) and lines[index].strip().startswith(">"):
                quote.append(lines[index].strip().lstrip(">").strip())
                index += 1
            flowables.append(
                Paragraph(
                    inline_markup(" ".join(quote)),
                    ParagraphStyle(
                        name=f"Quote{index}",
                        parent=styles["ReportBody"],
                        leftIndent=10,
                        borderColor=colors.HexColor("#E3B44B"),
                        borderWidth=2,
                        borderPadding=5,
                    ),
                )
            )
            continue
        paragraph_lines = [stripped]
        index += 1
        while index < len(lines):
            next_line = lines[index].strip()
            if not next_line or next_line.startswith(("#", "|", "```", ">")) or next_line.startswith("- ") or re.match(r"^\d+\.\s", next_line):
                break
            paragraph_lines.append(next_line)
            index += 1
        flowables.append(Paragraph(inline_markup(" ".join(paragraph_lines)), styles["ReportBody"]))

    def page_footer(canvas: Any, doc: Any) -> None:
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D9E1E8"))
        canvas.line(17 * mm, 13 * mm, A4[0] - 17 * mm, 13 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#617487"))
        canvas.drawString(17 * mm, 9 * mm, "Grupo 7 · Proyecto final · Yelp Open Dataset")
        canvas.drawRightString(A4[0] - 17 * mm, 9 * mm, f"Página {doc.page}")
        canvas.restoreState()

    output.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(output),
        pagesize=A4,
        rightMargin=17 * mm,
        leftMargin=17 * mm,
        topMargin=17 * mm,
        bottomMargin=19 * mm,
        title=title,
        author="Grupo 7",
    )
    document.build(flowables, onFirstPage=page_footer, onLaterPages=page_footer)


def _ppt_color(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def _add_text(
    slide: Any,
    text: str,
    x: float,
    y: float,
    width: float,
    height: float,
    size: float = 18,
    color: str | None = None,
    bold: bool = False,
    align: Any = PP_ALIGN.LEFT,
) -> Any:
    shape = slide.shapes.add_textbox(
        PptxInches(x), PptxInches(y), PptxInches(width), PptxInches(height)
    )
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = PptxInches(0.02)
    frame.margin_right = PptxInches(0.02)
    frame.margin_top = PptxInches(0.01)
    frame.margin_bottom = PptxInches(0.01)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Aptos"
    run.font.size = PptxPt(size)
    run.font.bold = bold
    run.font.color.rgb = _ppt_color(color or COLORS["ink"])
    return shape


def _add_bullets(
    slide: Any, lines: list[str], x: float, y: float, width: float, height: float, size: float = 18
) -> None:
    shape = slide.shapes.add_textbox(
        PptxInches(x), PptxInches(y), PptxInches(width), PptxInches(height)
    )
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = PptxInches(0.08)
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.level = 0
        paragraph.font.name = "Aptos"
        paragraph.font.size = PptxPt(size)
        paragraph.font.color.rgb = _ppt_color(COLORS["ink"])
        paragraph.space_after = PptxPt(10)


def _add_card(slide: Any, x: float, y: float, width: float, label: str, value: str, note: str = "") -> None:
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        PptxInches(x),
        PptxInches(y),
        PptxInches(width),
        PptxInches(1.05),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _ppt_color(COLORS["white"])
    shape.line.color.rgb = _ppt_color(COLORS["line"])
    _add_text(slide, label, x + 0.16, y + 0.12, width - 0.3, 0.22, 10, COLORS["muted"], True)
    _add_text(slide, value, x + 0.16, y + 0.34, width - 0.3, 0.4, 22, COLORS["navy"], True)
    if note:
        _add_text(slide, note, x + 0.16, y + 0.77, width - 0.3, 0.2, 8, COLORS["muted"])


def _new_slide(prs: Presentation, title: str, subtitle: str = "") -> Any:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _ppt_color(COLORS["paper"])
    _add_text(slide, title, 0.55, 0.28, 12.2, 0.48, 25, COLORS["navy"], True)
    if subtitle:
        _add_text(slide, subtitle, 0.58, 0.83, 12.0, 0.36, 11, COLORS["muted"])
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, PptxInches(0.58), PptxInches(1.23), PptxInches(0.72), PptxInches(0.045)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = _ppt_color(COLORS["teal"])
    line.line.fill.background()
    return slide


def _add_footer(slide: Any, slide_num: int) -> None:
    _add_text(slide, f"GRUPO 7  ·  PROYECTO 7                                      {slide_num:02d}", 0.6, 7.13, 12.0, 0.2, 8, COLORS["muted"])


def _add_bars(
    slide: Any,
    title: str,
    items: list[dict[str, Any]],
    label_key: str,
    value_key: str,
    x: float,
    y: float,
    width: float,
    height: float,
    formatter: Any = None,
) -> None:
    _add_text(slide, title, x, y, width, 0.3, 14, COLORS["navy"], True)
    if not items:
        _add_text(slide, "Sin datos disponibles", x, y + 0.45, width, 0.35, 11, COLORS["muted"])
        return
    rows = items[:10]
    row_height = min(0.42, (height - 0.42) / len(rows))
    label_width = width * 0.38
    bar_width = width * 0.45
    value_width = width - label_width - bar_width - 0.1
    maximum = max(float(row[value_key]) for row in rows) or 1.0
    for index, row in enumerate(rows):
        top = y + 0.38 + index * row_height
        label = str(row[label_key])
        if len(label) > 27:
            label = label[:24] + "..."
        _add_text(slide, label, x, top, label_width, row_height * 0.8, 9.2, COLORS["ink"])
        background = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            PptxInches(x + label_width),
            PptxInches(top + 0.05),
            PptxInches(bar_width - 0.08),
            PptxInches(0.16),
        )
        background.fill.solid()
        background.fill.fore_color.rgb = _ppt_color(COLORS["line"])
        background.line.fill.background()
        fill_width = max(0.03, (bar_width - 0.08) * float(row[value_key]) / maximum)
        fill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            PptxInches(x + label_width),
            PptxInches(top + 0.05),
            PptxInches(fill_width),
            PptxInches(0.16),
        )
        fill.fill.solid()
        fill.fill.fore_color.rgb = _ppt_color(COLORS["teal"])
        fill.line.fill.background()
        value = formatter(row[value_key]) if formatter else f"{row[value_key]:,.0f}"
        _add_text(slide, value, x + label_width + bar_width - 0.02, top, value_width, row_height * 0.8, 9, COLORS["muted"], align=PP_ALIGN.RIGHT)


def build_presentation(summary: dict[str, Any], output: Path) -> None:
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    slide_num = 0

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _ppt_color(COLORS["navy"])
    _add_text(slide, "GRUPO 7  ·  PROYECTO FINAL  ·  EJERCICIO 7", 0.85, 0.9, 11.6, 0.35, 13, "B8D7DC", True)
    _add_text(slide, "Opiniones de clientes\nen restaurantes", 0.85, 1.7, 11.6, 1.7, 38, COLORS["white"], True)
    _add_text(slide, "ETL · Power BI · DAX · ANOVA", 0.9, 3.75, 10.5, 0.5, 21, "75C7BF", True)
    _add_text(slide, "Yelp Open Dataset | Resultados descriptivos de una muestra de reseñas", 0.9, 4.45, 11.4, 0.55, 15, "E1EAF0")
    _add_text(slide, "Guion visual para exposición oral · No implica causalidad", 0.9, 6.55, 11.0, 0.32, 11, "B8D7DC")
    slide_num += 1

    slide = _new_slide(prs, "1. Pregunta y alcance", "¿Qué patrones muestran las reseñas de negocios etiquetados como restaurantes?")
    _add_bullets(slide, [
        "Objetivo: procesar reseñas, medir satisfacción y explorar diferencias por tiempo, precio, ciudad y etiquetas Yelp.",
        "El ETL filtra negocios cuya lista de categorías contiene “Restaurants”.",
        "Las reseñas son autoseleccionadas; describimos este conjunto, no a todos los clientes.",
        "El sentimiento textual se calcula con VADER, un método léxico orientado al inglés.",
    ], 0.8, 1.65, 7.2, 4.2, 18)
    _add_card(slide, 8.55, 1.9, 3.5, "RESEÑAS", f"{summary['meta']['reviews']:,}", "Tabla compacta del modelo")
    _add_card(slide, 8.55, 3.2, 3.5, "CALIFICACIÓN MEDIA", f"{summary['meta']['average_stars']:.2f} / 5")
    _add_card(slide, 8.55, 4.5, 3.5, "NEGOCIOS", f"{summary['meta']['businesses']:,}", "Con reseñas incluidas")
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "2. Flujo ETL y modelo", "De JSON Lines a un modelo tabular para análisis reproducible")
    _add_bullets(slide, [
        "Extract: lectura validada de negocios, reseñas y usuarios.",
        "Transform: filtro Restaurants, limpieza y fechas; sentimiento VADER; etiquetas y términos de queja.",
        "Load: dimensiones de negocio, usuario y fecha relacionadas con el hecho de reseñas.",
        "Controles: errores explícitos para archivos ausentes, JSON inválido y salidas vacías.",
    ], 0.75, 1.55, 6.0, 4.7, 16)
    _add_text(slide, "Esquema de estrella", 7.3, 1.55, 4.8, 0.3, 15, COLORS["navy"], True)
    for label, x, y, color in [
        ("dim_business\nciudad · precio · ubicación", 7.1, 2.05, COLORS["blue"]),
        ("dim_user\nactividad · perfil", 9.8, 2.05, COLORS["coral"]),
        ("dim_date\naño · trimestre · mes · fecha", 7.1, 4.5, COLORS["gold"]),
        ("fact_review\nestrellas · sentimiento · votos", 9.8, 4.5, COLORS["teal"]),
    ]:
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(x), PptxInches(y), PptxInches(2.35), PptxInches(1.1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = _ppt_color(color)
        shape.line.fill.background()
        _add_text(slide, label, x + 0.12, y + 0.15, 2.12, 0.8, 12, COLORS["white"], True, PP_ALIGN.CENTER)
    _add_text(slide, "Las etiquetas Yelp no siempre significan cocina; se explican como una limitación del campo category.", 7.15, 6.0, 5.1, 0.55, 10, COLORS["muted"])
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "3. Diez análisis propios del Proyecto 7", "El tablero y el ETL responden las preguntas del ejercicio sobre restaurantes")
    _add_bullets(slide, [
        "1–2: calificación por negocio y sentimiento por etiqueta Yelp.",
        "3–4: palabras frecuentes y tendencia mensual de satisfacción.",
        "5–6: promedio por ciudad y relación descriptiva precio–calificación.",
        "7–8: Top 10 mejor calificados y frecuencia de términos de queja.",
        "9–10: reseñas positivas/negativas y segmentación por actividad de usuario.",
    ], 0.85, 1.7, 11.4, 4.0, 20)
    _add_text(slide, "La nube de palabras y el sentimiento textual se preparan en el ETL; no se presentan como sustitutos de DAX.", 0.9, 6.2, 11.4, 0.5, 12, COLORS["muted"])
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "4. Diez consultas clave DAX", "Las diez consultas tabulares EVALUATE se ejecutaron en Power BI y devolvieron «Correcto».")
    _add_bullets(slide, [
        "1–3: COUNT de reseñas por ciudad; AVERAGE por precio; SUM de votos útiles.",
        "4–5: CALCULATE/KPI positivo; segmentación año–trimestre–mes–día.",
        "6–7: MTD, YTD, SAMEPERIODLASTYEAR y VAR/RETURN.",
        "8–9: RANKX por volumen; SEARCH para términos de queja.",
        "10: STDEV.P para dispersión de estrellas.",
    ], 0.8, 1.6, 7.0, 3.7, 17)
    rows = [848, 5, 445, 445, 6018, 31, 204, 10, 420, 445]
    _add_text(slide, "Filas devueltas por consulta", 8.2, 1.65, 4.0, 0.35, 14, COLORS["navy"], True)
    for idx, count in enumerate(rows):
        x = 8.25 + (idx % 5) * 0.83
        y = 2.25 + (idx // 5) * 1.15
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            PptxInches(x),
            PptxInches(y),
            PptxInches(0.72),
            PptxInches(0.84),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = _ppt_color(COLORS["white"])
        shape.line.color.rgb = _ppt_color(COLORS["line"])
        _add_text(slide, f"Q{idx + 1}", x + 0.05, y + 0.1, 0.62, 0.2, 9, COLORS["muted"], True, PP_ALIGN.CENTER)
        _add_text(slide, f"{count:,}", x + 0.02, y + 0.4, 0.68, 0.25, 10, COLORS["navy"], True, PP_ALIGN.CENTER)
    _add_text(slide, "Las consultas 3, 4, 9 y 10 agrupan etiquetas Yelp, no necesariamente tipos de cocina.", 0.85, 5.72, 11.6, 0.38, 11, COLORS["coral"], True)
    _add_text(slide, "Q1: 848 ciudades y 687.553 reseñas en Philadelphia. La agregación local normaliza mayúsculas y espacios para replicar el agrupamiento de Power BI; total: 4.724.471 reseñas.", 0.85, 6.12, 11.6, 0.62, 9.5, COLORS["muted"])
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "5. Precio y satisfacción", "Calificación promedio por rango de precio, con el volumen como contexto")
    _add_bars(slide, "Promedio observado (1–5 estrellas)", summary["price"], "label", "average_stars", 0.8, 1.55, 7.5, 4.7, lambda value: f"{value:.2f}")
    table_rows = [
        [row["label"], f"{row['reviews']:,}", f"{row['average_stars']:.2f}"]
        for row in summary["price"]
    ]
    _add_text(slide, "Rango 4 muestra promedio más alto en la consulta (4,15); rango 1, 3,65.", 8.7, 2.0, 3.6, 0.9, 16, COLORS["navy"], True)
    _add_text(slide, "La categoría “Sin dato” conserva rangos de precio faltantes. La relación es descriptiva, no causal.", 8.7, 3.45, 3.7, 1.1, 12, COLORS["muted"])
    _add_text(slide, f"Las cinco filas suman {sum(row['reviews'] for row in summary['price']):,} reseñas.", 8.7, 5.15, 3.5, 0.5, 12, COLORS["muted"])
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "6. Tiempo y satisfacción", "Las consultas temporales devuelven promedios en distintos contextos")
    last_year = summary["monthly"][-1]["month"][:4] if summary["monthly"] else ""
    last_month = [row for row in summary["monthly"] if row["month"].startswith(last_year)][-12:]
    _add_bars(slide, f"Últimos meses disponibles ({last_year})", last_month, "month", "average_stars", 0.75, 1.55, 7.8, 4.8, lambda value: f"{value:.2f}")
    _add_bullets(slide, [
        "Consulta 5: 6.018 filas con fecha, mes, trimestre y año.",
        "Consulta 6: 31 fechas recientes con promedio MTD; el 19/01/2022 se observa 3,83.",
        "Consulta 7: 204 filas año–mes; el cambio interanual queda vacío si no existe período comparable.",
        "La tabla de fechas incluye días sin reseñas; SUMMARIZECOLUMNS devuelve fechas con resultados.",
    ], 8.9, 1.65, 3.6, 4.4, 13)
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "7. Sentimiento y quejas", "El sentimiento automático y las quejas son indicadores complementarios")
    sentiment_items = [
        {"label": "Positivas", "value": summary["meta"]["sentiment"].get("Positive", 0)},
        {"label": "Neutrales", "value": summary["meta"]["sentiment"].get("Neutral", 0)},
        {"label": "Negativas", "value": summary["meta"]["sentiment"].get("Negative", 0)},
    ]
    _add_bars(slide, "Volumen de etiquetas VADER", sentiment_items, "label", "value", 0.75, 1.55, 5.6, 3.4, lambda value: f"{value:,}")
    complaints = [
        {"label": row.get("complaint_term", ""), "frequency": int(row.get("frequency") or 0)}
        for row in summary["complaints"]
    ]
    _add_bars(slide, "Términos de queja", complaints, "label", "frequency", 6.8, 1.55, 5.7, 3.4, lambda value: f"{value:,}")
    top_cat = summary["category_volume"][:5]
    _add_text(slide, "La consulta DAX cuenta cada reseña una sola vez si contiene al menos un término; el conteo por término puede solaparse.", 0.8, 5.4, 11.6, 0.7, 12, COLORS["muted"])
    _add_text(slide, "VADER usa reglas léxicas en inglés; conviene validar manualmente una muestra antes de hacer afirmaciones fuertes.", 0.8, 6.15, 11.6, 0.55, 12, COLORS["coral"], True)
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "8. Restaurantes y etiquetas Yelp", "El volumen de reseñas no equivale a calidad ni a satisfacción")
    _add_bars(slide, "Más reseñas", summary["top_volume"][:8], "name", "reviews", 0.65, 1.5, 6.1, 4.8, lambda value: f"{value:,}")
    _add_bars(slide, "Mejor promedio con mínimo 100 reseñas", summary["top_rating_min100"][:8], "name", "average_stars", 6.9, 1.5, 5.8, 4.8, lambda value: f"{value:.2f}")
    _add_text(slide, "El umbral de 100 reduce rankings basados en muy pocas reseñas; no debe confundirse con el Top 10 DAX por volumen.", 0.75, 6.45, 11.8, 0.45, 10, COLORS["muted"])
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "9. ANOVA: diferencias entre grupos", "Welch ANOVA es la prueba principal porque Levene detecta varianzas desiguales")
    _add_card(slide, 0.8, 1.55, 2.5, "MUESTRA", "1.210.971", "5 etiquetas Yelp")
    _add_card(slide, 3.55, 1.55, 2.5, "LEVENE", "p < 0,001", "No hay homogeneidad")
    _add_card(slide, 6.3, 1.55, 2.5, "WELCH F", "976,19", "df = 4; 594.272,81")
    _add_card(slide, 9.05, 1.55, 2.5, "RESULTADO", "p < 0,001", "Rechazar H0")
    means = [
        {"label": "Breakfast & Brunch", "value": 3.88},
        {"label": "American (New)", "value": 3.84},
        {"label": "Bars", "value": 3.74},
        {"label": "Mexican", "value": 3.73},
        {"label": "American (Traditional)", "value": 3.66},
    ]
    _add_bars(slide, "Promedios observados por etiqueta", means, "label", "value", 0.85, 3.05, 6.5, 2.7, lambda value: f"{value:.2f}")
    _add_bullets(slide, [
        "H0: las medias de estrellas son iguales.",
        "H1: al menos una media difiere.",
        "Shapiro-Wilk fue significativo; las estrellas son discretas y N es muy grande.",
        "La prueba global no identifica pares ni demuestra causalidad.",
    ], 7.7, 3.0, 4.3, 2.7, 13)
    _add_text(slide, "El factor es la primera etiqueta Yelp no genérica asociada a un negocio con Restaurants; no necesariamente es una cocina.", 0.85, 6.3, 11.5, 0.55, 11, COLORS["coral"], True)
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "10. Conclusiones y límites", "Qué podemos afirmar y qué no")
    _add_bullets(slide, [
        "En la muestra, la calificación varía por período, rango de precio y etiqueta Yelp.",
        "Welch ANOVA detecta diferencias entre las cinco etiquetas más observadas; no prueba causalidad.",
        "El sentimiento VADER describe el texto según reglas en inglés, no una evaluación humana.",
        "La categoría almacenada puede ser un tag secundario; no llamarla cocina sin revisar su semántica.",
        "Los usuarios de Yelp se autoseleccionan; los hallazgos no generalizan automáticamente a todos los clientes.",
        "La correlación entre precio y calificación no demuestra que pagar más cause satisfacción.",
    ], 0.85, 1.65, 11.5, 4.8, 17)
    slide_num += 1
    _add_footer(slide, slide_num)

    slide = _new_slide(prs, "11. Entrega y reproducibilidad", "Archivos fuente, resultados y vista HTML")
    _add_bullets(slide, [
        "ETL: etl_pipeline.py; preparación para Power BI: prepare_powerbi.py.",
        "Medidas y consultas: dax_measures.dax y dax_queries.dax.",
        "Informe: report.md y versiones DOCX/PDF generadas.",
        "Presentación oral: este PPTX y el guion presentation.md.",
        "Dashboard HTML: versión estática agregada; no sustituye al PBIX ni a una publicación autenticada de Power BI Service.",
        "Datos: Yelp Open Dataset; no redistribuir los JSON originales y respetar sus términos.",
    ], 0.85, 1.65, 11.6, 4.7, 16)
    _add_text(slide, "Fuente: https://business.yelp.com/data/resources/open-dataset/", 0.85, 6.5, 11.2, 0.3, 10, COLORS["blue"])
    slide_num += 1
    _add_footer(slide, slide_num)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera informe, presentación y dashboard HTML estático del Grupo 7."
    )
    parser.add_argument("--data-dir", type=Path, default=ROOT / "data" / "processed" / "powerbi")
    parser.add_argument("--report", type=Path, default=ROOT / "report.md")
    parser.add_argument("--dax-guide", type=Path, default=ROOT / "dax_guide.md")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "deliverables")
    parser.add_argument("--chunk-size", type=int, default=200_000)
    args = parser.parse_args()
    if args.chunk_size < 1:
        parser.error("--chunk-size debe ser mayor que cero.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_summary(args.data_dir, args.chunk_size)
    build_dashboard(summary, args.output_dir / "dashboard.html")
    report_html = args.output_dir / "report.html"
    build_report_html(args.report, report_html)
    build_docx(args.report, args.output_dir / "report.docx")
    build_pdf(args.report, args.output_dir / "report.pdf")
    build_pdf(
        args.dax_guide,
        args.output_dir / "guia_dax.pdf",
        title="Guía de estudio DAX — Grupo 7, Proyecto 7",
    )
    build_presentation(summary, args.output_dir / "presentation.pptx")
    print(f"Generados: {args.output_dir.resolve()}", flush=True)
    print(
        "Entregables: report.docx, report.pdf, report.html, guia_dax.pdf, presentation.pptx, dashboard.html",
        flush=True,
    )


if __name__ == "__main__":
    main()
