from __future__ import annotations

import hashlib
import io
import json
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
import os
import re

from PIL import Image
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------
# Config & Paths
# -------------------------
SIZE_ORDER = ["UNICA", "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]
SIZE_ALIAS = {"XXL": "2XL", "2XL": "2XL"}

APP_SUPPORT = Path(__file__).resolve().parent / "wupi_data"
PROJECTS_DIR = APP_SUPPORT / "projects"

BIBBIA_MANUAL_PATH = APP_SUPPORT / "bibbia_manual.json"
CUSTOM_SUPP_PATH = APP_SUPPORT / "custom_suppliers.json"

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
if not ASSETS_DIR.exists():
    ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
LOGO_PATH = ASSETS_DIR / "wupi.png"
FAVICON_PATH = ASSETS_DIR / "favicon.png"

# Normalizzazione colori
COLOR_ALIAS_MAP = {
  "arancione": "arancione", "bianco": "bianco", "bk": "nero", "black": "nero",
  "blu": "navy", "blu_navy": "navy", "blu_royal": "royal", "blunavy": "navy",
  "bluroyal": "royal", "bu": "burgundy", "burgundy": "burgundy", "ca": "cardinal",
  "caramel": "dark_caramel", "cardinal": "cardinal", "cardinal_red": "cardinal",
  "cardinalred": "cardinal", "ch": "chocolate", "chocolate": "chocolate",
  "dark_caramel": "dark_caramel", "dark_grey": "dark_grey", "darkcaramel": "dark_caramel",
  "darkgrey": "dark_grey", "dc": "dark_caramel", "dg": "dark_grey", "dp": "dusty_pink",
  "dusty_pink": "dusty_pink", "dustypink": "dusty_pink", "du": "dusty_green",
  "dusty_green": "dusty_green", "dustygreen": "dusty_green", "dusty_green_": "dusty_green",
  "earth_green": "military", "earthgreen": "military", "eg": "military", "fg": "forest",
  "forest": "forest", "forest_green": "forest", "forestgreen": "forest",
  "giallo": "gold", "go": "gold", "gold": "gold", "grey_heater": "grey_heather",
  "grey_heather": "grey_heather", "greyheater": "grey_heather", "greyheather": "grey_heather",
  "grigio_melange": "grey_heather", "grigiomelange": "grey_heather", "gy": "grey_heather",
  "ib": "ink_blue", "ig": "irish_green", "ink_blue": "ink_blue", "inkblue": "ink_blue",
  "irish_green": "irish_green", "irishgreen": "irish_green", "jade": "salvia",
  "jade_green": "salvia", "jadegreen": "salvia", "ma": "mastic", "marrone": "chocolate",
  "mastic": "mastic", "mb": "mineral_blue", "military": "military", "mineral_blue": "mineral_blue",
  "mineralblue": "mineral_blue", "mocha": "chocolate", "moka": "chocolate", "mu": "mustard",
  "mustard": "mustard", "navy": "navy", "navy_blue": "navy", "navyblue": "navy",
  "nero": "nero", "ny": "navy", "off_white": "off_white", "offwhite": "off_white",
  "ol": "olive", "olive": "olive", "or": "arancione", "orange": "arancione",
  "ow": "off_white", "pe": "petroleum", "peacock": "peacock_ink_blue",
  "petroleum": "petroleum", "pink": "rosa", "pu": "purple", "purple": "purple",
  "rb": "royal", "rd": "red", "red": "red", "ro": "rosa", "rosa": "rosa",
  "rosso": "red", "rosso_cardinal": "cardinal", "rossocardinal": "cardinal",
  "royal": "royal", "royal_blue": "royal", "royalblue": "royal", "ru": "rust",
  "rust": "rust", "sa": "salvia", "salvia": "salvia", "sand": "mastic",
  "sky": "mineral_blue", "urban_slate": "urban_slate", "urban_slathe": "urban_slate",
  "urbanslate": "urban_slate", "urbanslathe": "urban_slate", "us": "urban_slate",
  "viola": "purple", "wh": "bianco", "white": "bianco", "pk": "peacock_ink_blue",
  "peacock_ink_blue": "peacock_ink_blue", "peacockinkblue": "peacock_ink_blue",
  "peacock_inkblue": "peacock_ink_blue", "peacockink_blue": "peacock_ink_blue",
  "lg": "light_grey", "lightgrey": "light_grey", "light_grey": "light_grey",
  "light_gray": "light_grey", "ac": "sand_almond_cream", "almond_cream": "sand_almond_cream",
  "sand_almond_cream": "sand_almond_cream", "sandalmondcream": "sand_almond_cream",
  "almond_sand_cream": "sand_almond_cream", "almondsandcream": "sand_almond_cream",
  "almond_sandcream": "sand_almond_cream", "almondsand_cream": "sand_almond_cream",
  "sand_cream_almond": "sand_almond_cream", "cream_almond_sand": "sand_almond_cream",
  "earthygreen": "military", "earthy_green": "military", "earthy green": "military",
  "gh": "grey_heather"
}

def color_to_canon_key(s: str) -> str:
    t = _norm_key(s)
    if not t: return ""
    return COLOR_ALIAS_MAP.get(t, t)

# -------------------------
# Database / JSON Helpers
# -------------------------
def clean_str(x) -> str:
    if isinstance(x, str):
        s = x.strip()
        return "" if s.lower() in ("nan", "none", "null") else s
    if x is None: return ""
    try:
        if pd.isna(x): return ""
    except ValueError: pass
    s = str(x).strip()
    return "" if s.lower() in ("nan", "none", "null") else s

def safe_dir_name(s: str) -> str:
    return re.sub(r'[^A-Za-z0-9_\-\s]', '', clean_str(s)).strip()

def _read_json(path: Path, default):
    if path.exists():
        try: return json.loads(path.read_text(encoding="utf-8"))
        except Exception: pass
    return default

def _write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def load_state(p: Path) -> list: return _read_json(p / "state.json", [])
def save_state(p: Path, d: list): _write_json(p / "state.json", d)

def load_subs(p: Path) -> dict: return _read_json(p / "subs.json", {})
def save_subs(p: Path, d: dict): _write_json(p / "subs.json", d)

def load_stock(p: Path) -> dict: return _read_json(p / "stock.json", {})
def save_stock(p: Path, d: dict): _write_json(p / "stock.json", d)

def load_costs(p: Path) -> dict: return _read_json(p / "costs.json", {})
def save_costs(p: Path, d: dict): _write_json(p / "costs.json", d)

def load_custom_suppliers() -> list: return _read_json(CUSTOM_SUPP_PATH, [])
def save_custom_suppliers(d: list): _write_json(CUSTOM_SUPP_PATH, d)

def load_manual_mockups() -> Dict:
    raw = _read_json(BIBBIA_MANUAL_PATH, {})
    out = {}
    for k, v in raw.items():
        try:
            if isinstance(v, str) and Path(v).exists():
                out[tuple(k.split("|||"))] = Path(v).read_bytes()
        except Exception: pass
    return out

def save_manual_mockups(rawmap: Dict[str, str]) -> None:
    _write_json(BIBBIA_MANUAL_PATH, rawmap)

def _cost_key(sku: str, product: str) -> str:
    return f"{clean_str(sku)}||{clean_str(product)}"

# -------------------------
# Core Logic
# -------------------------
def normalize_size(s: str) -> str:
    s2 = clean_str(s).upper()
    if not s2: return ""
    return SIZE_ALIAS.get(s2, s2)

def _norm_colname(c: str) -> str:
    c2 = str(c).strip().lower()
    c2 = re.sub(r"\s+", " ", c2)
    c2 = c2.replace("à","a").replace("è","e").replace("é","e").replace("ì","i").replace("ò","o").replace("ù","u")
    c2 = c2.replace(".", "").replace("/", " ").replace("-", " ").replace("_", " ")
    c2 = re.sub(r"\s+", " ", c2).strip()
    return c2

def standardize_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    alias = {
        "SKU": ["sku", "codice", "codice articolo", "codice art", "articolo sku", "sku articolo"],
        "Nome Prodotto": ["nome prodotto", "prodotto", "articolo", "descrizione", "nome"],
        "Colore": ["colore", "color", "colour", "variante colore"],
        "Taglia": ["taglia", "size", "misura", "variante taglia"],
        "Pezzi": ["pezzi", "qta", "quantita", "quantita pezzi", "quantita'", "qty", "quantità", "quantità pezzi"],
        "N. Ordine": ["n ordine", "numero ordine", "ordine", "n ord", "n. ordine"],
        "Nome Studente": ["nome studente", "nome"],
        "Cognome Studente": ["cognome studente", "cognome"],
        "Classe": ["classe", "class"],
        "Docente/ATA": ["docente/ata", "docente ata", "docenti/ata", "docenti ata", "docente"],
        "Nome incisione": ["nome incisione", "incisione", "testo incisione", "engraving", "nome personalizzazione"],
        "Prezzo unitario": ["prezzo unitario", "prezzo", "prezzo vendita", "unit price", "price", "prezzo cad", "prezzo iva inclusa"],
        "Prezzo acquisto": ["prezzo acquisto", "costo", "costo unitario", "purchase price", "cost"],
        "Importo ordine": ["importo ordine", "importo totale", "totale ordine", "importo"],
    }
    cols = list(df.columns)
    norm_map = { _norm_colname(c): c for c in cols }
    rename = {}
    for std, keys in alias.items():
        if std in cols: continue
        for k in keys:
            k2 = _norm_colname(k)
            if k2 in norm_map:
                rename[norm_map[k2]] = std
                break
    if rename: df = df.rename(columns=rename)
    return df

def ensure_cols(df: pd.DataFrame, cols: List[str]) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing: raise ValueError(f"Colonne mancanti: {', '.join(missing)}")

def _parse_taglie_items(taglie_str: str) -> list[tuple[str, int]]:
    items = []
    for part in clean_str(taglie_str).split():
        if ":" in part:
            t, q = part.split(":", 1)
            try: items.append((t, int(q)))
            except Exception: pass
    items.sort(key=lambda x: sort_size_key(x[0]))
    return items

def key_row(sku: str, prod: str, color: str) -> str:
    return f"{sku}||{prod}||{color}"

def normalize_key(k: str) -> str:
    try:
        parts = (k or "").split("||")
        if len(parts) != 3: return clean_str(k)
        return key_row(clean_str(parts[0]), clean_str(parts[1]), clean_str(parts[2]))
    except Exception:
        return clean_str(k)

def canon_key(sku: str, prod: str, color: str) -> str:
    return f"{clean_str(sku).lower()}||{clean_str(prod).lower()}||{clean_str(color).lower()}"

def sort_size_key(taglia: str) -> int:
    t = (taglia or "").upper().strip()
    if t == "XXL": t = "2XL"
    if not t: return 999
    return SIZE_ORDER.index(t) if t in SIZE_ORDER else 998

def df_normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = standardize_required_columns(df)
    ensure_cols(df, ["SKU", "Nome Prodotto", "Colore", "Taglia", "Pezzi"])
    out = df.copy()

    out["SKU"] = out["SKU"].map(clean_str)
    out["Nome Prodotto"] = out["Nome Prodotto"].map(clean_str)
    out["Colore"] = out["Colore"].map(clean_str)  
    out["Taglia"] = out["Taglia"].map(normalize_size)
    
    # IL FIX MAGICO: Trasforma subito le taglie vuote in UNICA alla radice
    out.loc[out["Taglia"].eq(""), "Taglia"] = "UNICA"
    
    out["Pezzi"] = pd.to_numeric(out["Pezzi"], errors="coerce").fillna(0).astype(int)
    if "Prezzo unitario" in out.columns:
        out["Prezzo unitario"] = pd.to_numeric(out["Prezzo unitario"], errors="coerce").fillna(0.0)
    else: out["Prezzo unitario"] = 0.0
    if "Prezzo acquisto" in out.columns:
        out["Prezzo acquisto"] = pd.to_numeric(out["Prezzo acquisto"], errors="coerce").fillna(0.0)
    else: out["Prezzo acquisto"] = 0.0

    for c in ["Nome Studente", "Cognome Studente", "Classe", "Docente/ATA", "N. Ordine", "Nome incisione"]:
        if c in out.columns: out[c] = out[c].map(clean_str)
        else: out[c] = "" if c != "N. Ordine" else ""

    out["Studente"] = (out["Nome Studente"].fillna("") + " " + out["Cognome Studente"].fillna("")).str.replace(r"\s+", " ", regex=True).str.strip()
    is_doc = out["Nome Studente"].eq("") & out["Cognome Studente"].eq("")
    out.loc[is_doc, "Studente"] = out.loc[is_doc, "Docente/ATA"].fillna("").astype(str).str.strip()

    out["GruppoEtichetta"] = out["Classe"].where(~is_doc, "Docenti / ATA")
    out.loc[out["GruppoEtichetta"].eq(""), "GruppoEtichetta"] = "Docenti / ATA"

    return out

def pivot_report(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d.loc[d["Taglia"].eq(""), "Taglia"] = "UNICA"

    piv = (
        d.groupby(["SKU", "Nome Prodotto", "Colore", "Taglia"], as_index=False)["Pezzi"].sum()
         .pivot_table(index=["SKU", "Nome Prodotto", "Colore"], columns="Taglia", values="Pezzi", aggfunc="sum", fill_value=0)
         .reset_index()
    )
    size_cols = [s for s in SIZE_ORDER if s in piv.columns]
    rest = [c for c in piv.columns if c not in ["SKU", "Nome Prodotto", "Colore"] + size_cols]
    piv = piv[["SKU", "Nome Prodotto", "Colore"] + size_cols + rest]
    qty_cols = [c for c in piv.columns if c not in ["SKU", "Nome Prodotto", "Colore"]]
    piv["Totale"] = piv[qty_cols].sum(axis=1).astype(int)
    return piv

def render_pivot_html(piv: pd.DataFrame, confirmed: set[str], subs: dict, file_stock: dict) -> None:
    view = piv.copy()
    
    cols = list(view.columns)
    size_cols = [c for c in cols if c not in ["SKU", "Nome Prodotto", "Colore", "Totale"]]

    for c in size_cols:
        view[c] = view[c].replace({0: ""})
    view["Totale"] = piv["Totale"].astype(int)

    css = f"""<style>
.table-wrap {{ overflow:auto; background-color: #ffffff; border: 1px solid #e5e5ea; border-radius: 12px; box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02); }}
table.wupi {{ border-collapse:separate; border-spacing:0; width:100%; font-size:14px; }}
table.wupi th, table.wupi td {{ padding:12px; border-bottom:1px solid #f0f0f2; vertical-align:middle; color: #1d1d1f; }}
table.wupi th {{ position:sticky; top:0; background-color: #fafafc; z-index:2; font-weight:600; letter-spacing: -0.2px; }}
table.wupi td {{ background-color: #ffffff; }}
table.wupi td.tot, table.wupi th.tot {{ position:sticky; right:0; z-index:3; font-weight:700; }}
table.wupi th.tot {{ background-color: #fafafc; border-left: 1px solid #f0f0f2; }}
table.wupi td.tot {{ background-color: #fafafc; border-left: 1px solid #f0f0f2; }}
tr.confirmed td {{ background-color: #e6f7e6; }}
tr.confirmed td.tot {{ background-color: #ccebcc; }}
tr.warehouse td {{ background-color: #e6f0fa; }}
tr.warehouse td.tot {{ background-color: #cce0f5; }}
.center {{ text-align:center; }}
</style>"""

    html: list[str] = []
    html.append(css)
    html.append('<div class="table-wrap"><table class="wupi">')
    html.append('<thead><tr>')

    for c in cols:
        cls = "tot" if c == "Totale" else ""
        align = "center" if c in size_cols + ["Totale"] else ""
        html.append(f'<th class="{cls} {align}">{c}</th>')

    html.append('</tr></thead><tbody>')

    for _, r in view.iterrows():
        sku_raw = clean_str(r.get("SKU", ""))
        col_raw = clean_str(r.get("Colore", ""))
        k = normalize_key(key_row(sku_raw, clean_str(r.get("Nome Prodotto", "")), col_raw))
        
        row_order_tot = 0
        row_stock_tot = 0
        row_original_tot = 0
        
        for sc in size_cols:
            val = r[sc]
            if pd.notna(val) and val != "":
                val_int = int(val)
                stock_k = f"{sku_raw}||{col_raw}||{clean_str(sc)}"
                stock_qty = file_stock.get(stock_k, 0)
                
                row_original_tot += val_int
                row_stock_tot += stock_qty
                row_order_tot += max(0, val_int - stock_qty)

        # LA NUOVA LOGICA DEI COLORI
        tr_cls = ""
        if row_original_tot > 0 and row_stock_tot >= row_original_tot:
            tr_cls = "warehouse"  # Azzurro solo se TUTTO è preso da magazzino
        elif k in confirmed:
            tr_cls = "confirmed"  # Verde se hai premuto Conferma (acquisto totale o parziale)

        html.append(f'<tr class="{tr_cls}">')

        for c in cols:
            cls = "tot" if c == "Totale" else ""
            align = "center" if c in size_cols + ["Totale"] else ""
            val = r[c]
            
            if c == "Totale": val = row_order_tot if row_order_tot > 0 else ""
            elif c in size_cols:
                if pd.notna(val) and val != "":
                    val_int = int(val)
                    if val_int == 0: val = ""
                    else:
                        stock_k = f"{sku_raw}||{col_raw}||{clean_str(c)}"
                        stock_qty = file_stock.get(stock_k, 0)
                        if stock_qty > 0:
                            order_qty = max(0, val_int - stock_qty)
                            if order_qty > 0: val = f"{order_qty} <span style='font-size:10.5px; color:#86868b; font-weight:normal;'>({stock_qty} mag)</span>"
                            else: val = f"<span style='font-size:11px; color:#1976d2; font-weight:700;'>{stock_qty} mag</span>"
                        else: val = str(val_int)
                else: val = ""
                        
            elif c == "SKU":
                sub_key = f"{sku_raw}||{col_raw}"
                sub_data = subs.get(sub_key, {})
                if isinstance(sub_data, str): sub_data = {"fornitore": "Altro", "sku": sub_data}
                if isinstance(sub_data, dict) and sub_data.get("sku"):
                    f_name = sub_data.get("fornitore", "").upper()
                    s_name = sub_data.get("sku", "").upper()
                    disp = f"{f_name}_{s_name}" if f_name and f_name != "ALTRO" else s_name
                    val = f'{val}<br><span style="font-size:11px; font-weight:800; color:#000000; letter-spacing:0.5px;">{disp}</span>'
            
            if pd.isna(val): val = ""
            html.append(f'<td class="{cls} {align}">{val}</td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    st.markdown("".join(html), unsafe_allow_html=True)
def cards_css() -> None:
    st.markdown("""
<style>
.wupi-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:16px; }
@media (max-width: 1100px) { .wupi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
@media (max-width: 720px) { .wupi-grid { grid-template-columns: repeat(1, minmax(0, 1fr)); } }

.wupi-card { background-color: #ffffff; border: 1px solid #e5e5ea; border-radius: 16px; padding: 16px; box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02); transition: transform 0.1s ease, box-shadow 0.1s ease; }
.wupi-card:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0, 0, 0, 0.06); }
.wupi-card.confirmed { background-color: #e6f7e6; border: 1px solid #ccebcc; }
.wupi-card.warehouse { background-color: #e6f0fa; border: 1px solid #cce0f5; } /* AZZURRO MAGAZZINO */

.card-head { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:12px; }
.color-name { font-weight:700; font-size:17px; letter-spacing:-0.3px; color: #1d1d1f; }
.color-tot { font-weight:600; font-size:15px; color: #86868b; }
.chips { display:flex; flex-wrap:wrap; gap:8px; }

.chip { display:inline-flex; gap:6px; align-items:center; padding: 4px 10px; border-radius: 8px; background-color: #f0f0f2; font-size: 13px; font-weight: 500; color: #555; }
.chip .q { font-weight:700; font-size:14px; color: #1d1d1f; }
.wupi-card.confirmed .chip, .wupi-card.warehouse .chip { background-color: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
</style>
""", unsafe_allow_html=True)

def render_color_cards(df: pd.DataFrame, sku: str, prod: str, confirmed: set[str], proj_dir: Path) -> None:
    cards_css()
    sub = df[(df["SKU"] == sku) & (df["Nome Prodotto"] == prod)].copy()
    if sub.empty:
        st.info("Nessun dato per questo SKU/prodotto.")
        return

    # IL FIX: Qualsiasi taglia vuota/anomala diventa "UNICA"
    sub["Taglia"] = sub["Taglia"].replace({"": "UNICA", None: "UNICA", "nan": "UNICA", "NaN": "UNICA"})

    file_stock = load_stock(proj_dir)
    total_sku = int(sub["Pezzi"].sum())
    st.subheader(f"{sku} - {total_sku} pz")
    st.caption("Colori e taglie per questo SKU. Usa 🛒 Conferma per segnare un colore come già ordinato.")

    blocks: list[tuple[str, int, list[tuple[str,int]]]] = []
    for color, g in sub.groupby("Colore", dropna=False):
        agg = g.groupby("Taglia", as_index=False)["Pezzi"].sum()
        items = [(str(r["Taglia"]), int(r["Pezzi"])) for _, r in agg.iterrows() if int(r["Pezzi"]) > 0]
        items.sort(key=lambda x: sort_size_key(x[0]))
        blocks.append((clean_str(color), sum(q for _, q in items), items))
    blocks.sort(key=lambda x: x[0])

    cols = st.columns(3)
    
    # ---- INIZIO CICLO DEI COLORI ----
    for i, (color, tot, items) in enumerate(blocks):
        col = cols[i % 3]
        k = normalize_key(key_row(clean_str(sku), clean_str(prod), clean_str(color)))
        is_done = k in confirmed
        
        chips = ""
        color_stock_tot = 0
        for t, q in items:
            stock_k = f"{clean_str(sku)}||{clean_str(color)}||{clean_str(t)}"
            sq = file_stock.get(stock_k, 0)
            color_stock_tot += sq
            if sq > 0:
                chips += f'<span class="chip">{t} <span class="q">{max(0, q-sq)} <span style="font-size:11px; color:#86868b; font-weight:500;">(+{sq}📦)</span></span></span>'
            else:
                chips += f'<span class="chip">{t} <span class="q">{q}</span></span>'

        is_fully_stocked = (color_stock_tot >= tot and tot > 0)
        card_cls = "warehouse" if is_fully_stocked else ("confirmed" if is_done else "")
        btn_disabled = is_done or is_fully_stocked

        with col:
            st.markdown(f"""
<div class="wupi-card {card_cls}">
  <div class="card-head">
    <div class="color-name">{color if color else '(colore vuoto)'}</div>
    <div class="color-tot">{tot} pz</div>
  </div>
  <div class="chips">{chips}</div>
</div>
""", unsafe_allow_html=True)
            st.markdown('<div style="height:16px"></div>', unsafe_allow_html=True)
            b1, b2, b3 = st.columns([4, 4, 2])
            with b1:
                if st.button("✓ Conferma", key=f"conf__{hashlib.md5(k.encode()).hexdigest()}", disabled=btn_disabled, use_container_width=True):
                    confirmed.add(normalize_key(k))
                    save_state(proj_dir, sorted(list(confirmed)))
                    st.session_state["confirmed"] = set(confirmed)
                    st.rerun()
            with b2:
                if st.button("↩︎ Annulla", key=f"undo__{hashlib.md5(k.encode()).hexdigest()}", disabled=not is_done, use_container_width=True):
                    confirmed.discard(normalize_key(k))
                    save_state(proj_dir, sorted(list(confirmed)))
                    st.session_state["confirmed"] = set(confirmed)
                    st.rerun()
            with b3:
                with st.popover("📦", use_container_width=True):
                    st.markdown(f"**Magazzino {color if color else ''}**")
                    with st.form(f"wh_form_{hashlib.md5(k.encode()).hexdigest()}"):
                        new_stock = {}
                        for t, q in items:
                            stock_k = f"{clean_str(sku)}||{clean_str(color)}||{clean_str(t)}"
                            curr_stock = file_stock.get(stock_k, 0)
                            val = st.number_input(f"{t} (Max {q})", min_value=0, max_value=q, value=min(curr_stock, q), step=1)
                            new_stock[stock_k] = val
                        
                        if st.form_submit_button("💾 Salva", type="primary", use_container_width=True):
                            stk = load_stock(proj_dir)
                            for sk, v in new_stock.items():
                                if v > 0: stk[sk] = v
                                else: stk.pop(sk, None)
                            save_stock(proj_dir, stk)
                            st.rerun()
    # ---- FINE CICLO DEI COLORI ----


    # QUESTI BOTTONI ORA SONO TORNATI RIGOROSAMENTE FUORI DAL CICLO!
    st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)
    _, r1, r2 = st.columns([6, 2, 2])
    with r1:
        if st.button("✓ Conferma tutto lo SKU", key=f"all_{sku}_{prod}", use_container_width=True):
            for color, _, _ in blocks:
                confirmed.add(normalize_key(key_row(sku, prod, color)))
            save_state(proj_dir, sorted(list(confirmed)))
            st.session_state['confirmed'] = set(confirmed)
            st.session_state['advance_next_sku'] = True
            st.rerun()
    with r2:
        if st.button("↩︎ Annulla tutto lo SKU", key=f"unall_{sku}_{prod}", use_container_width=True):
            for color, _, _ in blocks:
                confirmed.discard(normalize_key(key_row(sku, prod, color)))
            save_state(proj_dir, sorted(list(confirmed)))
            st.session_state["confirmed"] = set(confirmed)
            st.rerun()
def global_css() -> None:
    st.markdown("""
<style>
.stApp { background-color: #ffffff; color: #1d1d1f; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
.stButton > button { background-color: #ffffff !important; color: #1d1d1f !important; border: 1px solid #d2d2d7 !important; border-radius: 8px !important; font-weight: 500 !important; box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important; transition: all 0.2s ease !important; }
.stButton > button:hover { border-color: #86868b !important; background-color: #f5f5f7 !important; }
.stButton > button[kind="primary"] { background-color: #1d1d1f !important; color: #ffffff !important; border: 1px solid #1d1d1f !important; }
.stButton > button[kind="primary"]:hover { background-color: #333336 !important; }
*:focus { outline:none !important; }
button:focus { box-shadow: 0 0 0 2px rgba(0,0,0,.1) !important; }
a, a:visited { color:#1d1d1f; }
[data-testid="stFileUploader"] { background-color: #fafafc !important; border: 1px dashed #d2d2d7 !important; border-radius: 12px !important; padding: 1.5rem !important; }
[data-testid="stMetric"], [data-testid="stDataFrame"] > div { background-color: #ffffff !important; border-radius: 12px !important; border: 1px solid #e5e5ea !important; box-shadow: 0 2px 8px rgba(0,0,0,0.02) !important; }
[data-testid="stMetric"] { padding: 16px; }
.wupi-gap-after-pivot { height: 14px; }
</style>
""", unsafe_allow_html=True)

# -------------------------
# PDF Ordine Fornitore
# -------------------------
def make_order_summary_pdf(piv_df: pd.DataFrame, subs: dict, file_stock: dict, school: str, proj: str) -> bytes:
    from reportlab.platypus import Image as RLImage
    buf = io.BytesIO()
    
    usable_width = A4[0] - 24 * mm
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=12*mm, leftMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    elements = []
    styles = getSampleStyleSheet()

    school_style = ParagraphStyle('SchoolStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1d1d1f"), spaceAfter=0)
    proj_style = ParagraphStyle('ProjStyle', parent=styles['Heading3'], fontSize=14, textColor=colors.HexColor("#86868b"), spaceBefore=0)
    sku_style = ParagraphStyle('SkuStyle', parent=styles['Heading2'], fontSize=13, spaceBefore=0, spaceAfter=0, textColor=colors.HexColor("#1d1d1f"))
    right_tot_style = ParagraphStyle('RightTotStyle', parent=styles['Normal'], alignment=2, textColor=colors.HexColor("#86868b"), fontSize=11)
    centered_style = ParagraphStyle('CenteredStyle', parent=styles['Normal'], alignment=1, leading=10)

    left_p = [Paragraph(f"<b>{school}</b>", school_style), Paragraph(proj, proj_style)]
    
    logo_img = ""
    if LOGO_PATH.exists():
        try:
            im = Image.open(LOGO_PATH)
            w, h = im.size
            aspect = w / h
            target_h = 14 * mm
            logo_img = RLImage(str(LOGO_PATH), width=target_h * aspect, height=target_h)
            logo_img.hAlign = 'RIGHT'
        except:
            logo_img = ""

    header_t = Table([[left_p, logo_img]], colWidths=[usable_width - 60*mm, 60*mm])
    header_t.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    
    elements.append(header_t)
    elements.append(Spacer(1, 10*mm))

    size_cols = [c for c in piv_df.columns if c not in ["SKU", "Nome Prodotto", "Colore", "Totale"]]
    colore_w = 42 * mm
    qty_w = (usable_width - colore_w) / max(1, (len(size_cols) + 1))
    col_widths = [colore_w] + [qty_w] * (len(size_cols) + 1)

    grouped = piv_df.groupby(["SKU", "Nome Prodotto"], sort=False)

    for (sku, prod), group in grouped:
        block = []
        tot_sku = int(group["Totale"].sum())
        
        p_left = Paragraph(f"<b>{sku}</b> — {prod}", sku_style)
        p_right = Paragraph(f"Totale {tot_sku} pz", right_tot_style)
        title_t = Table([[p_left, p_right]], colWidths=[usable_width - 40*mm, 40*mm])
        title_t.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        block.append(title_t)

        header = ["Colore"] + size_cols + ["Totale"]
        data = [header]

        for _, r in group.iterrows():
            col = str(r["Colore"])
            sub_key = f"{clean_str(sku)}||{clean_str(col)}"
            
            sub_data = subs.get(sub_key, {})
            if isinstance(sub_data, str): sub_data = {"fornitore": "Altro", "sku": sub_data}

            if isinstance(sub_data, dict) and sub_data.get("sku"):
                f_name = sub_data.get("fornitore", "").upper()
                s_name = sub_data.get("sku", "").upper()
                disp = f"{f_name}_{s_name}" if f_name and f_name != "ALTRO" else s_name
                col_p = Paragraph(f"{col}<br/><font color='#1d1d1f' size='8'><b>{disp}</b></font>", styles['Normal'])
                row = [col_p]
            else: row = [col]

            row_order_tot = 0
            for sc in size_cols:
                val = r[sc]
                if pd.notna(val) and val != "":
                    val_int = int(val)
                    if val_int == 0: row.append("")
                    else:
                        stock_k = f"{clean_str(sku)}||{clean_str(col)}||{clean_str(sc)}"
                        stock_qty = file_stock.get(stock_k, 0)
                        order_qty = max(0, val_int - stock_qty)
                        row_order_tot += order_qty
                        
                        if stock_qty > 0:
                            if order_qty > 0: cell_p = Paragraph(f"<font size='9'><b>{order_qty}</b></font> <font size='7' color='#86868b'>({stock_qty} mag)</font>", centered_style)
                            else: cell_p = Paragraph(f"<font size='8' color='#1976d2'><b>{stock_qty} mag</b></font>", centered_style)
                            row.append(cell_p)
                        else: row.append(str(order_qty) if order_qty > 0 else "")
                else: row.append("")
                    
            row.append(str(row_order_tot) if row_order_tot > 0 else "")
            data.append(row)

        t = Table(data, repeatRows=1, colWidths=col_widths, hAlign='LEFT')
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#fafafc")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#1d1d1f")),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#e5e5ea")),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (-1, 0), (-1, -1), 'Helvetica-Bold'),
        ]))
        block.append(t)
        block.append(Spacer(1, 6*mm))
        elements.append(KeepTogether(block))

    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()

# -------------------------
# Labels
# -------------------------
@dataclass
class LabelCfg:
    w_mm: float = 152.0
    h_mm: float = 102.0
    margin_mm: float = 8.0
    logo_w_mm: float = 28.0
    title_pt: float = 18.0
    header_pt: float = 9.0
    row_pt: float = 9.0
    row_h_mm: float = 4.2
    strip_modello: bool = False

def make_labels_pdf(df: pd.DataFrame, logo_bytes: bytes | None, cfg: LabelCfg) -> bytes:
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase.pdfmetrics import stringWidth

    w = cfg.w_mm * mm
    h = cfg.h_mm * mm
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))

    logo_img = None
    if logo_bytes:
        try: logo_img = ImageReader(io.BytesIO(logo_bytes))
        except Exception: logo_img = None

    for col in ["N. Ordine", "Nome Prodotto", "Colore", "Taglia", "Studente", "GruppoEtichetta", "Pezzi"]:
        if col not in df.columns: df[col] = ""

    groups = sorted(
        df["GruppoEtichetta"].dropna().unique().tolist(),
        key=lambda x: (0 if str(x) == "Docenti / ATA" else 1, str(x)),
    )

    x_left = cfg.margin_mm * mm
    x_right = w - cfg.margin_mm * mm
    y_top = h - cfg.margin_mm * mm
    logo_h = 12 * mm
    logo_w = (cfg.logo_w_mm * 1.15) * mm
    gap = 6 * mm
    gap_col = 1.2 * mm

    usable_w = x_right - x_left
    col_ord_x = x_left
    col_ord_w = 14 * mm
    remaining = usable_w - col_ord_w - (4 * gap_col)

    col_art_w  = remaining * 0.44
    col_col_w  = remaining * 0.18
    col_size_w = remaining * 0.12
    col_name_w = remaining * 0.26

    col_art_x  = col_ord_x + col_ord_w + gap_col
    col_col_x  = col_art_x + col_art_w + gap_col
    col_size_x = col_col_x + col_col_w + gap_col
    col_name_x = col_size_x + col_size_w + gap_col

    row_pt = max(7.0, float(cfg.row_pt))
    header_pt = max(7.5, float(getattr(cfg, "header_pt", row_pt)))
    row_h = max(3.6, float(getattr(cfg, "row_h_mm", 4.2))) * mm

    def fit(text: str, max_w: float, font: str, size: float) -> str:
        t = (text or "").strip()
        if not t: return ""
        if stringWidth(t, font, size) <= max_w: return t
        s = t
        while s and stringWidth(s + "…", font, size) > max_w: s = s[:-1]
        return (s + "…") if s else "…"

    def wrap_two_lines_right(t: str, max_w: float, font: str, size: float):
        t = (t or "").strip()
        if not t: return [""]
        if stringWidth(t, font, size) <= max_w: return [t]
        words = t.split()
        lines, cur = [], ""
        for w0 in words:
            cand = (cur + " " + w0).strip()
            if stringWidth(cand, font, size) <= max_w:
                cur = cand
            else:
                if cur: lines.append(cur)
                cur = w0
                if len(lines) == 1: break
        if len(lines) < 2: lines.append(cur)
        if stringWidth(lines[-1], font, size) > max_w:
            lines[-1] = fit(lines[-1], max_w, font, size)
        return lines[:2]

    def draw_header(title: str, page_idx: int, page_count: int):
        if logo_img:
            c.drawImage(logo_img, x_left, y_top - logo_h, width=logo_w, height=logo_h, preserveAspectRatio=True, mask="auto")
        title_area_left = x_left + logo_w + gap
        title_max_w = x_right - title_area_left
        c.setFont("Helvetica-Bold", cfg.title_pt)
        lines = wrap_two_lines_right(title, title_max_w, "Helvetica-Bold", cfg.title_pt)
        y_title = y_top - 6 * mm
        for i, line in enumerate(lines):
            c.drawRightString(x_right, y_title - i * (6.5 * mm), line)
        if page_count > 1:
            c.setFont("Helvetica", 8)
            c.drawRightString(x_right, cfg.margin_mm * mm + 3 * mm, f"{page_idx}/{page_count}")
        y = y_top - 30 * mm
        c.setFont("Helvetica-Bold", header_pt)
        c.drawString(col_ord_x, y, "Ordine")
        c.drawString(col_art_x, y, "Articolo")
        c.drawString(col_col_x, y, "Colore")
        c.drawString(col_size_x, y, "Taglia")
        c.drawString(col_name_x, y, "Nome Cognome")
        c.setFont("Helvetica", row_pt)
        return y - 6 * mm

    def paginate_rows(rows: pd.DataFrame) -> list[pd.DataFrame]:
        y0 = y_top - 36 * mm
        usable = y0 - (cfg.margin_mm * mm + 8 * mm)
        per_page = max(1, int(usable // row_h))
        return [rows.iloc[i : i + per_page].copy() for i in range(0, len(rows), per_page)]

    for grp in groups:
        gdf = df[df["GruppoEtichetta"] == grp].copy()
        gdf["__r"] = gdf["Taglia"].map(sort_size_key)
        gdf = gdf.sort_values(["Nome Prodotto", "Colore", "__r", "Taglia", "N. Ordine"], kind="stable")
        pages = paginate_rows(gdf)
        total_pages = len(pages)
        for pi, page_df in enumerate(pages, start=1):
            y = draw_header(str(grp), pi, total_pages)
            for _, r in page_df.iterrows():
                ordine = fit(str(r.get("N. Ordine", "") or ""), col_ord_w, "Helvetica", row_pt)
                raw_art = str(r.get("Nome Prodotto", "") or "").strip()
                if cfg.strip_modello: raw_art = re.sub(r"(?i)\bmodello\b\s*", "", raw_art)
                if len(raw_art) > 20: raw_art = raw_art[:20] + "…"
                articolo = fit(raw_art, col_art_w, "Helvetica", row_pt)
                colore = fit(str(r.get("Colore", "") or ""), col_col_w, "Helvetica", row_pt)
                taglia = fit(str(r.get("Taglia", "") or ""), col_size_w, "Helvetica", row_pt)
                persona_raw = str(r.get("Studente", "") or "").strip()
                name_size = row_pt
                while name_size > 6.5 and stringWidth(persona_raw, "Helvetica", name_size) > col_name_w: name_size -= 0.2
                c.drawString(col_ord_x, y, ordine)
                c.drawString(col_art_x, y, articolo)
                c.drawString(col_col_x, y, colore)
                c.drawString(col_size_x, y, taglia)
                c.setFont("Helvetica", name_size)
                c.drawString(col_name_x, y, persona_raw)
                c.setFont("Helvetica", row_pt)
                y -= row_h
            c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

# -------------------------
# Ordini da pagare (Pending)
# -------------------------
def normalize_pagamento(v) -> str:
    return clean_str(v).lower()

def to_number_it(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)): return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    if not s: return 0.0
    import re as _re
    s = _re.sub(r"[^\d\.,\-]", "", s)
    has_dot = "." in s
    has_comma = "," in s
    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif has_comma and not has_dot:
        s = s.replace(".", "").replace(",", ".")
    else:
        if _re.match(r"^-?\d{1,3}(\.\d{3})+(\.\d+)?$", s):
            parts = s.split(".")
            if len(parts) > 2:
                dec = parts.pop()
                s = "".join(parts) + "." + dec
            else: s = s.replace(".", "")
    try: return float(s)
    except Exception: return 0.0

def build_pending_model(df: pd.DataFrame) -> dict:
    d = df.copy()
    importo_col = "Importo ordine" 
    for col in d.columns:
        c_norm = str(col).lower().replace("_", " ").strip()
        if any(x in c_norm for x in ["importo", "totale", "prezzo", "da pagare"]):
            importo_col = col
            break
    required = ["N. Ordine","Classe","Cognome Studente","Nome Studente","Docente/ATA","Pagamento", importo_col]
    for c in required:
        if c not in d.columns: d[c] = ""
        d[c] = d[c].map(clean_str)
    pending_keywords = ["pending", "in attesa", "non pagato", "da pagare", "on-hold", "on hold"]
    pending = d[d["Pagamento"].map(normalize_pagamento).isin(pending_keywords)].copy()
    seen = {}
    for _, r in pending.iterrows():
        order_id = clean_str(r["N. Ordine"])
        if not order_id or order_id in seen: continue
        raw_classe = clean_str(r["Classe"])
        is_doc = (raw_classe == "")
        classe = "Docenti/ATA" if is_doc else raw_classe
        docente = clean_str(r["Docente/ATA"])
        cognome = clean_str(r["Cognome Studente"])
        nome = clean_str(r["Nome Studente"])
        display_name = docente if is_doc else clean_str(f"{cognome} {nome}")
        if not display_name: display_name = "(Docente/ATA non indicato)" if is_doc else "(Studente non indicato)"
        seen[order_id] = {
            "orderId": order_id, "classe": classe if classe else "(Senza classe)",
            "displayName": display_name, "importo": to_number_it(r[importo_col]),
        }
    unique_orders = list(seen.values())
    by_class = {}
    for o in unique_orders: by_class.setdefault(o["classe"], []).append(o)
    classes = sorted(by_class.keys(), key=lambda x: (0 if x == "Docenti/ATA" else 1, str(x)))
    for c in classes: by_class[c].sort(key=lambda o: (o["displayName"] or "").lower())
    summary = []
    grand_total = 0.0
    for c in classes:
        arr = by_class[c]
        tot = sum(x["importo"] for x in arr)
        grand_total += tot
        summary.append({"Classe": c, "N ordini": len(arr), "Totale €": round(tot, 2)})
    return { "classes": classes, "by_class": by_class, "summary": pd.DataFrame(summary), "grand_total": round(grand_total, 2), "n_orders": len(unique_orders) }

def pending_pdf_per_class_students(model: dict) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase.pdfmetrics import stringWidth
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    margin = 18 * mm
    row_h = 6 * mm
    def fmt_eur(x: float) -> str: return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    for classe in model["classes"]:
        y = h - margin
        c.setFont("Helvetica-Bold", 18)
        c.drawString(margin, y, f"Classe: {classe}")
        y -= 10 * mm
        c.setFont("Helvetica-Bold", 11)
        c.drawString(margin, y, "Studente")
        c.drawRightString(w - margin, y, "Importo (€)")
        y -= 7 * mm
        c.setFont("Helvetica", 11)
        tot_cls = 0.0
        tmp = {}
        for o in model["by_class"][classe]:
            nm = o.get("displayName") or ""
            tmp[nm] = tmp.get(nm, 0.0) + float(o.get("importo") or 0.0)
        rows = sorted(tmp.items(), key=lambda x: (x[0] or "").lower())
        for name, imp in rows:
            if y < margin + 18 * mm:
                c.showPage()
                y = h - margin
                c.setFont("Helvetica-Bold", 18)
                c.drawString(margin, y, f"Classe: {classe}")
                y -= 10 * mm
                c.setFont("Helvetica-Bold", 11)
                c.drawString(margin, y, "Studente")
                c.drawRightString(w - margin, y, "Importo (€)")
                y -= 7 * mm
                c.setFont("Helvetica", 11)
            max_w = (w - 2 * margin) - 35 * mm
            nm = (name or "").strip() or "(Nome non indicato)"
            while stringWidth(nm, "Helvetica", 11) > max_w and len(nm) > 2: nm = nm[:-1]
            if nm != (name or "").strip(): nm = nm.rstrip() + "…"
            c.drawString(margin, y, nm)
            c.drawRightString(w - margin, y, fmt_eur(imp))
            y -= row_h
            tot_cls += imp
        y -= 4 * mm
        c.setFont("Helvetica-Bold", 14)
        c.drawString(margin, y, "TOTALE CLASSE")
        c.drawRightString(w - margin, y, fmt_eur(tot_cls))
        c.showPage()
    y = h - margin
    c.setFont("Helvetica-Bold", 22)
    c.drawString(margin, y, "TOTALE GENERALE DA RACCOGLIERE")
    y -= 14 * mm
    c.setFont("Helvetica-Bold", 28)
    c.drawString(margin, y, f"€ {fmt_eur(float(model.get('grand_total') or 0.0))}")
    c.save()
    buf.seek(0)
    return buf.getvalue()

def pending_pdf_totals_only(model: dict) -> bytes:
    from reportlab.lib.pagesizes import A4
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    margin = 18 * mm
    y = h - margin
    def fmt_eur(x: float) -> str: return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    c.setFont("Helvetica-Bold", 18)
    c.drawString(margin, y, "Riepilogo ordini pending")
    y -= 10 * mm
    c.setFont("Helvetica", 11)
    c.drawString(margin, y, "Totali per classe (Pagamento = pending)")
    y -= 10 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin, y, "Classe")
    c.drawRightString(w - margin, y, "Totale da raccogliere (€)")
    y -= 8 * mm
    c.setFont("Helvetica", 12)
    class_totals = []
    for classe in model["classes"]:
        tot = sum(float(o.get("importo") or 0.0) for o in model["by_class"][classe])
        class_totals.append((classe, tot))
    row_h = 7 * mm
    for classe, tot in class_totals:
        if y < margin + 25 * mm:
            c.showPage()
            y = h - margin
            c.setFont("Helvetica-Bold", 18)
            c.drawString(margin, y, "Riepilogo ordini pending")
            y -= 20 * mm
            c.setFont("Helvetica-Bold", 12)
            c.drawString(margin, y, "Classe")
            c.drawRightString(w - margin, y, "Totale da raccogliere (€)")
            y -= 8 * mm
            c.setFont("Helvetica", 12)
        c.drawString(margin, y, str(classe))
        c.drawRightString(w - margin, y, fmt_eur(tot))
        y -= row_h
    y -= 8 * mm
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, y, "TOTALE GENERALE")
    c.drawRightString(w - margin, y, fmt_eur(float(model.get("grand_total") or 0.0)))
    c.save()
    buf.seek(0)
    return buf.getvalue()

def page_pending(df_target: pd.DataFrame) -> None:
    st.subheader("💸 Ordini da pagare")
    st.caption("Estrae gli ordini non saldati, consolida per N. Ordine e raggruppa per classe.")
    if "Pagamento" not in df_target.columns and "N. Ordine" not in df_target.columns:
        st.error("Mancano le colonne chiave 'N. Ordine' e 'Pagamento' nel file.")
        return
    model = build_pending_model(df_target)
    top1, top2, top3 = st.columns([2,2,3])
    with top1: st.metric("Ordini pending", model["n_orders"])
    with top2: st.metric("Totale €", f'{model["grand_total"]:.2f}')
    with top3:
        if model["n_orders"] > 0:
            pdfA = pending_pdf_per_class_students(model)
            pdfB = pending_pdf_totals_only(model)
            st.download_button("⬇️ PDF per classe", data=pdfA, file_name="wupi_ordini_pending_dettaglio.pdf", mime="application/pdf")
            st.download_button("⬇️ PDF riepilogo", data=pdfB, file_name="wupi_ordini_pending_riepilogo.pdf", mime="application/pdf")
    st.markdown("### Riepilogo per classe")
    if not model["summary"].empty:
        st.dataframe(model["summary"], use_container_width=True, hide_index=True)
    else: st.info("Nessun ordine pending trovato in questo file.")
    st.markdown("### Dettaglio")
    cls = st.selectbox("Classe", options=model["classes"], index=0 if model["classes"] else None)
    if cls:
        det = pd.DataFrame(model["by_class"][cls])
        det = det.rename(columns={"orderId":"N. Ordine","displayName":"Nome","importo":"Importo €"})[["N. Ordine","Nome","Importo €"]]
        st.dataframe(det, use_container_width=True, hide_index=True)

# -------------------------
# Bibbia maker (A3)
# -------------------------
@dataclass
class BibbiaCfg:
    margin_mm: float = 10.0
    gap_mm: float = 6.0
    header_pt: float = 18.0
    caption_pt: float = 11.0
    show_missing_boxes: bool = True

def _norm_key(s: str) -> str:
    s = clean_str(s).lower()
    s = s.replace("-", "_").replace(" ", "_")
    s = re.sub(r"[\s]+", "_", s)
    s = re.sub(r"[^a-z0-9_]+", "", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _sku_base(sku: str) -> str:
    s = clean_str(sku)
    if "_" in s:
        p1, p2 = s.split("_", 1)
        if 1 <= len(p1) <= 2: return p2
    return s

def sku_base_key(sku: str) -> str:
    s = clean_str(sku).upper()
    if not s: return ""
    s = re.sub(r"[^A-Z0-9]+", "_", s)
    parts = [p for p in s.split("_") if p]
    if not parts: return ""
    return parts[-1]

def product_model_key(nome_prodotto: str) -> str:
    s = clean_str(nome_prodotto)
    if not s: return ""
    part = s.split("|", 1)[1].strip() if "|" in s else s
    part_low = part.lower()
    part_low = re.sub(r"^\s*modello\s+", "", part_low, flags=re.IGNORECASE)
    clothing_pattern = r"(?i)\b(hoodie|t\-shirt|tshirt|shirt|sweatshirt|felpa|maglia|maglietta|pant|pants|pantalone|pantaloni|short|shorts|zip|crew|giacca|jacket|kway|k\-way|polo|penne|kit)\b"
    part_low = re.sub(clothing_pattern, "", part_low)
    part_low = re.sub(r"[^a-z0-9]+", "_", part_low)
    part_low = re.sub(r"_+", "_", part_low).strip("_")
    return part_low

def find_mockup_bytes(mock_map: dict, sku_key: str, model_key: str, col_key: str, side: str) -> bytes | None:
    side = side or ""
    model_key = (model_key or "").strip()
    k = (sku_key, model_key, col_key, side)
    if k in mock_map: return mock_map[k]
    k = (sku_key, "", col_key, side)
    if k in mock_map: return mock_map[k]
    if side == "fronte":
        k = (sku_key, model_key, col_key, "")
        if k in mock_map: return mock_map[k]
        k = (sku_key, "", col_key, "")
        if k in mock_map: return mock_map[k]
    return None

def parse_mockup_files(files: list) -> dict:
    def norm_token(t: str) -> str:
        t = (t or "").strip().lower()
        t = re.sub(r"[^a-z0-9]+", "_", t)
        return re.sub(r"_+", "_", t).strip("_")
    alias2canon = {}
    for a, canon in COLOR_ALIAS_MAP.items():
        alias2canon[norm_token(a)] = canon
        alias2canon[norm_token(canon)] = canon
    def detect_color(tokens: list[str]) -> tuple[str, int | None, str]:
        if not tokens: return "", None, ""
        max_len = min(3, len(tokens))
        for L in range(max_len, 0, -1):
            raw = norm_token("_".join(tokens[-L:]))
            if not raw: continue
            canon = alias2canon.get(raw)
            if canon: return canon, len(tokens) - L, raw
            if L == 1: return raw, len(tokens) - 1, raw
        for i, tok in enumerate(tokens):
            key = norm_token(tok)
            if key in alias2canon: return alias2canon[key], i, key
        raw = norm_token(tokens[-1])
        return raw, len(tokens) - 1, raw
    def detect_side(tokens: list[str]) -> tuple[str, int | None]:
        for i, tok in enumerate(tokens):
            k = norm_token(tok)
            if k in ("fronte", "front", "f"): return "fronte", i
            if k in ("retro", "back", "r"): return "retro", i
        return "", None
    out = {}
    for f in files:
        name = getattr(f, "name", "")
        base = Path(name).stem
        parts = [p for p in re.split(r"[\s_\-]+", base) if p]
        if not parts: continue
        sku_tok = parts[0]
        sku_key = _norm_key(sku_base_key(sku_tok))
        rest = parts[1:]
        side_key, side_i = detect_side(rest)
        rest_wo_side = [t for j, t in enumerate(rest) if not (side_i is not None and j == side_i)]
        color_key, color_start, raw_color = detect_color(rest_wo_side)
        model_tokens = rest_wo_side[:color_start] if color_start is not None else []
        model_key = product_model_key(" ".join(model_tokens))
        try: data = f.getvalue()
        except Exception:
            try: data = f.read()
            except Exception: continue
        out[(sku_key, model_key, color_key, side_key)] = data
        if raw_color and raw_color != color_key: out[(sku_key, model_key, raw_color, side_key)] = data
    return out

def bibbia_variants(df_norm: pd.DataFrame) -> pd.DataFrame:
    d = df_norm.copy()
    for col in ["SKU", "Nome Prodotto", "Colore", "Taglia", "Pezzi", "Nome incisione", "N. Ordine", "Classe"]:
        if col not in d.columns: d[col] = ""
    d["SKU_BASE"] = d["SKU"].map(_sku_base)
    d["SKU_KEY"] = d["SKU_BASE"].map(_norm_key)
    d["COL_KEY"] = d["Colore"].map(color_to_canon_key)
    d["MODEL_KEY"] = d["Nome Prodotto"].map(product_model_key)
    dd = d.copy()
    dd.loc[dd["Taglia"].eq(""), "Taglia"] = "UNICA"
    breakdown = dd.groupby(["SKU", "Nome Prodotto", "Colore", "Taglia"], as_index=False)["Pezzi"].sum().sort_values(["SKU", "Nome Prodotto", "Colore", "Taglia"], kind="stable")
    def fmt_breakdown(g: pd.DataFrame) -> str:
        items = [(str(r["Taglia"]), int(r["Pezzi"])) for _, r in g.iterrows() if int(r["Pezzi"]) > 0]
        items.sort(key=lambda x: sort_size_key(x[0]))
        return " ".join([f"{t}:{q}" for t, q in items])
    bd = breakdown.groupby(["SKU", "Nome Prodotto", "Colore"], as_index=False).apply(fmt_breakdown)
    if isinstance(bd, pd.Series): bd = bd.reset_index().rename(columns={0: "Taglie"})
    else: bd = bd.rename(columns={None: "Taglie"})
    if "Taglie" not in bd.columns: bd["Taglie"] = ""

    tmp = d.copy()
    tmp["INC_RAW"] = tmp["Nome incisione"].map(clean_str)
    tmp["ORD_SHOW"] = tmp["N. Ordine"].map(clean_str)
    tmp["CLS_SHOW"] = tmp["Classe"].map(clean_str)
    tmp.loc[tmp["CLS_SHOW"].eq(""), "CLS_SHOW"] = "Docenti / ATA"
    clothing_pattern = r"(?i)\b(hoodie|t\-shirt|tshirt|shirt|sweatshirt|felpa|maglia|maglietta|pant|pants|pantalone|pantaloni|short|shorts|zip|crew|giacca|jacket|kway|k\-way|polo|penne|kit)\b"
    is_clothing = tmp["Nome Prodotto"].astype(str).str.contains(clothing_pattern, regex=True)
    tmp.loc[is_clothing, "INC_RAW"] = ""
    def fmt_incisioni(g: pd.DataFrame) -> str:
        g = g.copy()
        personalizzati = g[g["INC_RAW"].ne("")]
        if personalizzati.empty: return ""
        neutri_qty = int(g[g["INC_RAW"].eq("")]["Pezzi"].sum())
        pers_qty = int(personalizzati["Pezzi"].sum())
        rows = [f"Neutri: {neutri_qty} pz", f"Personalizzati: {pers_qty} pz"]
        agg = personalizzati.groupby(["ORD_SHOW", "CLS_SHOW", "INC_RAW"], as_index=False)["Pezzi"].sum().sort_values(["ORD_SHOW", "CLS_SHOW", "INC_RAW"], kind="stable")
        for _, r in agg.iterrows():
            ordn = clean_str(r["ORD_SHOW"])
            cls = clean_str(r["CLS_SHOW"])
            inc = clean_str(r["INC_RAW"])
            qty = int(r["Pezzi"])
            base = " · ".join([x for x in [f"#{ordn}" if ordn else "", cls, inc] if x])
            rows.append(f"{base} ({qty})" if qty > 1 else base)
        return "\n".join(rows)
    inc = tmp.groupby(["SKU", "Nome Prodotto", "Colore"], as_index=False).apply(fmt_incisioni)
    if isinstance(inc, pd.Series): inc = inc.reset_index().rename(columns={0: "Incisioni"})
    else: inc = inc.rename(columns={None: "Incisioni"})
    if "Incisioni" not in inc.columns: inc["Incisioni"] = ""

    tot = d.groupby(["SKU", "Nome Prodotto", "Colore", "SKU_KEY", "COL_KEY", "MODEL_KEY"], as_index=False)["Pezzi"].sum().rename(columns={"Pezzi": "Totale"})
    out = tot.merge(bd[["SKU", "Nome Prodotto", "Colore", "Taglie"]], on=["SKU", "Nome Prodotto", "Colore"], how="left")
    out = out.merge(inc[["SKU", "Nome Prodotto", "Colore", "Incisioni"]], on=["SKU", "Nome Prodotto", "Colore"], how="left")
    out["Taglie"] = out["Taglie"].fillna("")
    out["Incisioni"] = out["Incisioni"].fillna("")
    out = out.sort_values(["SKU", "Nome Prodotto", "Colore"], kind="stable").reset_index(drop=True)
    return out

def _draw_image_fit(c: canvas.Canvas, img_bytes: bytes, x: float, y: float, w: float, h: float):
    img = ImageReader(io.BytesIO(img_bytes))
    iw, ih = img.getSize()
    if iw <= 0 or ih <= 0: return
    scale = min(w / iw, h / ih)
    dw, dh = iw * scale, ih * scale
    dx = x + (w - dw) / 2
    dy = y + (h - dh) / 2
    c.drawImage(img, dx, dy, width=dw, height=dh, preserveAspectRatio=True, mask="auto")

def _draw_bibbia_variant(c: canvas.Canvas, r: pd.Series, mock_map: dict, cfg: BibbiaCfg, logo_img, w: float, h: float):
    mm_to_pt = mm
    margin = cfg.margin_mm * mm_to_pt
    gap = cfg.gap_mm * mm_to_pt

    sku = str(r.get("SKU", ""))
    sku_base = _sku_base(sku)
    sku_key = str(r.get("SKU_KEY", _norm_key(sku_base)))
    prod = str(r.get("Nome Prodotto", ""))
    model_key = str(r.get("MODEL_KEY", product_model_key(prod)))
    col = str(r.get("Colore", ""))
    col_key = str(r.get("COL_KEY", _norm_key(col)))

    taglie = str(r.get("Taglie", ""))
    incisioni = str(r.get("Incisioni", ""))
    totale = int(r.get("Totale", 0))
    has_incisioni = clean_str(incisioni) != ""

    header_h = 24 * mm_to_pt
    footer_h = 34 * mm_to_pt

    if logo_img: c.drawImage(logo_img, margin, h - margin - 16 * mm_to_pt, width=26 * mm_to_pt, height=16 * mm_to_pt, preserveAspectRatio=True, mask="auto")

    c.setFillGray(0)
    c.setFont("Helvetica-Bold", cfg.header_pt)
    c.drawString(margin + (32 * mm_to_pt if logo_img else 0), h - margin - 10 * mm_to_pt, f"{sku_base} — {prod}")
    c.drawRightString(w - margin, h - margin - 10 * mm_to_pt, f"{col}")

    img_y0 = margin + footer_h + gap
    img_h = h - margin - header_h - img_y0
    img_w = (w - 2 * margin - gap) / 2
    box1 = (margin, img_y0, img_w, img_h)
    box2 = (margin + img_w + gap, img_y0, img_w, img_h)

    front = find_mockup_bytes(mock_map, sku_key, model_key, col_key, "fronte")
    back = find_mockup_bytes(mock_map, sku_key, model_key, col_key, "retro")

    if front: _draw_image_fit(c, front, *box1)
    elif cfg.show_missing_boxes:
        c.setLineWidth(1); c.setStrokeGray(0.8); c.rect(*box1)
        c.setFillGray(0.5); c.setFont("Helvetica", 14)
        c.drawCentredString(box1[0] + box1[2] / 2, box1[1] + box1[3] / 2, "FRONTE MANCANTE")

    if back: _draw_image_fit(c, back, *box2)
    elif cfg.show_missing_boxes:
        c.setLineWidth(1); c.setStrokeGray(0.8); c.rect(*box2)
        c.setFillGray(0.5); c.setFont("Helvetica", 14)
        c.drawCentredString(box2[0] + box2[2] / 2, box2[1] + box2[3] / 2, "RETRO MANCANTE")

    c.setFillGray(0)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, margin + footer_h - 8 * mm_to_pt, f"SKU: {sku_base}   Colore: {col}   Totale: {totale}")

    box_w = 82 * mm_to_pt if has_incisioni else 0
    pills_left_x = margin
    pills_y = margin + footer_h - 20 * mm_to_pt
    pills_max_w = (w - 2 * margin - box_w - 10 * mm_to_pt) if has_incisioni else (w - 2 * margin)

    items = _parse_taglie_items(taglie)
    cur_x = pills_left_x

    if items:
        for taglia, qty in items:
            c.setFont("Helvetica", 15)
            w1 = c.stringWidth(str(taglia), "Helvetica", 15)
            c.setFont("Helvetica-Bold", 15)
            w2 = c.stringWidth(str(qty), "Helvetica-Bold", 15)
            pw = w1 + w2 + 24 + 8
            if cur_x + pw > pills_left_x + pills_max_w: break
            c.setFillColorRGB(0.92, 0.92, 0.93)
            c.roundRect(cur_x, pills_y - 6, pw, 26, 13, stroke=0, fill=1)
            c.setFillGray(0)
            c.setFont("Helvetica", 15)
            c.drawString(cur_x + 12, pills_y + 2, str(taglia))
            c.setFont("Helvetica-Bold", 15)
            c.drawString(cur_x + 12 + w1 + 8, pills_y + 2, str(qty))
            cur_x += pw + 12

    if has_incisioni:
        bx = w - margin - box_w
        by = margin
        box_h = footer_h
        c.setFillColorRGB(0.94, 0.94, 0.95)
        c.roundRect(bx, by, box_w, box_h, 8, stroke=0, fill=1)
        c.setFillGray(0)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(bx + 12, by + box_h - 16, "PERSONALIZZAZIONI")
        lines = incisioni.split("\n")
        if len(lines) >= 2:
            c.setFont("Helvetica-Bold", 9)
            c.drawString(bx + 12, by + box_h - 30, f"{lines[0]}   |   {lines[1]}")
            c.setLineWidth(0.5); c.setStrokeColorRGB(0.85, 0.85, 0.85)
            c.line(bx + 12, by + box_h - 36, bx + box_w - 12, by + box_h - 36)
            y = by + box_h - 50
            for line in lines[2:]:
                if y < by + 8: break
                txt = line.strip()
                if not txt: continue
                c.setFillColorRGB(0.2, 0.2, 0.2)
                c.circle(bx + 15, y + 2.5, 1.5, stroke=0, fill=1)
                c.setFillGray(0); c.setFont("Helvetica", 9)
                while c.stringWidth(txt, "Helvetica", 9) > (box_w - 36) and len(txt) > 2: txt = txt[:-1]
                if txt != line.strip(): txt = txt.rstrip() + "…"
                c.drawString(bx + 22, y, txt)
                y -= 12

def make_bibbia_pdf_single(variants: pd.DataFrame, mock_map: dict, cfg: BibbiaCfg, brand_logo: bytes | None = None) -> bytes:
    w, h = landscape(A3)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    logo_img = None
    if brand_logo:
        try: logo_img = ImageReader(io.BytesIO(brand_logo))
        except Exception: logo_img = None
    for _, r in variants.iterrows():
        _draw_bibbia_variant(c, r, mock_map, cfg, logo_img, w, h)
        c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

def make_bibbia_pdf_grid(variants: pd.DataFrame, mock_map: dict, cfg: BibbiaCfg, brand_logo: bytes | None = None) -> bytes:
    from reportlab.lib.pagesizes import A3
    page_w, page_h = A3 
    vw, vh = landscape(A3) 
    cols, rows = 2, 4
    cell_w, cell_h = page_w / cols, page_h / rows
    scale_factor = cell_w / vw 
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    logo_img = None
    if brand_logo:
        try: logo_img = ImageReader(io.BytesIO(brand_logo))
        except Exception: logo_img = None
    count = 0
    for _, r in variants.iterrows():
        px = (count % cols) * cell_w
        py = page_h - ((count // cols) % rows + 1) * cell_h
        c.saveState()
        c.translate(px, py)
        c.scale(scale_factor, scale_factor)
        _draw_bibbia_variant(c, r, mock_map, cfg, logo_img, vw, vh)
        c.restoreState()
        c.setLineWidth(0.5); c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.rect(px, py, cell_w, cell_h)
        count += 1
        if count % (cols * rows) == 0: c.showPage()
    if count % (cols * rows) != 0: c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

@st.dialog("Gestione SKU Sostitutivi")
def substitute_modal(sku_color_options, proj_dir: Path):
    if "sub_idx" not in st.session_state: st.session_state.sub_idx = 0
    if st.session_state.sub_idx >= len(sku_color_options): st.session_state.sub_idx = 0
        
    c_nav1, c_nav2 = st.columns([4, 1])
    with c_nav1:
        sel_val = st.selectbox("Seleziona rapidamente (oppure usa le frecce):", options=sku_color_options, index=st.session_state.sub_idx)
        if sel_val != sku_color_options[st.session_state.sub_idx]:
            st.session_state.sub_idx = sku_color_options.index(sel_val)
            st.rerun()
    with c_nav2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("❌ Chiudi", use_container_width=True):
            st.session_state.show_sub_modal = False
            st.rerun()

    current_sel = sku_color_options[st.session_state.sub_idx]
    sku, color = [x.strip() for x in current_sel.split(" — ", 1)]
    k = f"{clean_str(sku)}||{clean_str(color)}"
    
    subs = load_subs(proj_dir)
    sub_data = subs.get(k, {})
    if isinstance(sub_data, str): sub_data = {"fornitore": "Altro", "sku": sub_data}
        
    # Memoria intelligente per ricordarsi sempre l'ultimo inserito nei passaggi "Avanti" o "Indietro"
    last_forn = st.session_state.get("last_sub_forn", "ActionWear")
    last_sku = st.session_state.get("last_sub_sku", "")
    
    curr_forn = sub_data.get("fornitore", last_forn) if sub_data else last_forn
    curr_sku = sub_data.get("sku", last_sku) if sub_data else last_sku

    custom_sups = load_custom_suppliers()
    base_sups = ["ActionWear", "Basic", "Roly", "Top-Tex", "Stanley/Stella", "Innova"]
    all_sups = base_sups + custom_sups + ["Altro"]
    idx_forn = all_sups.index(curr_forn) if curr_forn in all_sups else all_sups.index("Altro")
    
    with st.form(f"sub_form_{st.session_state.sub_idx}", clear_on_submit=False):
        st.markdown(f"Stai modificando: **{sku}** — Variante **{color}**")
        f1, f2 = st.columns(2)
        with f1:
            sel_forn = st.selectbox("Fornitore", options=all_sups, index=idx_forn)
            altro_forn = st.text_input("Specifica nuovo fornitore:", value=curr_forn if curr_forn not in base_sups+custom_sups else "") if sel_forn == "Altro" else ""
        with f2: 
            new_sku = st.text_input("SKU Sostitutivo", value=curr_sku, placeholder="Es. DOG123")
            
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
        b1, b2, b3 = st.columns(3)
        
        # IL TRUCCO DEL TASTO INVIO: Streamlit assegna il tasto Invio al PRIMO bottone dichiarato nel codice!
        # Noi lo mettiamo visivamente in b2 (al centro), ma lo scriviamo per primo nel codice!
        with b2: btn_save = st.form_submit_button("💾 Salva e Chiudi", type="primary", use_container_width=True)
        with b1: btn_prev = st.form_submit_button("⬅️ Salva e Precedente", use_container_width=True)
        with b3: btn_next = st.form_submit_button("Salva e Successivo ➡️", use_container_width=True)
        
        if btn_prev or btn_save or btn_next:
            final_forn = altro_forn.strip() if sel_forn == "Altro" else sel_forn
            final_sku = new_sku.strip()
            
            st.session_state["last_sub_forn"] = final_forn
            st.session_state["last_sub_sku"] = final_sku
            
            if final_forn and final_forn not in base_sups and final_forn not in custom_sups and final_forn != "Altro":
                custom_sups.append(final_forn)
                save_custom_suppliers(custom_sups)
            
            if final_sku: subs[k] = {"fornitore": final_forn, "sku": final_sku}
            else: subs.pop(k, None)
            
            save_subs(proj_dir, subs)
            
            # Se premi Avanti o Indietro, naviga. Se premi Salva e Chiudi, spegne la modale!
            if btn_prev: 
                st.session_state.sub_idx = max(0, st.session_state.sub_idx - 1)
            elif btn_next: 
                st.session_state.sub_idx = min(len(sku_color_options)-1, st.session_state.sub_idx + 1)
            elif btn_save:
                st.session_state.show_sub_modal = False
            
            st.rerun()

def page_bibbia(df_norm: pd.DataFrame) -> None:
    st.subheader("Bibbia maker (A3)")
    st.caption("Carica i mockup in batch (JPG/PNG) con naming permissivo: SKU_modello_colore_fronte / SKU_modello_colore_retro.")

    if "bibbia_uploader_ver" not in st.session_state: st.session_state["bibbia_uploader_ver"] = 0

    mock_files = st.file_uploader("Carica qui le immagini Mockup", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"bibbia_mockups_{st.session_state['bibbia_uploader_ver']}")
    
    col_btn, _ = st.columns([1, 4])
    with col_btn:
        if st.button("🗑 Svuota immagini", use_container_width=True):
            st.session_state["bibbia_uploader_ver"] += 1
            st.rerun()
            
    st.markdown("<br>", unsafe_allow_html=True)
    mock_map = parse_mockup_files(mock_files or [])
    mock_map.update(load_manual_mockups())

    variants = bibbia_variants(df_norm)
    variants["Fronte"] = variants.apply(lambda r: "✅" if find_mockup_bytes(mock_map, r["SKU_KEY"], r.get("MODEL_KEY",""), r["COL_KEY"], "fronte") is not None else "❌", axis=1)
    variants["Retro"]  = variants.apply(lambda r: "✅" if find_mockup_bytes(mock_map, r["SKU_KEY"], r.get("MODEL_KEY",""), r["COL_KEY"], "retro") is not None else "❌", axis=1)

    k1, k2, k3 = st.columns(3)
    with k1: st.metric("Varianti", len(variants))
    with k2: st.metric("Complete (F+R)", int(((variants["Fronte"] == "✅") & (variants["Retro"] == "✅")).sum()))
    with k3: st.metric("Con mancanti", int(((variants["Fronte"] == "❌") | (variants["Retro"] == "❌")).sum()))

    st.markdown("### Controllo match")
    st.markdown("""<style>.bibbia-header { font-weight: 600; padding-bottom: 8px; border-bottom: 2px solid #e5e5ea; color: #86868b; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }</style>""", unsafe_allow_html=True)
    
    h1, h2, h3, h4, h5, h6 = st.columns([1.5, 2.5, 3, 2.5, 1, 1], vertical_alignment="bottom")
    with h1: st.markdown("<div class='bibbia-header'>SKU / Colore</div>", unsafe_allow_html=True)
    with h2: st.markdown("<div class='bibbia-header'>Prodotto</div>", unsafe_allow_html=True)
    with h3: st.markdown("<div class='bibbia-header'>Taglie</div>", unsafe_allow_html=True)
    with h4: st.markdown("<div class='bibbia-header'>Incisioni</div>", unsafe_allow_html=True)
    with h5: st.markdown("<div class='bibbia-header'>Fronte</div>", unsafe_allow_html=True)
    with h6: st.markdown("<div class='bibbia-header'>Retro</div>", unsafe_allow_html=True)

    for idx, r in variants.iterrows():
        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 2.5, 3, 2.5, 1, 1], vertical_alignment="center")
        with c1: st.markdown(f"**{r['SKU']}**<br><span style='color:#555; font-size:14px;'>{r['Colore']}</span>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div style='font-weight:500; margin-top:2px;'>{r['Nome Prodotto']}</div>", unsafe_allow_html=True)
        with c3:
            items = _parse_taglie_items(r['Taglie'])
            if items:
                pills = "".join([f'<span class="chip" style="margin-bottom:4px;">{t} <span class="q">{q}</span></span>' for t, q in items])
                st.markdown(f"<div class='chips'>{pills}</div>", unsafe_allow_html=True)
        with c4:
            inc_html = str(r['Incisioni']).replace('\n', '<br>')
            st.markdown(f"<div style='font-size:12px; color:#666; line-height:1.3;'>{inc_html}</div>", unsafe_allow_html=True)
        with c5:
            if r["Fronte"] == "✅": st.markdown("✅")
            else:
                if st.button("❌", key=f"f_{idx}_{r['SKU_KEY']}_{r['COL_KEY']}_{r.get('MODEL_KEY','')}", help="Aggiungi Fronte"): upload_missing_modal(r, "fronte")
        with c6:
            if r["Retro"] == "✅": st.markdown("✅")
            else:
                if st.button("❌", key=f"r_{idx}_{r['SKU_KEY']}_{r['COL_KEY']}_{r.get('MODEL_KEY','')}", help="Aggiungi Retro"): upload_missing_modal(r, "retro")
        st.markdown("<hr style='margin:0.25em 0; border:none; border-bottom:1px solid #f0f0f2;'>", unsafe_allow_html=True)

    with st.expander("Opzioni PDF A3", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1: margin_mm = st.number_input("Margine (mm)", min_value=4.0, max_value=30.0, value=10.0, step=1.0)
        with c2: gap_mm = st.number_input("Spazio tra fronte/retro (mm)", min_value=2.0, max_value=30.0, value=6.0, step=1.0)
        with c3: show_missing = st.checkbox("Mostra riquadri MANCANTE", value=True)

    with st.expander("Font", expanded=False):
        f1, f2 = st.columns(2)
        with f1: header_pt = st.number_input("Titolo (pt)", min_value=12.0, max_value=36.0, value=18.0, step=0.5)
        with f2: caption_pt = st.number_input("Caption (pt)", min_value=8.0, max_value=20.0, value=11.0, step=0.5)

    cfg = BibbiaCfg(margin_mm=float(margin_mm), gap_mm=float(gap_mm), header_pt=float(header_pt), caption_pt=float(caption_pt), show_missing_boxes=bool(show_missing))
    logo_bytes = (LOGO_PATH.read_bytes() if LOGO_PATH.exists() else None)

    st.markdown("### Generazione PDF")
    colA, colB = st.columns(2)
    with colA:
        if st.button("📄 Prepara PDF Singolo (1 per A3)", use_container_width=True):
            st.session_state['bibbia_mode'] = 'single'; st.rerun()
    with colB:
        if st.button("🗂 Prepara PDF Griglia (8 per A3)", type="primary", use_container_width=True):
            st.session_state['bibbia_mode'] = 'grid'; st.rerun()
            
    mode = st.session_state.get('bibbia_mode')
    if mode == 'single':
        pdf = make_bibbia_pdf_single(variants, mock_map, cfg, brand_logo=logo_bytes)
        st.success("PDF Singolo pronto per il download!")
        st.download_button("⬇️ Scarica PDF Singolo", data=pdf, file_name="wupi_bibbia_singola.pdf", mime="application/pdf", use_container_width=True)
    elif mode == 'grid':
        pdf = make_bibbia_pdf_grid(variants, mock_map, cfg, brand_logo=logo_bytes)
        st.success("PDF Griglia pronto per il download!")
        st.download_button("⬇️ Scarica PDF Griglia (8 in 1)", data=pdf, file_name="wupi_bibbia_griglia.pdf", mime="application/pdf", use_container_width=True)

def finance_summary(df_norm: pd.DataFrame, costs: Dict[str, float] | None = None):
    costs = costs or {}
    d = df_norm.copy()
    for col in ["SKU", "Pezzi", "Prezzo unitario", "Prezzo acquisto", "Classe", "Nome Prodotto", "Colore", "N. Ordine"]:
        if col not in d.columns: d[col] = 0 if col in ("Pezzi", "Prezzo unitario", "Prezzo acquisto") else ""
    d["Pezzi"] = pd.to_numeric(d["Pezzi"], errors="coerce").fillna(0).astype(int)
    d["Prezzo unitario"] = pd.to_numeric(d["Prezzo unitario"], errors="coerce").fillna(0.0)
    d["Prezzo vendita ex IVA"] = d["Prezzo unitario"] / 1.22
    d["Prezzo acquisto"] = pd.to_numeric(d["Prezzo acquisto"], errors="coerce").fillna(0.0)
    d["CostKey"] = d.apply(lambda r: _cost_key(r.get("SKU", ""), r.get("Nome Prodotto", "")), axis=1)
    d["Prezzo acquisto"] = d.apply(lambda r: float(costs.get(r["CostKey"], r["Prezzo acquisto"] if pd.notna(r["Prezzo acquisto"]) else 0.0)), axis=1)
    d["Margine unitario"] = d["Prezzo vendita ex IVA"] - d["Prezzo acquisto"]
    d["Importo ex IVA"] = d["Pezzi"] * d["Prezzo vendita ex IVA"]
    d["Margine"] = d["Pezzi"] * d["Margine unitario"]
    d["Classe"] = d["Classe"].map(clean_str)
    d.loc[d["Classe"].eq(""), "Classe"] = "Docenti / ATA"

    total_orders = int(d["N. Ordine"].map(clean_str).replace("", pd.NA).dropna().nunique())
    total_pieces = int(d["Pezzi"].sum())
    total_amount = float(d["Importo ex IVA"].sum())
    total_margin = float(d["Margine"].sum())

    by_class = d.groupby("Classe", dropna=False, as_index=False).agg(Pezzi=("Pezzi", "sum"), **{"Importo ex IVA": ("Importo ex IVA", "sum"), "Margine": ("Margine", "sum")}).sort_values(["Classe"], kind="stable").reset_index(drop=True)
    if not by_class.empty and (by_class["Classe"] == "Docenti / ATA").any():
        first = by_class[by_class["Classe"] == "Docenti / ATA"]
        rest = by_class[by_class["Classe"] != "Docenti / ATA"]
        by_class = pd.concat([first, rest], ignore_index=True)

    by_product = d.groupby(["SKU", "Nome Prodotto"], dropna=False, as_index=False).agg(Pezzi=("Pezzi", "sum"), **{"Prezzo vendita ex IVA": ("Prezzo vendita ex IVA", "mean"), "Prezzo acquisto": ("Prezzo acquisto", "mean"), "Importo ex IVA": ("Importo ex IVA", "sum"), "Margine": ("Margine", "sum")}).sort_values(["Pezzi", "Nome Prodotto"], ascending=[False, True], kind="stable").reset_index(drop=True)
    by_color = d.groupby(["SKU", "Colore"], dropna=False, as_index=False).agg(Pezzi=("Pezzi", "sum"), **{"Prezzo vendita ex IVA": ("Prezzo vendita ex IVA", "mean"), "Prezzo acquisto": ("Prezzo acquisto", "mean"), "Importo ex IVA": ("Importo ex IVA", "sum"), "Margine": ("Margine", "sum")}).sort_values(["Pezzi", "Colore"], ascending=[False, True], kind="stable").reset_index(drop=True)
    return d, total_orders, total_pieces, total_amount, total_margin, by_class, by_product, by_color

def _eur(x: float) -> str: return f"€ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def page_finanze(df_norm: pd.DataFrame, proj_dir: Path) -> None:
    st.subheader("Finanze")
    st.caption("Riepilogo economico della tornata/scuola. Tutti i valori qui sono ex IVA 22% (Prezzo unitario / 1,22).")
    costs = load_costs(proj_dir)
    _, total_orders, total_pieces, total_amount, total_margin, by_class, by_product, by_color = finance_summary(df_norm, costs)

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Ordini", total_orders)
    with c2: st.metric("Pezzi totali", total_pieces)
    with c3: st.metric("Importo ex IVA", _eur(total_amount))
    with c4: st.metric("Margine", _eur(total_margin))

    st.markdown("### Prezzi acquisto")
    cost_view = by_product[["SKU", "Nome Prodotto", "Prezzo acquisto"]].copy()
    cost_view["Chiave"] = cost_view.apply(lambda r: _cost_key(r["SKU"], r["Nome Prodotto"]), axis=1)
    edited = st.data_editor(cost_view, hide_index=True, use_container_width=True, disabled=["SKU", "Nome Prodotto", "Chiave"], column_config={"Prezzo acquisto": st.column_config.NumberColumn("Prezzo acquisto", step=0.01, format="%.2f"), "Chiave": None}, key="finance_cost_editor")
    
    if st.button("💾 Salva prezzi acquisto"):
        new_costs = load_costs(proj_dir)
        for _, r in edited.iterrows():
            try: new_costs[str(r["Chiave"])] = float(r["Prezzo acquisto"])
            except Exception: pass
        save_costs(proj_dir, new_costs)
        st.success("Prezzi acquisto salvati.")
        st.rerun()

    costs = load_costs(proj_dir)
    _, total_orders, total_pieces, total_amount, total_margin, by_class, by_product, by_color = finance_summary(df_norm, costs)
    t1, t2, t3 = st.tabs(["Per classe", "Per prodotto", "Per colore"])
    with t1:
        show = by_class.copy()
        for c in ["Importo ex IVA", "Margine"]: show[c] = show[c].map(_eur)
        st.dataframe(show, use_container_width=True, hide_index=True)
    with t2:
        show = by_product.copy()
        for c in ["Prezzo vendita ex IVA", "Prezzo acquisto", "Importo ex IVA", "Margine"]: show[c] = show[c].map(_eur)
        st.dataframe(show, use_container_width=True, hide_index=True)
    with t3:
        show = by_color.copy()
        for c in ["Prezzo vendita ex IVA", "Prezzo acquisto", "Importo ex IVA", "Margine"]: show[c] = show[c].map(_eur)
        st.dataframe(show, use_container_width=True, hide_index=True)

@st.dialog("Gestione SKU Sostitutivi")
def substitute_modal(sku_color_options, proj_dir: Path):
    if "sub_idx" not in st.session_state: st.session_state.sub_idx = 0
    if st.session_state.sub_idx >= len(sku_color_options): st.session_state.sub_idx = 0
        
    sel_val = st.selectbox("Seleziona rapidamente (oppure usa le frecce):", options=sku_color_options, index=st.session_state.sub_idx)
    if sel_val != sku_color_options[st.session_state.sub_idx]:
        st.session_state.sub_idx = sku_color_options.index(sel_val)
        st.rerun()

    current_sel = sku_color_options[st.session_state.sub_idx]
    sku, color = [x.strip() for x in current_sel.split(" — ", 1)]
    k = f"{clean_str(sku)}||{clean_str(color)}"
    
    subs = load_subs(proj_dir)
    sub_data = subs.get(k, {})
    if isinstance(sub_data, str): sub_data = {"fornitore": "Altro", "sku": sub_data}
        
    # Memoria intelligente
    last_forn = st.session_state.get("last_sub_forn", "ActionWear")
    last_sku = st.session_state.get("last_sub_sku", "")
    
    curr_forn = sub_data.get("fornitore", last_forn) if sub_data else last_forn
    curr_sku = sub_data.get("sku", last_sku) if sub_data else last_sku

    custom_sups = load_custom_suppliers()
    base_sups = ["ActionWear", "Basic", "Roly", "Top-Tex", "Stanley/Stella", "Innova"]
    all_sups = base_sups + custom_sups + ["Altro"]
    idx_forn = all_sups.index(curr_forn) if curr_forn in all_sups else all_sups.index("Altro")
    
    with st.form(f"sub_form_{st.session_state.sub_idx}", clear_on_submit=False):
        st.markdown(f"Stai modificando: **{sku}** — Variante **{color}**")
        f1, f2 = st.columns(2)
        with f1:
            sel_forn = st.selectbox("Fornitore", options=all_sups, index=idx_forn)
            altro_forn = st.text_input("Specifica nuovo fornitore:", value=curr_forn if curr_forn not in base_sups+custom_sups else "") if sel_forn == "Altro" else ""
        with f2: 
            new_sku = st.text_input("SKU Sostitutivo", value=curr_sku, placeholder="Es. DOG123")
            
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
        b1, b2, b3 = st.columns(3)
        with b1: btn_prev = st.form_submit_button("⬅️ Salva e Precedente", use_container_width=True)
        with b2: btn_save = st.form_submit_button("💾 Salva (Invio)", type="primary", use_container_width=True)
        with b3: btn_next = st.form_submit_button("Salva e Successivo ➡️", use_container_width=True)
        
        if btn_prev or btn_save or btn_next:
            final_forn = altro_forn.strip() if sel_forn == "Altro" else sel_forn
            final_sku = new_sku.strip()
            
            # Salvataggio in cache per la compilazione automatica al prossimo articolo
            st.session_state["last_sub_forn"] = final_forn
            st.session_state["last_sub_sku"] = final_sku
            
            if final_forn and final_forn not in base_sups and final_forn not in custom_sups and final_forn != "Altro":
                custom_sups.append(final_forn)
                save_custom_suppliers(custom_sups)
            
            if final_sku: subs[k] = {"fornitore": final_forn, "sku": final_sku}
            else: subs.pop(k, None)
            
            save_subs(proj_dir, subs)
            
            if btn_prev: st.session_state.sub_idx = max(0, st.session_state.sub_idx - 1)
            if btn_next: st.session_state.sub_idx = min(len(sku_color_options)-1, st.session_state.sub_idx + 1)
            
            st.rerun()
            
@st.cache_data
def get_cached_dataframe(file_path_str: str, file_mtime: float) -> pd.DataFrame:
    """Carica e normalizza l'Excel solo se il file è stato modificato, velocizzando l'app del 1000%"""
    df_raw = pd.read_excel(file_path_str)
    return df_normalize(df_raw)
    
def main() -> None:
    st.set_page_config(page_title="WUPI Suite", layout="wide", page_icon=str(FAVICON_PATH) if FAVICON_PATH.exists() else "🧰")
    
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    
    # ---------------------
    # WORKSPACE SIDEBAR
    # ---------------------
    st.sidebar.title("WUPI Workspace")
    schools = [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()]
    sel_school = st.sidebar.selectbox("1️⃣ Seleziona Scuola", options=["-- Seleziona --"] + sorted(schools))
    
    if sel_school and sel_school != "-- Seleziona --":
        with st.sidebar.expander("⚙️ Gestisci Scuola"):
            new_school_name = st.text_input("Rinomina", value=sel_school, key="ren_sch")
            if st.button("Salva Nuovo Nome") and new_school_name != sel_school:
                os.rename(PROJECTS_DIR / sel_school, PROJECTS_DIR / safe_dir_name(new_school_name))
                st.rerun()
            if st.button("❌ Elimina Scuola in modo permanente"):
                shutil.rmtree(PROJECTS_DIR / sel_school)
                st.rerun()
                
    with st.sidebar.expander("➕ Nuova Scuola"):
        new_school = st.text_input("Nome Scuola", placeholder="Es. Liceo Respighi")
        if st.button("Crea Scuola", use_container_width=True) and new_school:
            (PROJECTS_DIR / safe_dir_name(new_school)).mkdir(parents=True, exist_ok=True)
            st.rerun()
            
    sel_proj = None
    if sel_school and sel_school != "-- Seleziona --":
        school_dir = PROJECTS_DIR / sel_school
        projs = [d.name for d in school_dir.iterdir() if d.is_dir()]
        sel_proj = st.sidebar.selectbox("2️⃣ Seleziona Tornata / Progetto", options=["-- Seleziona --"] + sorted(projs))
        
        if sel_proj and sel_proj != "-- Seleziona --":
            with st.sidebar.expander("⚙️ Gestisci Tornata"):
                new_proj_name = st.text_input("Rinomina", value=sel_proj, key="ren_prj")
                if st.button("Salva Nuovo Nome Tornata") and new_proj_name != sel_proj:
                    os.rename(school_dir / sel_proj, school_dir / safe_dir_name(new_proj_name))
                    st.rerun()
                if st.button("❌ Elimina Tornata in modo permanente"):
                    shutil.rmtree(school_dir / sel_proj)
                    st.rerun()
        
        with st.sidebar.expander("➕ Nuova Tornata"):
            new_project = st.text_input("Nome Tornata", placeholder="Es. Febbraio 2026")
            if st.button("Crea Tornata", use_container_width=True) and new_project:
                (school_dir / safe_dir_name(new_project)).mkdir(parents=True, exist_ok=True)
                st.rerun()

    top_l, top_r = st.columns([7, 1])
    with top_l:
        st.title("WUPI Suite")
        st.caption(f"Build: STUDIO_v13_CLEAN_PDF (stable)")
    with top_r:
        if LOGO_PATH.exists(): st.image(str(LOGO_PATH), use_container_width=True)

    if not sel_school or sel_school == "-- Seleziona --" or not sel_proj or sel_proj == "-- Seleziona --":
        st.info("👈 **Apri il menu a sinistra** per selezionare (o creare) una Scuola e una Tornata per iniziare a lavorare.")
        return

    proj_dir = PROJECTS_DIR / sel_school / sel_proj
    excel_path = proj_dir / "data.xlsx"

    if not excel_path.exists():
        st.warning(f"Nessun file Excel caricato per il progetto **{sel_school} ➔ {sel_proj}**.")
        uploaded = st.file_uploader("Carica Excel Ordini (.xlsx)", type=["xlsx"])
        if uploaded:
            excel_path.write_bytes(uploaded.getvalue())
            st.success("File caricato con successo!")
            st.rerun()
        return

    c_head, c_del = st.columns([8, 2])
    with c_head:
        st.success(f"📂 Workspace Attivo: **{sel_school}** ➔ **{sel_proj}**")
    with c_del:
        if st.button("🗑 Sostituisci Excel", help="Elimina il file Excel caricato per poterne caricare un altro."):
            excel_path.unlink()
            st.rerun()

   # Caricamento fulmineo con Cache
    mtime = os.path.getmtime(excel_path)
    df = get_cached_dataframe(str(excel_path), mtime)

    current_proj_id = str(proj_dir)
    if "confirmed" not in st.session_state or st.session_state.get("confirmed_proj") != current_proj_id:
        st.session_state["confirmed"] = set(load_state(proj_dir))
        st.session_state["confirmed_proj"] = current_proj_id
    confirmed = st.session_state["confirmed"]

    subs = load_subs(proj_dir)
    file_stock = load_stock(proj_dir)

    tabs = st.tabs(["📦 Report acquisto", "🏷 Etichette", "💸 Ordini da pagare", "📖 Bibbia maker", "💰 Finanze"])
    
    with tabs[0]:
        st.subheader("Pivot ordine fornitore")
        st.caption("0 nascosti, Totale fisso a destra (bold). Le righe confermate diventano VERDI.")
        piv_full = pivot_report(df)

        c_search, c_sub, c_pdf = st.columns([4, 1.5, 1.5])
        
        if "show_sub_modal" not in st.session_state: 
            st.session_state.show_sub_modal = False
            
        with c_search:
            q = st.text_input("🔍 Cerca (SKU / Prodotto / Colore)", key="pivot_search")
            
        with c_sub:
            st.markdown("<br>", unsafe_allow_html=True)
            # Questo attiva la variabile di memoria, così la modale non si chiude ai ricaricamenti!
            if st.button("🔄 SKU Sostitutivo", use_container_width=True):
                st.session_state.show_sub_modal = True

        if st.session_state.show_sub_modal:
            pairs_sub = piv_full[["SKU", "Colore"]].drop_duplicates().sort_values(["SKU", "Colore"], kind="stable")
            sku_color_options = [f'{r["SKU"]} — {r["Colore"]}' for _, r in pairs_sub.iterrows()]
            substitute_modal(sku_color_options, proj_dir)

        with c_pdf:
            st.markdown("<br>", unsafe_allow_html=True)
            pdf_bytes = make_order_summary_pdf(piv_full, subs, file_stock, sel_school, sel_proj)
            st.download_button("📄 PDF Ordine", data=pdf_bytes, file_name=f"ordine_{sel_school}_{sel_proj}.pdf", mime="application/pdf", use_container_width=True)
        piv_view = piv_full
        if q:
            qq = str(q).strip()
            piv_view = piv_full[piv_full["SKU"].astype(str).str.contains(qq, case=False, na=False) | piv_full["Nome Prodotto"].astype(str).str.contains(qq, case=False, na=False) | piv_full["Colore"].astype(str).str.contains(qq, case=False, na=False)].copy()

        render_pivot_html(piv_view, confirmed, subs, file_stock)
        st.markdown('<div class="wupi-gap-after-pivot"></div>', unsafe_allow_html=True)

        pairs = piv_full[["SKU", "Nome Prodotto"]].drop_duplicates().sort_values(["SKU", "Nome Prodotto"], kind="stable")
        options = [f'{r["SKU"]} — {r["Nome Prodotto"]}' for _, r in pairs.iterrows()]
        
        if "pair_idx" not in st.session_state: st.session_state["pair_idx"] = 0
        if "pair_widget_ver" not in st.session_state: st.session_state["pair_widget_ver"] = 0
        if st.session_state.get("advance_next_sku", False) and len(options) > 0:
            st.session_state["advance_next_sku"] = False
            if st.session_state["pair_idx"] < len(options) - 1:
                st.session_state["pair_idx"] += 1
                st.session_state["pair_widget_ver"] += 1

        a, b, c = st.columns([1, 8, 1], gap='small')
        with a:
            if st.button("‹", key="prev_pair", use_container_width=True):
                st.session_state["pair_idx"] = max(0, st.session_state["pair_idx"] - 1)
                st.session_state["pair_widget_ver"] += 1; st.rerun()
        with b:
            sel = st.selectbox("Seleziona SKU", options=options, index=st.session_state["pair_idx"] if len(options) else 0, key=f"pair_select_{st.session_state['pair_widget_ver']}", label_visibility="collapsed")
        with c:
            if st.button("›", key="next_pair", use_container_width=True):
                st.session_state["pair_idx"] = min(len(options) - 1, st.session_state["pair_idx"] + 1)
                st.session_state["pair_widget_ver"] += 1; st.rerun()

        if len(options) > 0 and sel in options: st.session_state["pair_idx"] = options.index(sel)

        sku = sel.split(" — ", 1)[0].strip()
        prod = sel.split(" — ", 1)[1].strip()
        render_color_cards(df, sku, prod, confirmed, proj_dir)

    with tabs[1]:
        st.subheader("Etichette (152×102 mm orizzontale)")
        st.caption("Se Nome/Cognome studente vuoti → Docenti / ATA come classe e nominativo da Docente/ATA.")
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write("Impostazioni rapide")
            w = st.number_input("Larghezza (mm)", value=152.0, step=1.0)
            h = st.number_input("Altezza (mm)", value=102.0, step=1.0)
            with st.expander("Opzioni avanzate (margini e font)", expanded=False):
                m = st.number_input("Margine (mm)", value=8.0, step=0.5)
                logo_w = st.number_input("Larghezza logo (mm)", value=28.0, step=1.0)
                title_pt = st.number_input("Font classe (pt)", value=18.0, step=0.5)
                header_pt = st.number_input("Font intestazioni (pt)", value=9.0, step=0.5)
                row_pt = st.number_input("Font righe (pt)", value=9.0, step=0.5)
                row_h = st.number_input("Altezza riga (mm)", value=4.2, step=0.2)
                strip_modello = st.checkbox('Elimina "Modello" (es: | Modello Joker → | Joker)', value=False)
        with col2: logo_up = st.file_uploader("Logo (PNG) opzionale", type=["png"])
        logo_bytes = logo_up.getvalue() if logo_up else (LOGO_PATH.read_bytes() if LOGO_PATH.exists() else None)

        cfg = LabelCfg(w_mm=float(w), h_mm=float(h), margin_mm=float(m), logo_w_mm=float(logo_w), title_pt=float(title_pt), header_pt=float(header_pt), row_pt=float(row_pt), row_h_mm=float(row_h), strip_modello=bool(strip_modello))
        if st.button("Genera PDF etichette", type="primary"):
            pdf = make_labels_pdf(df, logo_bytes, cfg)
            st.download_button("⬇️ Scarica PDF", data=pdf, file_name=f"etichette_{sel_school}.pdf", mime="application/pdf")

    with tabs[2]: page_pending(df)
    with tabs[3]: page_bibbia(df)
    with tabs[4]: page_finanze(df, proj_dir)

if __name__ == "__main__":
    global_css()
    main()
