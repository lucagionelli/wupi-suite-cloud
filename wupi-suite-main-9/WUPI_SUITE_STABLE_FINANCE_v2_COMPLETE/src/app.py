from __future__ import annotations

import hashlib
import io
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
import os
import re

from PIL import Image
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# -------------------------
# Config
# -------------------------
SIZE_ORDER = ["UNICA", "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]
SIZE_ALIAS = {"XXL": "2XL", "2XL": "2XL"}

APP_SUPPORT = Path.home() / "Library" / "Application Support" / "WUPI Suite"
STATE_PATH = APP_SUPPORT / "state_confirm.json"
COSTS_PATH = APP_SUPPORT / "costs.json"
BIBBIA_MANUAL_PATH = APP_SUPPORT / "bibbia_manual.json"

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
if not ASSETS_DIR.exists():
    ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
LOGO_PATH = ASSETS_DIR / "wupi.png"
FAVICON_PATH = ASSETS_DIR / "favicon.png"

GREEN = "#d7f7d7"

# Normalizzazione colori per match mockup
COLOR_ALIAS_MAP = {
  "arancione": "arancione",
  "bianco": "bianco",
  "bk": "nero",
  "black": "nero",
  "blu": "navy",
  "blu_navy": "navy",
  "blu_royal": "royal",
  "blunavy": "navy",
  "bluroyal": "royal",
  "bu": "burgundy",
  "burgundy": "burgundy",
  "ca": "cardinal",
  "caramel": "dark_caramel",
  "cardinal": "cardinal",
  "cardinal_red": "cardinal",
  "cardinalred": "cardinal",
  "ch": "chocolate",
  "chocolate": "chocolate",
  "dark_caramel": "dark_caramel",
  "dark_grey": "dark_grey",
  "darkcaramel": "dark_caramel",
  "darkgrey": "dark_grey",
  "dc": "dark_caramel",
  "dg": "dark_grey",
  "dp": "dusty_pink",
  "dusty_pink": "dusty_pink",
  "dustypink": "dusty_pink",
  "du": "dusty_green",
  "dusty_green": "dusty_green",
  "dustygreen": "dusty_green",
  "dusty_green_": "dusty_green",
  "earth_green": "military",
  "earthgreen": "military",
  "eg": "military",
  "fg": "forest",
  "forest": "forest",
  "forest_green": "forest",
  "forestgreen": "forest",
  "giallo": "gold",
  "go": "gold",
  "gold": "gold",
  "grey_heater": "grey_heather",
  "grey_heather": "grey_heather",
  "greyheater": "grey_heather",
  "greyheather": "grey_heather",
  "grigio_melange": "grey_heather",
  "grigiomelange": "grey_heather",
  "gy": "grey_heather",
  "ib": "ink_blue",
  "ig": "irish_green",
  "ink_blue": "ink_blue",
  "inkblue": "ink_blue",
  "irish_green": "irish_green",
  "irishgreen": "irish_green",
  "jade": "salvia",
  "jade_green": "salvia",
  "jadegreen": "salvia",
  "ma": "mastic",
  "marrone": "chocolate",
  "mastic": "mastic",
  "mb": "mineral_blue",
  "military": "military",
  "mineral_blue": "mineral_blue",
  "mineralblue": "mineral_blue",
  "mocha": "chocolate",
  "moka": "chocolate",
  "mu": "mustard",
  "mustard": "mustard",
  "navy": "navy",
  "navy_blue": "navy",
  "navyblue": "navy",
  "nero": "nero",
  "ny": "navy",
  "off_white": "off_white",
  "offwhite": "off_white",
  "ol": "olive",
  "olive": "olive",
  "or": "arancione",
  "orange": "arancione",
  "ow": "off_white",
  "pe": "petroleum",
  "peacock": "peacock_ink_blue",
  "petroleum": "petroleum",
  "pink": "rosa",
  "pu": "purple",
  "purple": "purple",
  "rb": "royal",
  "rd": "red",
  "red": "red",
  "ro": "rosa",
  "rosa": "rosa",
  "rosso": "red",
  "rosso_cardinal": "cardinal",
  "rossocardinal": "cardinal",
  "royal": "royal",
  "royal_blue": "royal",
  "royalblue": "royal",
  "ru": "rust",
  "rust": "rust",
  "sa": "salvia",
  "salvia": "salvia",
  "sand": "mastic",
  "sky": "mineral_blue",
  "urban_slate": "urban_slate",
  "urban_slathe": "urban_slate",
  "urbanslate": "urban_slate",
  "urbanslathe": "urban_slate",
  "us": "urban_slate",
  "viola": "purple",
  "wh": "bianco",
  "white": "bianco",
  "pk": "peacock_ink_blue",
  "peacock_ink_blue": "peacock_ink_blue",
  "peacockinkblue": "peacock_ink_blue",
  "peacock_inkblue": "peacock_ink_blue",
  "peacockink_blue": "peacock_ink_blue",
  "lg": "light_grey",
  "lightgrey": "light_grey",
  "light_grey": "light_grey",
  "light_gray": "light_grey",
  "ac": "sand_almond_cream",
  "almond_cream": "sand_almond_cream",
  "sand_almond_cream": "sand_almond_cream",
  "sandalmondcream": "sand_almond_cream",
  "almond_sand_cream": "sand_almond_cream",
  "almondsandcream": "sand_almond_cream",
  "almond_sandcream": "sand_almond_cream",
  "almondsand_cream": "sand_almond_cream",
  "sand_cream_almond": "sand_almond_cream",
  "cream_almond_sand": "sand_almond_cream",
  "earthygreen": "military",
  "earthy_green": "military",
  "earthy green": "military",
  "gh": "grey_heather"
}

def color_to_canon_key(s: str) -> str:
    t = _norm_key(s)
    if not t:
        return ""
    return COLOR_ALIAS_MAP.get(t, t)

# -------------------------
# Helpers
# -------------------------
def file_sig(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()[:12]

def clean_str(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s

def normalize_size(s: str) -> str:
    s2 = clean_str(s).upper()
    if not s2:
        return ""
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
    }
    cols = list(df.columns)
    norm_map = { _norm_colname(c): c for c in cols }
    rename = {}
    for std, keys in alias.items():
        if std in cols:
            continue
        for k in keys:
            k2 = _norm_colname(k)
            if k2 in norm_map:
                rename[norm_map[k2]] = std
                break
    if rename:
        df = df.rename(columns=rename)
    return df

def ensure_cols(df: pd.DataFrame, cols: List[str]) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti: {', '.join(missing)}")

def load_state() -> Dict:
    try:
        return json_loads(STATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_state(data: Dict) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json_dumps(data), encoding="utf-8")

def json_dumps(d: Dict) -> str:
    import json
    return json.dumps(d, ensure_ascii=False, indent=2)

def json_loads(s: str) -> Dict:
    import json
    return json.loads(s)

def load_costs() -> Dict:
    try:
        return json_loads(COSTS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_costs(data: Dict) -> None:
    COSTS_PATH.parent.mkdir(parents=True, exist_ok=True)
    COSTS_PATH.write_text(json_dumps(data), encoding="utf-8")

def _cost_key(sku: str, product: str) -> str:
    return f"{clean_str(sku)}||{clean_str(product)}"

def load_manual_mockups() -> Dict:
    try:
        raw = json_loads(BIBBIA_MANUAL_PATH.read_text(encoding="utf-8"))
    except Exception:
        raw = {}
    out = {}
    for k, v in raw.items():
        try:
            if isinstance(v, str) and Path(v).exists():
                out[tuple(k.split("|||"))] = Path(v).read_bytes()
        except Exception:
            pass
    return out

def save_manual_mockups(rawmap: Dict[str, str]) -> None:
    BIBBIA_MANUAL_PATH.parent.mkdir(parents=True, exist_ok=True)
    BIBBIA_MANUAL_PATH.write_text(json_dumps(rawmap), encoding="utf-8")

def _parse_taglie_items(taglie_str: str) -> list[tuple[str, int]]:
    items = []
    for part in clean_str(taglie_str).split():
        if ":" in part:
            t, q = part.split(":", 1)
            try:
                items.append((t, int(q)))
            except Exception:
                pass
    items.sort(key=lambda x: sort_size_key(x[0]))
    return items

def key_row(sku: str, prod: str, color: str) -> str:
    return f"{sku}||{prod}||{color}"

def normalize_key(k: str) -> str:
    try:
        parts = (k or "").split("||")
        if len(parts) != 3:
            return clean_str(k)
        return key_row(clean_str(parts[0]), clean_str(parts[1]), clean_str(parts[2]))
    except Exception:
        return clean_str(k)

def canon_key(sku: str, prod: str, color: str) -> str:
    return f"{clean_str(sku).lower()}||{clean_str(prod).lower()}||{clean_str(color).lower()}"

def sort_size_key(taglia: str) -> int:
    t = (taglia or "").upper().strip()
    if t == "XXL":
        t = "2XL"
    if not t:
        return 999
    return SIZE_ORDER.index(t) if t in SIZE_ORDER else 998

def df_normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = standardize_required_columns(df)
    ensure_cols(df, ["SKU", "Nome Prodotto", "Colore", "Taglia", "Pezzi"])
    out = df.copy()

    out["SKU"] = out["SKU"].map(clean_str)
    out["Nome Prodotto"] = out["Nome Prodotto"].map(clean_str)
    out["Colore"] = out["Colore"].map(clean_str)  
    out["Taglia"] = out["Taglia"].map(normalize_size)
    out["Pezzi"] = pd.to_numeric(out["Pezzi"], errors="coerce").fillna(0).astype(int)
    if "Prezzo unitario" in out.columns:
        out["Prezzo unitario"] = pd.to_numeric(out["Prezzo unitario"], errors="coerce").fillna(0.0)
    else:
        out["Prezzo unitario"] = 0.0
    if "Prezzo acquisto" in out.columns:
        out["Prezzo acquisto"] = pd.to_numeric(out["Prezzo acquisto"], errors="coerce").fillna(0.0)
    else:
        out["Prezzo acquisto"] = 0.0

    for c in ["Nome Studente", "Cognome Studente", "Classe", "Docente/ATA", "N. Ordine", "Nome incisione"]:
        if c in out.columns:
            out[c] = out[c].map(clean_str)
        else:
            out[c] = "" if c != "N. Ordine" else ""

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

def render_pivot_html(piv: pd.DataFrame, confirmed: set[str]) -> None:
    view = piv.copy()
    for c in [s for s in SIZE_ORDER if s in view.columns]:
        view[c] = view[c].replace({0: ""})
    view["Totale"] = piv["Totale"].astype(int)

    cols = list(view.columns)

    css = f"""<style>
/* Tabella piatta e solida */
.table-wrap {{ 
    overflow:auto; 
    background-color: #ffffff;
    border: 1px solid #e5e5ea; 
    border-radius: 12px;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02);
}}
table.wupi {{ border-collapse:separate; border-spacing:0; width:100%; font-size:14px; }}
table.wupi th, table.wupi td {{ padding:12px; border-bottom:1px solid #f0f0f2; vertical-align:middle; color: #1d1d1f; }}
table.wupi th {{ 
    position:sticky; top:0; 
    background-color: #fafafc; 
    z-index:2; font-weight:600; letter-spacing: -0.2px;
}}
table.wupi td {{ background-color: #ffffff; }}
table.wupi td.tot, table.wupi th.tot {{ position:sticky; right:0; z-index:3; font-weight:700; }}
table.wupi th.tot {{ background-color: #fafafc; border-left: 1px solid #f0f0f2; }}
table.wupi td.tot {{ background-color: #fafafc; border-left: 1px solid #f0f0f2; }}
tr.confirmed td {{ background-color: #f5f5f7; }}
tr.confirmed td.tot {{ background-color: #eaeaef; }}
.center {{ text-align:center; }}
</style>"""

    html: list[str] = []
    html.append(css)
    html.append('<div class="table-wrap"><table class="wupi">')
    html.append('<thead><tr>')

    for c in cols:
        cls = "tot" if c == "Totale" else ""
        align = "center" if c in SIZE_ORDER + ["Totale"] else ""
        html.append(f'<th class="{cls} {align}">{c}</th>')

    html.append('</tr></thead><tbody>')

    for _, r in view.iterrows():
        k = normalize_key(key_row(clean_str(r.get("SKU", "")), clean_str(r.get("Nome Prodotto", "")), clean_str(r.get("Colore", ""))))
        tr_cls = "confirmed" if k in confirmed else ""
        html.append(f'<tr class="{tr_cls}">')
        for c in cols:
            cls = "tot" if c == "Totale" else ""
            align = "center" if c in SIZE_ORDER + ["Totale"] else ""
            val = r[c]
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

/* Cards solide e pulite */
.wupi-card {
  background-color: #ffffff;
  border: 1px solid #e5e5ea;
  border-radius: 16px;
  padding: 16px;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02);
  transition: transform 0.1s ease, box-shadow 0.1s ease;
}
.wupi-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 16px rgba(0, 0, 0, 0.06);
}
/* Card confermata grigio solido */
.wupi-card.confirmed {
  background-color: #f5f5f7;
  border: 2px solid #1d1d1f;
}
.card-head { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:12px; }
.color-name { font-weight:700; font-size:17px; letter-spacing:-0.3px; color: #1d1d1f; }
.color-tot { font-weight:600; font-size:15px; color: #86868b; }
.chips { display:flex; flex-wrap:wrap; gap:8px; }

/* Taglie (pills) solide */
.chip {
  display:inline-flex; gap:6px; align-items:center;
  padding: 4px 10px; border-radius: 8px;
  background-color: #f0f0f2;
  font-size: 13px; font-weight: 500; color: #555;
}
.chip .q { font-weight:700; font-size:14px; color: #1d1d1f; }
.btn-row { display:flex; gap:12px; justify-content:center; margin-top:18px; }
</style>
""", unsafe_allow_html=True)

def render_color_cards(df: pd.DataFrame, sku: str, prod: str, confirmed: set[str], state_sig: str, state: Dict) -> None:
    cards_css()
    sub = df[(df["SKU"] == sku) & (df["Nome Prodotto"] == prod)].copy()
    if sub.empty:
        st.info("Nessun dato per questo SKU/prodotto.")
        return

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
    for i, (color, tot, items) in enumerate(blocks):
        col = cols[i % 3]
        k = normalize_key(key_row(clean_str(sku), clean_str(prod), clean_str(color)))
        is_done = k in confirmed
        chips = "".join([f'<span class="chip">{t}<span class="q">{q}</span></span>' for t, q in items])

        with col:
            st.markdown(f"""
<div class="wupi-card {'confirmed' if is_done else ''}">
  <div class="card-head">
    <div class="color-name">{color if color else '(colore vuoto)'}</div>
    <div class="color-tot">{tot} pz</div>
  </div>
  <div class="chips">{chips}</div>
</div>
""", unsafe_allow_html=True)
            st.markdown('<div style="height:16px"></div>', unsafe_allow_html=True)
            b1, b2 = st.columns(2)
            with b1:
                if st.button("✓ Conferma", key=f"conf__{hashlib.md5(k.encode()).hexdigest()}", disabled=is_done, use_container_width=True):
                    confirmed.add(normalize_key(k))
                    state[state_sig] = sorted(list(normalize_key(k) for k in confirmed))
                    save_state(state)
                    st.session_state["confirmed"] = set(confirmed)
                    st.rerun()
            with b2:
                if st.button("↩︎ Annulla", key=f"undo__{hashlib.md5(k.encode()).hexdigest()}", disabled=not is_done, use_container_width=True):
                    confirmed.discard(normalize_key(k))
                    state[state_sig] = sorted(list(normalize_key(k) for k in confirmed))
                    save_state(state)
                    st.session_state["confirmed"] = set(confirmed)
                    st.rerun()

    _, r1, r2 = st.columns([6, 2, 2])
    with r1:
        if st.button("✓ Conferma tutto lo SKU", key=f"all_{sku}_{prod}", use_container_width=True):
            for color, _, _ in blocks:
                confirmed.add(normalize_key(key_row(sku, prod, color)))
            state[state_sig] = sorted(list(normalize_key(k) for k in confirmed))
            save_state(state)
            st.session_state['confirmed'] = set(confirmed)
            st.session_state['advance_next_sku'] = True
            st.rerun()
    with r2:
        if st.button("↩︎ Annulla tutto lo SKU", key=f"unall_{sku}_{prod}", use_container_width=True):
            for color, _, _ in blocks:
                confirmed.discard(normalize_key(key_row(sku, prod, color)))
            state[state_sig] = sorted(list(confirmed))
            save_state(state)
            st.session_state["confirmed"] = set(confirmed)
            st.rerun()

def global_css() -> None:
    st.markdown("""
<style>
/* Sfondo principale pulito e solido */
.stApp {
    background-color: #ffffff;
    color: #1d1d1f;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}

/* Stile bottoni pulito (Nero/Grigio/Bianco) */
.stButton > button {
    background-color: #ffffff !important;
    color: #1d1d1f !important;
    border: 1px solid #d2d2d7 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    border-color: #86868b !important;
    background-color: #f5f5f7 !important;
}
/* Bottone Primary */
.stButton > button[kind="primary"] {
    background-color: #1d1d1f !important;
    color: #ffffff !important;
    border: 1px solid #1d1d1f !important;
}
.stButton > button[kind="primary"]:hover {
    background-color: #333336 !important;
}
*:focus { outline:none !important; }
button:focus { box-shadow: 0 0 0 2px rgba(0,0,0,.1) !important; }
a, a:visited { color:#1d1d1f; }

/* Box di caricamento File piatto */
[data-testid="stFileUploader"] {
    background-color: #fafafc !important;
    border: 1px dashed #d2d2d7 !important;
    border-radius: 12px !important;
    padding: 1.5rem !important;
}

/* Metriche e Dataframe piatti */
[data-testid="stMetric"], [data-testid="stDataFrame"] > div {
    background-color: #ffffff !important;
    border-radius: 12px !important;
    border: 1px solid #e5e5ea !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.02) !important;
}
[data-testid="stMetric"] { padding: 16px; }

/* Taglie (pills) globali */
.chips { display:flex; flex-wrap:wrap; gap:6px; }
.chip {
  display:inline-flex; gap:6px; align-items:center;
  padding: 4px 10px; border-radius: 8px;
  background-color: #f0f0f2;
  font-size: 13px; font-weight: 500; color: #555;
}
.chip .q { font-weight:700; font-size:14px; color: #1d1d1f; }

.wupi-gap-after-pivot { height: 14px; }
</style>
""", unsafe_allow_html=True)

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
        try:
            logo_img = ImageReader(io.BytesIO(logo_bytes))
        except Exception:
            logo_img = None

    for col in ["N. Ordine", "Nome Prodotto", "Colore", "Taglia", "Studente", "GruppoEtichetta", "Pezzi"]:
        if col not in df.columns:
            df[col] = ""

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
        while s and stringWidth(s + "…", font, size) > max_w:
            s = s[:-1]
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
                raw_art = str(r.get("Nome Prodotto", "") or "")
                raw_art = raw_art.strip()
                if cfg.strip_modello:
                    raw_art = re.sub(r"(?i)\bmodello\b\s*", "", raw_art)
                if len(raw_art) > 20:
                    raw_art = raw_art[:20] + "…"
                articolo = fit(raw_art, col_art_w, "Helvetica", row_pt)
                colore = fit(str(r.get("Colore", "") or ""), col_col_w, "Helvetica", row_pt)
                taglia = fit(str(r.get("Taglia", "") or ""), col_size_w, "Helvetica", row_pt)
                persona_raw = str(r.get("Studente", "") or "").strip()
                name_size = row_pt
                while name_size > 6.5 and stringWidth(persona_raw, "Helvetica", name_size) > col_name_w:
                    name_size -= 0.2

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
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s: return 0.0
    import re as _re
    s = _re.sub(r"[^\d\.,\-]", "", s)

    has_dot = "." in s
    has_comma = "," in s

    if has_dot and has_comma:
        last_dot = s.rfind(".")
        last_comma = s.rfind(",")
        if last_comma > last_dot:
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma and not has_dot:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        import re as _re2
        thousands_pattern = _re2.compile(r"^-?\d{1,3}(\.\d{3})+(\.\d+)?$")
        if thousands_pattern.match(s):
            parts = s.split(".")
            if len(parts) > 2:
                dec = parts.pop()
                s = "".join(parts) + "." + dec
            else:
                s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        return 0.0

def build_pending_model(df: pd.DataFrame) -> dict:
    required = ["N. Ordine","Classe","Cognome Studente","Nome Studente","Docente/ATA","Pagamento","Importo ordine"]
    ensure_cols(df, [c for c in required if c in df.columns] + ["N. Ordine"])

    d = df.copy()
    for c in required:
        if c not in d.columns: d[c] = ""
        d[c] = d[c].map(clean_str)

    pending = d[d["Pagamento"].map(normalize_pagamento).eq("pending")].copy()

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
        if not display_name:
            display_name = "(Docente/ATA non indicato)" if is_doc else "(Studente non indicato)"

        seen[order_id] = {
            "orderId": order_id,
            "classe": classe if classe else "(Senza classe)",
            "displayName": display_name,
            "importo": to_number_it(r["Importo ordine"]),
        }

    unique_orders = list(seen.values())
    by_class = {}
    for o in unique_orders:
        by_class.setdefault(o["classe"], []).append(o)

    classes = sorted(by_class.keys(), key=lambda x: (0 if x == "Docenti/ATA" else 1, str(x)))
    for c in classes:
        by_class[c].sort(key=lambda o: (o["displayName"] or "").lower())

    summary = []
    grand_total = 0.0
    for c in classes:
        arr = by_class[c]
        tot = sum(x["importo"] for x in arr)
        grand_total += tot
        summary.append({"Classe": c, "N ordini": len(arr), "Totale €": round(tot, 2)})

    return {
        "classes": classes,
        "by_class": by_class,
        "summary": pd.DataFrame(summary),
        "grand_total": round(grand_total, 2),
        "n_orders": len(unique_orders),
    }

def pending_pdf_per_class_students(model: dict) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase.pdfmetrics import stringWidth
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    margin = 18 * mm
    row_h = 6 * mm

    def fmt_eur(x: float) -> str:
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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
            while stringWidth(nm, "Helvetica", 11) > max_w and len(nm) > 2:
                nm = nm[:-1]
            if nm != (name or "").strip():
                nm = nm.rstrip() + "…"

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

    def fmt_eur(x: float) -> str:
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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

def page_pending(df_raw: pd.DataFrame) -> None:
    st.subheader("💸 Ordini da pagare")
    st.caption("Estrae gli ordini con Pagamento = pending, consolida per N. Ordine e raggruppa per classe (Docenti/ATA se Classe vuota).")

    required = ["N. Ordine","Classe","Cognome Studente","Nome Studente","Docente/ATA","Pagamento","Importo ordine"]
    missing = [c for c in required if c not in df_raw.columns]
    if missing:
        st.error("Nel file mancano queste colonne per 'Ordini da pagare': " + ", ".join(missing))
        return

    model = build_pending_model(df_raw)

    top1, top2, top3 = st.columns([2,2,3])
    with top1:
        st.metric("Ordini pending", model["n_orders"])
    with top2:
        st.metric("Totale €", f'{model["grand_total"]:.2f}')
    with top3:
        pdfA = pending_pdf_per_class_students(model)
        pdfB = pending_pdf_totals_only(model)
        st.download_button("⬇️ PDF per classe (dettaglio studenti)", data=pdfA, file_name="wupi_ordini_pending_dettaglio_per_classe.pdf", mime="application/pdf")
        st.download_button("⬇️ PDF riepilogo totali (un foglio)", data=pdfB, file_name="wupi_ordini_pending_riepilogo_totali.pdf", mime="application/pdf")

    st.markdown("### Riepilogo per classe")
    st.dataframe(model["summary"], use_container_width=True, hide_index=True)

    st.markdown("### Dettaglio")
    cls = st.selectbox("Classe", options=model["classes"], index=0 if model["classes"] else None)
    if cls:
        det = pd.DataFrame(model["by_class"][cls])
        det = det.rename(columns={"orderId":"N. Ordine","displayName":"Nome","importo":"Importo €"})[["N. Ordine","Nome","Importo €"]]
        st.dataframe(det, use_container_width=True, hide_index=True)


# -------------------------
# Bibbia maker (A3) — da XLSX + mockup batch
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
        if 1 <= len(p1) <= 2:
            return p2
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

    k = (sku_key, model_key, col_key, "")
    if k in mock_map: return mock_map[k]

    k = (sku_key, "", col_key, "")
    if k in mock_map: return mock_map[k]

    return None

def parse_mockup_files(files: list) -> dict:
    def norm_token(t: str) -> str:
        t = (t or "").strip().lower()
        t = re.sub(r"[^a-z0-9]+", "_", t)
        t = re.sub(r"_+", "_", t).strip("_")
        return t

    alias2canon: dict[str, str] = {}
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

    out: dict[tuple[str, str, str, str], bytes] = {}

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
        raw_model_str = " ".join(model_tokens)
        model_key = product_model_key(raw_model_str)

        try:
            data = f.getvalue()
        except Exception:
            try: data = f.read()
            except Exception: continue

        out[(sku_key, model_key, color_key, side_key)] = data
        if raw_color and raw_color != color_key:
            out[(sku_key, model_key, raw_color, side_key)] = data

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
    breakdown = (
        dd.groupby(["SKU", "Nome Prodotto", "Colore", "Taglia"], as_index=False)["Pezzi"].sum()
          .sort_values(["SKU", "Nome Prodotto", "Colore", "Taglia"], kind="stable")
    )

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

        rows = []
        rows.append(f"Neutri: {neutri_qty} pz")
        rows.append(f"Personalizzati: {pers_qty} pz")

        agg = (
            personalizzati.groupby(["ORD_SHOW", "CLS_SHOW", "INC_RAW"], as_index=False)["Pezzi"].sum()
            .sort_values(["ORD_SHOW", "CLS_SHOW", "INC_RAW"], kind="stable")
        )

        for _, r in agg.iterrows():
            ordn = clean_str(r["ORD_SHOW"])
            cls = clean_str(r["CLS_SHOW"])
            inc = clean_str(r["INC_RAW"])
            qty = int(r["Pezzi"])

            left = f"#{ordn}" if ordn else ""
            mid = cls if cls else ""
            base = " · ".join([x for x in [left, mid, inc] if x])
            if qty > 1: rows.append(f"{base} ({qty})")
            else: rows.append(base)

        return "\n".join(rows)

    inc = tmp.groupby(["SKU", "Nome Prodotto", "Colore"], as_index=False).apply(fmt_incisioni)
    if isinstance(inc, pd.Series): inc = inc.reset_index().rename(columns={0: "Incisioni"})
    else: inc = inc.rename(columns={None: "Incisioni"})
    if "Incisioni" not in inc.columns: inc["Incisioni"] = ""

    tot = (
        d.groupby(["SKU", "Nome Prodotto", "Colore", "SKU_KEY", "COL_KEY", "MODEL_KEY"], as_index=False)["Pezzi"].sum()
         .rename(columns={"Pezzi": "Totale"})
    )

    out = tot.merge(
        bd[["SKU", "Nome Prodotto", "Colore", "Taglie"]],
        on=["SKU", "Nome Prodotto", "Colore"],
        how="left"
    )
    out = out.merge(
        inc[["SKU", "Nome Prodotto", "Colore", "Incisioni"]],
        on=["SKU", "Nome Prodotto", "Colore"],
        how="left"
    )

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

    # 1. Logo
    if logo_img:
        c.drawImage(logo_img, margin, h - margin - 16 * mm_to_pt, width=26 * mm_to_pt, height=16 * mm_to_pt, preserveAspectRatio=True, mask="auto")

    # 2. Intestazione (Titolo SKU e Colore)
    c.setFillGray(0)
    c.setFont("Helvetica-Bold", cfg.header_pt)
    c.drawString(margin + (32 * mm_to_pt if logo_img else 0), h - margin - 10 * mm_to_pt, f"{sku_base} — {prod}")
    c.drawRightString(w - margin, h - margin - 10 * mm_to_pt, f"{col}")

    # 3. Disegno delle Immagini (Fronte / Retro)
    img_y0 = margin + footer_h + gap
    img_h = h - margin - header_h - img_y0
    img_w = (w - 2 * margin - gap) / 2

    box1 = (margin, img_y0, img_w, img_h)
    box2 = (margin + img_w + gap, img_y0, img_w, img_h)

    front = find_mockup_bytes(mock_map, sku_key, model_key, col_key, "fronte")
    back = find_mockup_bytes(mock_map, sku_key, model_key, col_key, "retro")

    if front:
        _draw_image_fit(c, front, *box1)
    elif cfg.show_missing_boxes:
        c.setLineWidth(1)
        c.setStrokeGray(0.8)
        c.rect(*box1)
        c.setFillGray(0.5)
        c.setFont("Helvetica", 14)
        c.drawCentredString(box1[0] + box1[2] / 2, box1[1] + box1[3] / 2, "FRONTE MANCANTE")

    if back:
        _draw_image_fit(c, back, *box2)
    elif cfg.show_missing_boxes:
        c.setLineWidth(1)
        c.setStrokeGray(0.8)
        c.rect(*box2)
        c.setFillGray(0.5)
        c.setFont("Helvetica", 14)
        c.drawCentredString(box2[0] + box2[2] / 2, box2[1] + box2[3] / 2, "RETRO MANCANTE")

    # 4. Testo SKU / Colore / Totale
    c.setFillGray(0)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, margin + footer_h - 8 * mm_to_pt, f"SKU: {sku_base}   Colore: {col}   Totale: {totale}")

    box_w = 0
    if has_incisioni:
        box_w = 82 * mm_to_pt

    # 5. Box delle Taglie (Pills)
    pills_left_x = margin
    pills_y = margin + footer_h - 20 * mm_to_pt
    pills_max_w = (w - 2 * margin - box_w - 10 * mm_to_pt) if has_incisioni else (w - 2 * margin)

    items = _parse_taglie_items(taglie)
    cur_x = pills_left_x

    if items:
        pill_font_regular = 15
        pill_font_bold = 15
        pill_h = 26
        pill_radius = 13
        pad_x = 12
        gap_inner = 8
        gap_between = 12
        
        rect_y = pills_y - 6
        text_baseline = rect_y + 8

        for taglia, qty in items:
            taglia_txt = str(taglia)
            qty_txt = str(qty)

            c.setFont("Helvetica", pill_font_regular)
            w1 = c.stringWidth(taglia_txt, "Helvetica", pill_font_regular)
            c.setFont("Helvetica-Bold", pill_font_bold)
            w2 = c.stringWidth(qty_txt, "Helvetica-Bold", pill_font_bold)

            pw = w1 + w2 + (pad_x * 2) + gap_inner

            if cur_x + pw > pills_left_x + pills_max_w:
                break

            c.setFillColorRGB(0.92, 0.92, 0.93)
            c.roundRect(cur_x, rect_y, pw, pill_h, pill_radius, stroke=0, fill=1)
            
            c.setFillGray(0)
            c.setFont("Helvetica", pill_font_regular)
            c.drawString(cur_x + pad_x, text_baseline, taglia_txt)

            c.setFont("Helvetica-Bold", pill_font_bold)
            c.drawString(cur_x + pad_x + w1 + gap_inner, text_baseline, qty_txt)

            cur_x += pw + gap_between

    # 6. Box Personalizzazioni in basso a destra
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
            neutri = lines[0]
            pers = lines[1]
            details = lines[2:]

            c.setFont("Helvetica-Bold", 9)
            c.drawString(bx + 12, by + box_h - 30, f"{neutri}   |   {pers}")
            
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0.85, 0.85, 0.85)
            c.line(bx + 12, by + box_h - 36, bx + box_w - 12, by + box_h - 36)

            y = by + box_h - 50
            max_w = box_w - 24
            for line in details:
                if y < by + 8:
                    break
                txt = line.strip()
                if not txt: continue

                c.setFillColorRGB(0.2, 0.2, 0.2)
                c.circle(bx + 15, y + 2.5, 1.5, stroke=0, fill=1)

                c.setFillGray(0)
                c.setFont("Helvetica", 9)
                while c.stringWidth(txt, "Helvetica", 9) > max_w - 12 and len(txt) > 2:
                    txt = txt[:-1]
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
    # A3 verticale come base della griglia
    page_w, page_h = A3 
    
    # La dimensione "virtuale" è l'A3 orizzontale originale
    vw, vh = landscape(A3) 
    
    cols = 2
    rows = 4
    
    cell_w = page_w / cols
    cell_h = page_h / rows
    
    # Scaliamo perfettamente l'A3 Orizzontale dentro la cella 1/8 di un A3 Verticale
    scale_factor = cell_w / vw 
    
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))

    logo_img = None
    if brand_logo:
        try: logo_img = ImageReader(io.BytesIO(brand_logo))
        except Exception: logo_img = None

    count = 0
    for _, r in variants.iterrows():
        col_idx = count % cols
        row_idx = (count // cols) % rows
        
        px = col_idx * cell_w
        py = page_h - (row_idx + 1) * cell_h
        
        c.saveState()
        c.translate(px, py)
        c.scale(scale_factor, scale_factor)
        _draw_bibbia_variant(c, r, mock_map, cfg, logo_img, vw, vh)
        c.restoreState()
        
        # Disegniamo una sottile linea grigia attorno alla cella come guida per il taglio
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.rect(px, py, cell_w, cell_h)
        
        count += 1
        if count % (cols * rows) == 0:
            c.showPage()
            
    if count % (cols * rows) != 0:
        c.showPage()

    c.save()
    buf.seek(0)
    return buf.getvalue()

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

    by_class = (
        d.groupby("Classe", dropna=False, as_index=False)
         .agg(Pezzi=("Pezzi", "sum"), **{"Importo ex IVA": ("Importo ex IVA", "sum"), "Margine": ("Margine", "sum")})
         .sort_values(["Classe"], kind="stable")
         .reset_index(drop=True)
    )
    if not by_class.empty and (by_class["Classe"] == "Docenti / ATA").any():
        first = by_class[by_class["Classe"] == "Docenti / ATA"]
        rest = by_class[by_class["Classe"] != "Docenti / ATA"]
        by_class = pd.concat([first, rest], ignore_index=True)

    by_product = (
        d.groupby(["SKU", "Nome Prodotto"], dropna=False, as_index=False)
         .agg(Pezzi=("Pezzi", "sum"),
              **{"Prezzo vendita ex IVA": ("Prezzo vendita ex IVA", "mean"),
                 "Prezzo acquisto": ("Prezzo acquisto", "mean"),
                 "Importo ex IVA": ("Importo ex IVA", "sum"),
                 "Margine": ("Margine", "sum")})
         .sort_values(["Pezzi", "Nome Prodotto"], ascending=[False, True], kind="stable")
         .reset_index(drop=True)
    )

    by_color = (
        d.groupby(["SKU", "Colore"], dropna=False, as_index=False)
         .agg(Pezzi=("Pezzi", "sum"),
              **{"Prezzo vendita ex IVA": ("Prezzo vendita ex IVA", "mean"),
                 "Prezzo acquisto": ("Prezzo acquisto", "mean"),
                 "Importo ex IVA": ("Importo ex IVA", "sum"),
                 "Margine": ("Margine", "sum")})
         .sort_values(["Pezzi", "Colore"], ascending=[False, True], kind="stable")
         .reset_index(drop=True)
    )

    return d, total_orders, total_pieces, total_amount, total_margin, by_class, by_product, by_color

def _eur(x: float) -> str:
    return f"€ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def page_finanze(df_norm: pd.DataFrame) -> None:
    st.subheader("Finanze")
    st.caption("Riepilogo economico della tornata/scuola. Tutti i valori qui sono ex IVA 22% (Prezzo unitario / 1,22).")

    costs = load_costs()
    _, total_orders, total_pieces, total_amount, total_margin, by_class, by_product, by_color = finance_summary(df_norm, costs)

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Ordini", total_orders)
    with c2: st.metric("Pezzi totali", total_pieces)
    with c3: st.metric("Importo ex IVA", _eur(total_amount))
    with c4: st.metric("Margine", _eur(total_margin))

    st.markdown("### Prezzi acquisto")
    cost_view = by_product[["SKU", "Nome Prodotto", "Prezzo acquisto"]].copy()
    cost_view["Chiave"] = cost_view.apply(lambda r: _cost_key(r["SKU"], r["Nome Prodotto"]), axis=1)
    edited = st.data_editor(
        cost_view,
        hide_index=True,
        use_container_width=True,
        disabled=["SKU", "Nome Prodotto", "Chiave"],
        column_config={
            "Prezzo acquisto": st.column_config.NumberColumn("Prezzo acquisto", step=0.01, format="%.2f"),
            "Chiave": None,
        },
        key="finance_cost_editor",
    )
    if st.button("💾 Salva prezzi acquisto"):
        new_costs = load_costs()
        for _, r in edited.iterrows():
            try: new_costs[str(r["Chiave"])] = float(r["Prezzo acquisto"])
            except Exception: pass
        save_costs(new_costs)
        st.success("Prezzi acquisto salvati.")
        st.rerun()

    costs = load_costs()
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

@st.dialog("Associa immagine mancante")
def upload_missing_modal(r, side):
    st.markdown(f"Stai caricando il **{side.upper()}** per:")
    st.markdown(f"**{r['SKU']}** — {r['Nome Prodotto']} ({r['Colore']})")
    
    uid = hashlib.md5(f"{r['SKU']}_{r['Nome Prodotto']}_{r['Colore']}_{side}".encode()).hexdigest()[:10]
    
    fup = st.file_uploader(f"Carica file (.jpg, .png)", type=["png","jpg","jpeg"], key=f"modal_up_{uid}")
    if fup:
        if st.button("Salva abbinamento", type="primary", use_container_width=True):
            save_path = APP_SUPPORT / "manual_mockups" / f"{r['SKU_KEY']}__{r.get('MODEL_KEY','')}__{r['COL_KEY']}__{side}{Path(fup.name).suffix or '.jpg'}"
            save_path.parent.mkdir(parents=True, exist_ok=True)
            save_path.write_bytes(fup.getvalue())
            
            try: raw_manual = json_loads(BIBBIA_MANUAL_PATH.read_text(encoding="utf-8"))
            except Exception: raw_manual = {}
            
            keybase = [str(r["SKU_KEY"]), str(r.get("MODEL_KEY","")), str(r["COL_KEY"])]
            raw_manual["|||".join(keybase + [side])] = str(save_path)
            save_manual_mockups(raw_manual)
            
            st.rerun()

def page_bibbia(df_norm: pd.DataFrame) -> None:
    st.subheader("Bibbia maker (A3)")
    st.caption("Carica i mockup in batch (JPG/PNG) con naming permissivo: SKU_modello_colore_fronte / SKU_modello_colore_retro.")

    if "bibbia_uploader_ver" not in st.session_state: st.session_state["bibbia_uploader_ver"] = 0

    mock_files = st.file_uploader(
        "Carica qui le immagini Mockup",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"bibbia_mockups_{st.session_state['bibbia_uploader_ver']}",
    )
    
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
    
    st.markdown("""
    <style>
    .bibbia-header { font-weight: 600; padding-bottom: 8px; border-bottom: 2px solid #e5e5ea; color: #86868b; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }
    </style>
    """, unsafe_allow_html=True)
    
    h1, h2, h3, h4, h5, h6 = st.columns([1.5, 2.5, 3, 2.5, 1, 1])
    with h1: st.markdown("<div class='bibbia-header'>SKU / Colore</div>", unsafe_allow_html=True)
    with h2: st.markdown("<div class='bibbia-header'>Prodotto</div>", unsafe_allow_html=True)
    with h3: st.markdown("<div class='bibbia-header'>Taglie</div>", unsafe_allow_html=True)
    with h4: st.markdown("<div class='bibbia-header'>Incisioni</div>", unsafe_allow_html=True)
    with h5: st.markdown("<div class='bibbia-header'>Fronte</div>", unsafe_allow_html=True)
    with h6: st.markdown("<div class='bibbia-header'>Retro</div>", unsafe_allow_html=True)

    for _, r in variants.iterrows():
        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 2.5, 3, 2.5, 1, 1])
        
        with c1:
            st.markdown(f"**{r['SKU']}**<br><span style='color:#555; font-size:14px;'>{r['Colore']}</span>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div style='font-weight:500; margin-top:2px;'>{r['Nome Prodotto']}</div>", unsafe_allow_html=True)
        with c3:
            items = _parse_taglie_items(r['Taglie'])
            if items:
                pills = "".join([f'<span class="chip" style="margin-bottom:4px;">{t} <span class="q">{q}</span></span>' for t, q in items])
                st.markdown(f"<div class='chips'>{pills}</div>", unsafe_allow_html=True)
        with c4:
            inc_html = str(r['Incisioni']).replace('\n', '<br>')
            st.markdown(f"<div style='font-size:12px; color:#666; line-height:1.3;'>{inc_html}</div>", unsafe_allow_html=True)
        with c5:
            if r["Fronte"] == "✅":
                st.markdown("<div style='margin-top:6px;'>✅</div>", unsafe_allow_html=True)
            else:
                if st.button("❌", key=f"f_{r['SKU_KEY']}_{r['COL_KEY']}_{r.get('MODEL_KEY','')}", help="Aggiungi Fronte"):
                    upload_missing_modal(r, "fronte")
        with c6:
            if r["Retro"] == "✅":
                st.markdown("<div style='margin-top:6px;'>✅</div>", unsafe_allow_html=True)
            else:
                if st.button("❌", key=f"r_{r['SKU_KEY']}_{r['COL_KEY']}_{r.get('MODEL_KEY','')}", help="Aggiungi Retro"):
                    upload_missing_modal(r, "retro")
        
        st.markdown("<hr style='margin:0.25em 0; border:none; border-bottom:1px solid #f0f0f2;'>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    with st.expander("Opzioni PDF A3", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1: margin_mm = st.number_input("Margine (mm)", min_value=4.0, max_value=30.0, value=10.0, step=1.0, key="bibbia_margin_mm")
        with c2: gap_mm = st.number_input("Spazio tra fronte/retro (mm)", min_value=2.0, max_value=30.0, value=6.0, step=1.0, key="bibbia_gap_mm")
        with c3: show_missing = st.checkbox("Mostra riquadri MANCANTE", value=True, key="bibbia_show_missing")

    with st.expander("Font", expanded=False):
        f1, f2 = st.columns(2)
        with f1: header_pt = st.number_input("Titolo (pt)", min_value=12.0, max_value=36.0, value=18.0, step=0.5, key="bibbia_header_pt")
        with f2: caption_pt = st.number_input("Caption (pt)", min_value=8.0, max_value=20.0, value=11.0, step=0.5, key="bibbia_caption_pt")

    cfg = BibbiaCfg(margin_mm=float(margin_mm), gap_mm=float(gap_mm), header_pt=float(header_pt), caption_pt=float(caption_pt), show_missing_boxes=bool(show_missing))
    logo_bytes = (LOGO_PATH.read_bytes() if LOGO_PATH.exists() else None)

    st.markdown("### Generazione PDF")
    colA, colB = st.columns(2)
    with colA:
        if st.button("📄 Prepara PDF Singolo (1 per A3)", use_container_width=True):
            st.session_state['bibbia_mode'] = 'single'
            st.rerun()
    with colB:
        if st.button("🗂 Prepara PDF Griglia (8 per A3)", type="primary", use_container_width=True):
            st.session_state['bibbia_mode'] = 'grid'
            st.rerun()
            
    mode = st.session_state.get('bibbia_mode')
    if mode == 'single':
        pdf = make_bibbia_pdf_single(variants, mock_map, cfg, brand_logo=logo_bytes)
        st.success("PDF Singolo pronto per il download!")
        st.download_button("⬇️ Scarica PDF Singolo", data=pdf, file_name="wupi_bibbia_singola.pdf", mime="application/pdf", use_container_width=True)
    elif mode == 'grid':
        pdf = make_bibbia_pdf_grid(variants, mock_map, cfg, brand_logo=logo_bytes)
        st.success("PDF Griglia pronto per il download!")
        st.download_button("⬇️ Scarica PDF Griglia (8 in 1)", data=pdf, file_name="wupi_bibbia_griglia.pdf", mime="application/pdf", use_container_width=True)

# -------------------------
# UI
# -------------------------
def main() -> None:
    st.set_page_config(
        page_title="WUPI Suite",
        layout="wide",
        page_icon=str(FAVICON_PATH) if FAVICON_PATH.exists() else "🧰",
    )

    top_l, top_r = st.columns([7, 1])
    with top_l:
        st.title("WUPI Suite")
        st.caption(f"Build: STUDIO_v4_GRID (stable) • {Path(__file__).resolve()}")
    with top_r:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), use_container_width=True)

    uploaded = st.file_uploader("Carica Excel (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.info("Carica un file per iniziare.")
        return

    file_bytes = uploaded.getvalue()
    sig = file_sig(file_bytes)

    df_raw = pd.read_excel(io.BytesIO(file_bytes))
    df = df_normalize(df_raw)

    state = load_state()
    if "confirmed" not in st.session_state or st.session_state.get("confirmed_sig") != sig:
        st.session_state["confirmed"] = set(normalize_key(k) for k in state.get(sig, []))
        st.session_state["confirmed_sig"] = sig
    confirmed = st.session_state["confirmed"]

    tabs = st.tabs(["📦 Report acquisto", "🏷 Etichette", "💸 Ordini da pagare", "📖 Bibbia maker", "💰 Finanze"])
    with tabs[0]:
        st.subheader("Pivot ordine fornitore")
        st.caption("0 nascosti, Totale fisso a destra (bold). Le righe confermate diventano grigie chiare.")
        piv_full = pivot_report(df)

        q = st.text_input("🔍 Cerca (SKU / Prodotto / Colore)", key="pivot_search")
        piv_view = piv_full
        if q:
            qq = str(q).strip()
            piv_view = piv_full[
                piv_full["SKU"].astype(str).str.contains(qq, case=False, na=False)
                | piv_full["Nome Prodotto"].astype(str).str.contains(qq, case=False, na=False)
                | piv_full["Colore"].astype(str).str.contains(qq, case=False, na=False)
            ].copy()

        render_pivot_html(piv_view, confirmed)
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
                st.session_state["pair_widget_ver"] += 1 
                st.rerun()
        with b:
            sel = st.selectbox(
                "Seleziona SKU",
                options=options,
                index=st.session_state["pair_idx"] if len(options) else 0,
                key=f"pair_select_{st.session_state['pair_widget_ver']}",
                label_visibility="collapsed",
            )
        with c:
            if st.button("›", key="next_pair", use_container_width=True):
                st.session_state["pair_idx"] = min(len(options) - 1, st.session_state["pair_idx"] + 1)
                st.session_state["pair_widget_ver"] += 1
                st.rerun()

        if len(options) > 0 and sel in options:
            st.session_state["pair_idx"] = options.index(sel)

        sku = sel.split(" — ", 1)[0].strip()
        prod = sel.split(" — ", 1)[1].strip()

        render_color_cards(df, sku, prod, confirmed, sig, state)

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
        with col2:
            logo_up = st.file_uploader("Logo (PNG) opzionale", type=["png"])
        logo_bytes = logo_up.getvalue() if logo_up else (LOGO_PATH.read_bytes() if LOGO_PATH.exists() else None)

        cfg = LabelCfg(w_mm=float(w), h_mm=float(h), margin_mm=float(m), logo_w_mm=float(logo_w), title_pt=float(title_pt), header_pt=float(header_pt), row_pt=float(row_pt), row_h_mm=float(row_h), strip_modello=bool(strip_modello))
        if st.button("Genera PDF etichette", type="primary"):
            pdf = make_labels_pdf(df, logo_bytes, cfg)
            st.download_button("⬇️ Scarica PDF", data=pdf, file_name="wupi_etichette.pdf", mime="application/pdf")

    with tabs[2]:
        page_pending(df_raw)

    with tabs[3]:
        page_bibbia(df)

    with tabs[4]:
        page_finanze(df)

if __name__ == "__main__":
    global_css()
    main()
