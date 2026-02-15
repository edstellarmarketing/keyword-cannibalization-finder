"""
Keyword Cannibalization Finder — Edstellar Edition
Built on top of Lee Foot's original concept with custom filtering for
geo-templated pages (corporate-training-companies-<country>, skills-in-demand-in-<country>, etc.)

Run with:
    pip install streamlit pandas openpyxl
    streamlit run keyword_cannibalization_app.py
"""

import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Keyword Cannibalization Finder",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Import fonts ── */
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;1,300&display=swap');

/* ── Root variables — works on both light & dark Streamlit themes ── */
:root {
    --orange:    #E8651A;
    --orange-lt: rgba(232,101,26,.18);
    --orange-br: rgba(232,101,26,.5);
    --sky:       #4A9FD5;
    --pale-blue: rgba(74,159,213,.15);
    --white:     #FFFFFF;
    --radius:    10px;
    /* Semantic — always visible regardless of theme */
    --txt:       #FFFFFF;          /* primary text on custom elements */
    --txt-muted: rgba(255,255,255,.65);
    --txt-label: rgba(255,255,255,.5);
    --card-bg:   rgba(255,255,255,.07);
    --card-bd:   rgba(255,255,255,.12);
    --shadow:    0 2px 14px rgba(0,0,0,.25);
    --shadow-lg: 0 6px 30px rgba(0,0,0,.35);
    /* Severity */
    --red:       #FF6B6B;
    --red-bg:    rgba(255,107,107,.15);
    --amber:     #FFB340;
    --amber-bg:  rgba(255,179,64,.15);
    --green:     #4ADE80;
    --green-bg:  rgba(74,222,128,.15);
}

/* ── Base typography ── */
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

/* ── Hide default Streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }

/* ── App header banner ── */
.app-header {
    background: linear-gradient(135deg, #0F2340 0%, #1B4F8A 60%, #2E6DA4 100%);
    border-radius: var(--radius);
    padding: 32px 36px 28px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.app-header::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 220px; height: 220px;
    border-radius: 50%;
    background: rgba(74,159,213,.15);
}
.app-header::after {
    content: '';
    position: absolute;
    bottom: -60px; left: 40%;
    width: 160px; height: 160px;
    border-radius: 50%;
    background: rgba(232,101,26,.12);
}
.app-header h1 {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 2rem;
    color: #FFFFFF !important;
    margin: 0 0 6px;
    position: relative; z-index: 1;
}
.app-header p {
    color: rgba(255,255,255,.75) !important;
    font-size: 0.95rem;
    margin: 0;
    position: relative; z-index: 1;
}
.app-header .badge {
    display: inline-block;
    background: rgba(232,101,26,.25);
    border: 1px solid rgba(232,101,26,.5);
    color: #FFB380 !important;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: .06em;
    text-transform: uppercase;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 10px;
}

/* ── Section headers ── */
.section-hdr {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 1.15rem;
    color: #FFFFFF !important;
    border-left: 4px solid var(--orange);
    padding-left: 12px;
    margin: 28px 0 16px;
}

/* ── KPI cards ── */
.kpi-row { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 24px; }
.kpi-card {
    flex: 1; min-width: 130px;
    background: var(--card-bg);
    border: 1px solid var(--card-bd);
    border-radius: var(--radius);
    padding: 18px 20px 14px;
    box-shadow: var(--shadow);
    transition: box-shadow .2s;
}
.kpi-card:hover { box-shadow: var(--shadow-lg); }
.kpi-card .kpi-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: .07em;
    text-transform: uppercase;
    color: var(--txt-label) !important;
    margin-bottom: 6px;
}
.kpi-card .kpi-value {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 1.9rem;
    color: #FFFFFF !important;
    line-height: 1;
}
.kpi-card .kpi-sub {
    font-size: 0.78rem;
    color: var(--txt-muted) !important;
    margin-top: 4px;
}
.kpi-card.accent .kpi-value { color: var(--orange) !important; }
.kpi-card.danger .kpi-value { color: var(--red) !important; }
.kpi-card.success .kpi-value { color: var(--green) !important; }

/* ── Info box ── */
.info-box {
    background: rgba(74,159,213,.15);
    border-left: 4px solid #4A9FD5;
    border-radius: 0 var(--radius) var(--radius) 0;
    padding: 14px 18px;
    font-size: 0.88rem;
    color: #A8D8F0 !important;
    margin: 16px 0;
}

/* ── Filter note ── */
.filter-note {
    background: rgba(232,101,26,.15);
    border-left: 4px solid var(--orange);
    border-radius: 0 var(--radius) var(--radius) 0;
    padding: 10px 14px;
    font-size: 0.82rem;
    color: #FFD0A8 !important;
    margin: 8px 0 16px;
}

/* ── Rec card ── */
.rec-card {
    background: var(--card-bg);
    border: 1px solid var(--card-bd);
    border-radius: var(--radius);
    padding: 16px 18px;
    margin-bottom: 10px;
}
.rec-card h4 {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 0.92rem;
    color: #FFFFFF !important;
    margin: 0 0 6px;
}
.rec-card p { font-size: 0.84rem; color: var(--txt-muted) !important; margin: 0; }

/* ── Streamlit dataframe tweak ── */
.stDataFrame { border-radius: var(--radius); overflow: hidden; }

/* ── Download btn override ── */
.stDownloadButton > button {
    background: rgba(255,255,255,.1) !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(255,255,255,.25) !important;
    border-radius: var(--radius) !important;
    font-weight: 600 !important;
    padding: 10px 22px !important;
    transition: all .2s !important;
}
.stDownloadButton > button:hover {
    background: rgba(255,255,255,.18) !important;
    border-color: rgba(255,255,255,.4) !important;
}

/* ── Primary button ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--orange) 0%, #c45510 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: var(--radius) !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 12px 32px !important;
    box-shadow: 0 4px 14px rgba(232,101,26,.4) !important;
    transition: all .2s !important;
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 20px rgba(232,101,26,.55) !important;
}

/* ── Sidebar: force visible label text ── */
[data-testid="stSidebar"] .stNumberInput label,
[data-testid="stSidebar"] .stCheckbox label,
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div {
    font-size: 0.83rem;
}
.sidebar-section {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 0.78rem;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: var(--orange) !important;
    padding: 12px 0 6px;
    border-top: 1px solid rgba(255,255,255,.1);
    margin-top: 8px;
}

/* ── Tab bar ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 2px solid rgba(255,255,255,.1);
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Syne', sans-serif;
    font-weight: 600;
    font-size: 0.88rem;
    padding: 8px 20px;
    border-radius: 8px 8px 0 0;
}
.stTabs [aria-selected="true"] {
    border-bottom: 3px solid var(--orange) !important;
}

/* ── Responsive ── */
@media (max-width: 768px) {
    .kpi-row { gap: 10px; }
    .kpi-card .kpi-value { font-size: 1.5rem; }
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS — Edstellar templatized page patterns
# ══════════════════════════════════════════════════════════════════════════════

COUNTRIES = (
    r'(singapore|australia|malaysia|canada|nigeria|ireland|philippines|south-africa|'
    r'new-zealand|egypt|kenya|greece|india|uk|usa|germany|france|uae|saudi-arabia|'
    r'italy|norway|sweden|belgium|south-korea|japan|china|brazil|austria|bahrain|'
    r'botswana|cyprus|denmark|finland|dubai|spain|portugal|netherlands|poland|'
    r'switzerland|turkey|thailand|indonesia|vietnam|qatar|kuwait|oman|jordan|'
    r'pakistan|bangladesh|sri-lanka|nepal|myanmar|hong-kong|taiwan|mexico|argentina|'
    r'colombia|chile|peru|ghana|tanzania|uganda|ethiopia|zimbabwe|zambia|morocco|'
    r'algeria|tunisia|senegal|ivory-coast|cameroon|new-york|london|texas|california|florida)'
)

TEMPLATE_PATTERNS = [
    (re.compile(r'corporate-training-companies-' + COUNTRIES, re.I),
     "corporate-training-companies-<country>"),
    (re.compile(r'skills-in-demand-in-' + COUNTRIES, re.I),
     "skills-in-demand-in-<country>"),
    (re.compile(r'skills-in-demand-' + COUNTRIES, re.I),
     "skills-in-demand-<country>"),
    (re.compile(r'^[a-z]+-work-culture$', re.I),
     "<country>-work-culture"),
    (re.compile(r'corporate-training-in-' + COUNTRIES, re.I),
     "corporate-training-in-<country>"),
    (re.compile(r'best-.*-training-companies-' + COUNTRIES, re.I),
     "best-*-training-companies-<country>"),
    (re.compile(r'top-.*-training-companies-' + COUNTRIES, re.I),
     "top-*-training-companies-<country>"),
]


def is_template(slug: str) -> bool:
    return any(rx.search(slug) for rx, _ in TEMPLATE_PATTERNS)


def get_base_slug(url: str) -> str:
    """Extract the slug portion from a full URL or bare slug."""
    # Remove protocol + domain if present
    url = re.sub(r'^https?://[^/]+/', '', str(url))
    # Remove trailing slashes
    url = url.rstrip('/')
    # Take only the last path segment
    return url.split('/')[-1] if '/' in url else url


# ══════════════════════════════════════════════════════════════════════════════
# DATA PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def read_gsc_data(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise column names from various GSC export formats.
    
    Primary format (Edstellar GSC export):
        Query | Landing Page | Url Clicks | Impressions | URL CTR | Average Position
    """
    mapping = {
        # Query
        'Query': 'query', 'Top queries': 'query', 'Queries': 'query',
        # Page / Landing Page
        'Landing Page': 'page', 'Page': 'page', 'Top pages': 'page',
        'Pages': 'page', 'URL': 'page',
        # Clicks
        'Url Clicks': 'clicks', 'Clicks': 'clicks',
        # Impressions
        'Impressions': 'impressions',
        # CTR
        'URL CTR': 'ctr', 'CTR': 'ctr', 'CTR (%)': 'ctr',
        'Click Through Rate': 'ctr',
        # Position
        'Average Position': 'position', 'Average position': 'position',
        'Avg Position': 'position', 'Avg. position': 'position',
        'Position': 'position',
        # Competing pages (optional)
        'Competing Pages': 'competing_pages_raw',
    }
    df = df.rename(columns=mapping)
    df.columns = df.columns.str.strip()

    # Normalise to lowercase for internal processing
    col_lower = {c: c.lower() for c in df.columns}
    df = df.rename(columns=col_lower)

    required = ['query', 'page', 'clicks', 'impressions', 'position']
    missing  = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required columns: {', '.join(missing)}. "
            f"Expected: Query, Landing Page, Url Clicks, Impressions, URL CTR, Average Position"
        )

    df['clicks']      = pd.to_numeric(df['clicks'],      errors='coerce').fillna(0).astype(int)
    df['impressions'] = pd.to_numeric(df['impressions'], errors='coerce').fillna(0).astype(int)
    df['position']    = pd.to_numeric(df['position'],    errors='coerce').fillna(0)

    if 'ctr' in df.columns:
        if df['ctr'].dtype == object:
            df['ctr'] = df['ctr'].astype(str).str.rstrip('%')
            df['ctr'] = pd.to_numeric(df['ctr'], errors='coerce').fillna(0)
            if df['ctr'].max() > 1:
                df['ctr'] = df['ctr'] / 100
        else:
            df['ctr'] = pd.to_numeric(df['ctr'], errors='coerce').fillna(0)
            if df['ctr'].max() > 1:
                df['ctr'] = df['ctr'] / 100
    else:
        df['ctr'] = 0.0

    return df


# Display column name mapping — internal name → Edstellar export label
DISPLAY_COLS = {
    'query':            'Query',
    'slug':             'Landing Page',
    'page':             'Landing Page',
    'clicks':           'Url Clicks',
    'impressions':      'Impressions',
    'ctr':              'URL CTR (%)',
    'position':         'Average Position',
    'competing_pages':  'Competing Pages',
    'severity':         'Severity',
}

def rename_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Rename internal column names to Edstellar GSC export labels."""
    return df.rename(columns=DISPLAY_COLS)


def apply_filters(df: pd.DataFrame,
                  pos_min: float, pos_max: float,
                  min_impressions: int, min_clicks: int,
                  filter_anchors: bool, filter_templates: bool) -> tuple[pd.DataFrame, dict]:
    """Apply all configured filters and return filtered df + audit log."""
    audit = {}
    audit['before'] = len(df)

    # Anchor filter
    if filter_anchors:
        before = len(df)
        df = df[~df['page'].astype(str).str.contains('#', na=False)].copy()
        audit['anchors_removed'] = before - len(df)
    else:
        audit['anchors_removed'] = 0

    # Templatized page filter
    if filter_templates:
        before = len(df)
        df['_slug'] = df['page'].astype(str).apply(get_base_slug)
        df = df[~df['_slug'].apply(is_template)].copy()
        audit['templates_removed'] = before - len(df)
    else:
        audit['templates_removed'] = 0

    # Position filter
    df = df[(df['position'] >= pos_min) & (df['position'] <= pos_max)].copy()

    # Impression / click filters
    df = df[(df['impressions'] >= min_impressions) & (df['clicks'] >= min_clicks)].copy()

    audit['after'] = len(df)
    return df, audit


def find_cannibalization(df: pd.DataFrame, min_pages: int) -> pd.DataFrame:
    """Identify queries where multiple pages compete."""
    if df.empty:
        return pd.DataFrame()

    slug_col = '_slug' if '_slug' in df.columns else 'page'

    agg = df.groupby(['query', slug_col]).agg(
        clicks=('clicks', 'sum'),
        impressions=('impressions', 'sum'),
        ctr=('ctr', 'mean'),
        position=('position', 'mean'),
    ).reset_index().rename(columns={slug_col: 'slug'})

    agg['position'] = agg['position'].round(1)
    agg['ctr']      = (agg['ctr'] * 100).round(2)

    pages_per_query           = agg.groupby('query')['slug'].transform('count')
    agg['competing_pages']    = pages_per_query
    cannibs                   = agg[agg['competing_pages'] >= min_pages].copy()

    return cannibs.sort_values(['competing_pages', 'impressions'], ascending=[False, False])


def build_query_summary(df: pd.DataFrame) -> pd.DataFrame:
    """One-row-per-query grouped view — uses Edstellar GSC column labels."""
    rows = []
    for q, grp in df.groupby('query'):
        # Pick the canonical "best" page by traffic authority:
        # Score = impressions + (clicks * 10) so clicks break ties on equal impressions.
        # This ensures we always recommend consolidating INTO the page with real traffic,
        # not the one that merely has the lowest position number.
        grp = grp.copy()
        grp['_score'] = grp['impressions'] + (grp['clicks'] * 10)
        best = grp.sort_values('_score', ascending=False).iloc[0]
        rows.append({
            'Query':                   q,
            'Competing Pages':         len(grp),
            'Url Clicks':              int(grp['clicks'].sum()),
            'Impressions':             int(grp['impressions'].sum()),
            'URL CTR (%)':             round(grp['ctr'].mean(), 2),
            'Best Average Position':   round(grp['position'].min(), 1),
            'Worst Average Position':  round(grp['position'].max(), 1),
            'Position Spread':         round(grp['position'].max() - grp['position'].min(), 1),
            'Best Landing Page':       best['slug'],
            'All Landing Pages':       ' | '.join(
                grp.sort_values('_score', ascending=False)['slug'].tolist()
            ),
        })
    out = pd.DataFrame(rows).sort_values('Impressions', ascending=False)
    return out


def severity(pos: float, impressions: int) -> str:
    if pos <= 10 and impressions >= 1000: return 'High'
    if pos <= 20 and impressions >= 200:  return 'Medium'
    return 'Low'


def to_excel(df_dict: dict) -> bytes:
    """Export multiple DataFrames to a single xlsx."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return buf.getvalue()


def to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


def generate_high_severity_docx(cannibs: pd.DataFrame,
                                 query_sum_df: pd.DataFrame) -> bytes:
    """
    Generate a Word .docx report matching the High Severity tab layout:
    - Cover section with summary stats
    - One section per query: position / impressions header, URL table, suggested action
    """
    import subprocess, json, tempfile, os

    high_queries = query_sum_df[query_sum_df['_sev'] == 'High']['Query'].tolist()
    if not high_queries:
        return b""

    # Build data structure to pass to JS
    report_data = []
    for q in high_queries:
        qdata = cannibs[cannibs['query'] == q].copy()
        qdata['_score'] = qdata['impressions'] + (qdata['clicks'] * 10)
        qdata = qdata.sort_values('_score', ascending=False)
        best_slug = qdata.iloc[0]['slug']
        weaker    = qdata.iloc[1:]['slug'].tolist()
        best_pos  = round(float(qdata['position'].min()), 1)
        total_imp = int(qdata['impressions'].sum())
        rows = []
        for _, r in qdata.iterrows():
            rows.append({
                'slug':      str(r['slug']),
                'clicks':    int(r['clicks']),
                'impressions': int(r['impressions']),
                'ctr':       round(float(r['ctr']), 2),
                'position':  round(float(r['position']), 1),
                'competing': int(r['competing_pages']),
                'isBest':    str(r['slug']) == best_slug,
            })
        report_data.append({
            'query':      q,
            'bestSlug':   best_slug,
            'weakerSlugs': weaker,
            'bestPos':    best_pos,
            'totalImp':   total_imp,
            'numPages':   len(qdata),
            'rows':       rows,
        })

    # Summary stats
    summary = {
        'totalHigh':   len(high_queries),
        'totalImp':    int(query_sum_df[query_sum_df['_sev']=='High']['Impressions'].sum()),
        'totalClicks': int(query_sum_df[query_sum_df['_sev']=='High']['Url Clicks'].sum()),
        'date':        pd.Timestamp.now().strftime('%B %d, %Y'),
    }

    payload = json.dumps({'summary': summary, 'queries': report_data})

    # Write JS script
    js_script = r"""
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require('docx');

const data    = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const summary = data.summary;
const queries = data.queries;

const NAVY='0F2340', BLUE='1B4F8A', MID_BLUE='2E6DA4', ORANGE='E8651A';
const WHITE='FFFFFF', BEST_BG='E8F5EE', TABLE_HD='1B4F8A';
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: 'D4DFE9' };
const cellBorder = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

function spacer(pts) {
  return new Paragraph({ children:[new TextRun('')], spacing:{before:pts*20,after:0} });
}
function sectionRule(color) {
  return new Paragraph({
    border:{ bottom:{style:BorderStyle.SINGLE, size:6, color:color||ORANGE, space:1} },
    spacing:{before:0,after:120}, children:[],
  });
}
function hdrCell(text, width) {
  return new TableCell({
    borders:cellBorder, width:{size:width,type:WidthType.DXA},
    shading:{fill:TABLE_HD, type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:100,right:100},
    children:[new Paragraph({children:[new TextRun({text,bold:true,color:WHITE,size:18,font:'Arial'})]})],
  });
}
function dataCell(text, width, isBest, alignRight) {
  return new TableCell({
    borders:cellBorder, width:{size:width,type:WidthType.DXA},
    shading:{fill:isBest?BEST_BG:WHITE, type:ShadingType.CLEAR},
    margins:{top:60,bottom:60,left:100,right:100},
    children:[new Paragraph({
      alignment:alignRight?AlignmentType.RIGHT:AlignmentType.LEFT,
      children:[new TextRun({text:String(text),size:18,font:'Arial',bold:isBest,color:isBest?'1A6B3A':'1A1A2E'})],
    })],
  });
}
function kpiCell(label, value, width) {
  return new TableCell({
    borders:cellBorder, width:{size:width,type:WidthType.DXA},
    shading:{fill:'D6E8F5', type:ShadingType.CLEAR},
    margins:{top:100,bottom:100,left:140,right:140},
    children:[
      new Paragraph({children:[new TextRun({text:String(value),bold:true,size:32,font:'Arial',color:NAVY})]}),
      new Paragraph({children:[new TextRun({text:label,size:16,font:'Arial',color:'5A7FA0'})]}),
    ],
  });
}

const children = [];

// Title block
children.push(new Paragraph({
  children:[new TextRun({text:'Edstellar  ·  Keyword Cannibalization Report',size:20,font:'Arial',bold:true,color:WHITE})],
  shading:{fill:NAVY,type:ShadingType.CLEAR},
  spacing:{before:0,after:0}, indent:{left:200},
}));
children.push(new Paragraph({
  children:[new TextRun({text:'High Severity Issues — Urgent Fixes',size:48,bold:true,font:'Arial',color:WHITE})],
  shading:{fill:NAVY,type:ShadingType.CLEAR},
  spacing:{before:120,after:0}, indent:{left:200},
}));
children.push(new Paragraph({
  children:[new TextRun({text:'Best position ≤10  ·  Impressions ≥1,000  ·  Generated: '+summary.date,size:20,font:'Arial',color:'FFB380'})],
  shading:{fill:NAVY,type:ShadingType.CLEAR},
  spacing:{before:80,after:280}, indent:{left:200},
}));

// KPI row
const kpiW = Math.floor(9360/3);
children.push(new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[kpiW,kpiW,9360-kpiW*2],
  rows:[new TableRow({children:[
    kpiCell('High Severity Queries', summary.totalHigh, kpiW),
    kpiCell('Total Impressions at Stake', summary.totalImp.toLocaleString(), kpiW),
    kpiCell('Total Clicks at Stake', summary.totalClicks.toLocaleString(), 9360-kpiW*2),
  ]})],
}));
children.push(spacer(10));
children.push(new Paragraph({
  children:[new TextRun({text:'🚨  These queries rank on page 1 but split click potential across multiple URLs. Consolidating them will have the most direct impact on organic traffic.',size:18,font:'Arial',color:'5A3000',italics:true})],
  shading:{fill:'FDE8D8',type:ShadingType.CLEAR},
  border:{left:{style:BorderStyle.SINGLE,size:16,color:ORANGE}},
  indent:{left:200,right:200}, spacing:{before:100,after:100},
}));
children.push(spacer(14));
children.push(sectionRule(ORANGE));

// Per-query sections
const COL_WIDTHS=[3200,900,1300,1000,1360,1600];
const COL_HEADERS=['Landing Page','Url Clicks','Impressions','URL CTR (%)','Average Position','Competing Pages'];

queries.forEach((q,qi)=>{
  children.push(new Paragraph({
    children:[
      new TextRun({text:'🔴  ',size:24,font:'Arial'}),
      new TextRun({text:q.query,size:26,bold:true,font:'Arial',color:NAVY}),
      new TextRun({text:'  —  '+q.numPages+' pages · pos '+q.bestPos+' · '+q.totalImp.toLocaleString()+' impressions',size:20,font:'Arial',color:'5A7FA0'}),
    ],
    spacing:{before:180,after:80},
  }));

  const headerCells = COL_HEADERS.map((h,i)=>hdrCell(h,COL_WIDTHS[i]));
  const dataRows = q.rows.map(row=>new TableRow({children:[
    dataCell(row.slug,       COL_WIDTHS[0],row.isBest,false),
    dataCell(row.clicks,     COL_WIDTHS[1],row.isBest,true),
    dataCell(row.impressions,COL_WIDTHS[2],row.isBest,true),
    dataCell(row.ctr,        COL_WIDTHS[3],row.isBest,true),
    dataCell(row.position,   COL_WIDTHS[4],row.isBest,true),
    dataCell(row.competing,  COL_WIDTHS[5],row.isBest,true),
  ]}));

  children.push(new Table({
    width:{size:9360,type:WidthType.DXA}, columnWidths:COL_WIDTHS,
    rows:[new TableRow({children:headerCells}),...dataRows],
  }));

  const weakerText = q.weakerSlugs.slice(0,2).join(', ')+(q.weakerSlugs.length>2?` +${q.weakerSlugs.length-2} more`:'');
  children.push(new Paragraph({
    children:[
      new TextRun({text:'Suggested action: ',bold:true,size:18,font:'Arial',color:NAVY}),
      new TextRun({text:'Consolidate ',size:18,font:'Arial',color:'5A3000'}),
      new TextRun({text:weakerText,bold:true,size:18,font:'Courier New',color:MID_BLUE}),
      new TextRun({text:' into ',size:18,font:'Arial',color:'5A3000'}),
      new TextRun({text:q.bestSlug,bold:true,size:18,font:'Courier New',color:'1A6B3A'}),
      new TextRun({text:' (highest traffic authority) · use ',size:18,font:'Arial',color:'5A3000'}),
      new TextRun({text:'rel=canonical',bold:true,size:18,font:'Courier New',color:MID_BLUE}),
      new TextRun({text:' or 301 redirect on weaker pages · strengthen internal links to ',size:18,font:'Arial',color:'5A3000'}),
      new TextRun({text:q.bestSlug,bold:true,size:18,font:'Courier New',color:'1A6B3A'}),
      new TextRun({text:'.',size:18,font:'Arial',color:'5A3000'}),
    ],
    shading:{fill:'FFF8E1',type:ShadingType.CLEAR},
    border:{left:{style:BorderStyle.SINGLE,size:14,color:'F0A500',space:1}},
    indent:{left:160,right:160}, spacing:{before:80,after:80},
  }));

  if(qi<queries.length-1){ children.push(spacer(6)); children.push(sectionRule('D4DFE9')); }
});

// Footer
children.push(spacer(20));
children.push(new Paragraph({
  children:[new TextRun({text:'Generated by Edstellar Keyword Cannibalization Finder  ·  '+summary.date,size:16,font:'Arial',color:'8BA3BC',italics:true})],
  alignment:AlignmentType.CENTER,
  border:{top:{style:BorderStyle.SINGLE,size:2,color:'D4DFE9',space:4}},
  spacing:{before:200,after:0},
}));

const doc = new Document({
  styles:{ default:{ document:{ run:{ font:'Arial', size:20 } } } },
  sections:[{
    properties:{ page:{ size:{width:12240,height:15840}, margin:{top:720,right:900,bottom:900,left:900} } },
    children,
  }],
});

Packer.toBuffer(doc).then(buf=>{ fs.writeFileSync(process.argv[3],buf); console.log('OK '+buf.length+' bytes'); })
  .catch(e=>{ console.error(e.message); process.exit(1); });
"""

    with tempfile.TemporaryDirectory() as tmpdir:
        data_file = os.path.join(tmpdir, 'data.json')
        out_file  = os.path.join(tmpdir, 'report.docx')
        js_file   = os.path.join(tmpdir, 'gen.js')

        with open(data_file, 'w') as f:
            f.write(payload)
        with open(js_file, 'w') as f:
            f.write(js_script)

        result = subprocess.run(
            ['node', js_file, data_file, out_file],
            capture_output=True, text=True, timeout=30,
            env={**os.environ, 'NODE_PATH': '/home/claude/.npm-global/lib/node_modules'}
        )
        if result.returncode != 0:
            raise RuntimeError(f"Doc generation failed: {result.stderr}")

        with open(out_file, 'rb') as f:
            return f.read()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown('<div class="sidebar-section">Position Range</div>', unsafe_allow_html=True)
    pos_min = st.number_input("Minimum Position", min_value=1,   max_value=100, value=1)
    pos_max = st.number_input("Maximum Position", min_value=1,   max_value=100, value=20)

    st.markdown('<div class="sidebar-section">Volume Thresholds</div>', unsafe_allow_html=True)
    min_impressions = st.number_input("Minimum Impressions", min_value=0, max_value=1_000_000, value=0, step=50)
    min_clicks      = st.number_input("Minimum Clicks",      min_value=0, max_value=100_000,   value=0)
    min_pages       = st.number_input("Minimum Competing Pages", min_value=2, max_value=20,    value=2)

    st.markdown('<div class="sidebar-section">Smart Filters</div>', unsafe_allow_html=True)
    filter_anchors   = st.checkbox("Remove anchor (#) URLs",       value=True,
                                   help="Strips URL variants with #section anchors — these are the same page")
    filter_templates = st.checkbox("Remove geo-templated pages",   value=True,
                                   help="Excludes corporate-training-companies-<country>, skills-in-demand-in-<country>, <country>-work-culture, etc. These are intentionally different pages targeting different regions")

    st.markdown('<div class="sidebar-section">Display</div>', unsafe_allow_html=True)
    show_full_urls   = st.checkbox("Show full URLs",          value=False)
    group_by_query   = st.checkbox("Group results by query",  value=True)

    st.markdown('<div class="sidebar-section">Recommended Settings</div>', unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:0.78rem; line-height:1.8; opacity:0.85;">
    <b>For highest impact:</b><br>
    • Max Position → <b>10</b><br>
    • Min Impressions → <b>500</b><br>
    • Smart filters → <b>Both ON</b><br><br>
    <b>For full audit:</b><br>
    • Max Position → <b>20</b><br>
    • Min Impressions → <b>100</b>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
    <div class="badge">SEO Intelligence</div>
    <h1>🎯 Keyword Cannibalization Finder</h1>
    <p>Identify queries where multiple pages compete · Filter geo-templates · Export prioritized fix list</p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="section-hdr">Upload GSC Export</div>', unsafe_allow_html=True)

with st.expander("📋 How to export from Google Search Console", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        **Standard GSC Export (Queries + Pages):**
        1. Go to GSC → **Performance** → **Search results**
        2. Click **Pages** tab, then **+New** → filter by your site
        3. Switch back to **Queries** tab
        4. Set date range (90+ days recommended)
        5. Click **Export** → **Download CSV**

        The export must contain: `Query`, `Page/Landing Page`, `Clicks`, `Impressions`, `Position`
        """)
    with c2:
        st.markdown("""
        **Third-party / custom export formats accepted:**
        - Google Search Console API exports
        - Screaming Frog GSC integration
        - Semrush / Ahrefs GSC-linked exports
        - Custom exports with `Avg Position` column names

        **Supported file types:** `.csv`
        """)

uploaded_file = st.file_uploader(
    "Drop your GSC CSV here",
    type=["csv"],
    label_visibility="collapsed",
)

if uploaded_file is None:
    st.markdown("""
    <div class="info-box">
    👆 Upload a CSV export from Google Search Console to get started.
    The file must contain columns for <strong>Query</strong>, <strong>Page</strong>,
    <strong>Clicks</strong>, <strong>Impressions</strong>, and <strong>Position</strong>.
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-hdr">What this tool does</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="rec-card">
        <h4>🔍 Detects Cannibalization</h4>
        <p>Finds every query where 2+ of your pages are competing for the same keyword in search results.</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="rec-card">
        <h4>🌍 Smart Geo-Template Filter</h4>
        <p>Automatically excludes intentional geo-targeted page series (e.g. corporate-training-companies-singapore) so you only see real conflicts.</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="rec-card">
        <h4>📊 Prioritised Fix List</h4>
        <p>Ranks conflicts by Severity (High / Medium / Low) based on position and impressions so you know exactly where to act first.</p>
        </div>
        """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# LOAD & PREVIEW
# ══════════════════════════════════════════════════════════════════════════════

try:
    raw_df = pd.read_csv(uploaded_file, dtype=str)
    raw_df = read_gsc_data(raw_df)
except Exception as e:
    st.error(f"❌ Could not read file: {e}")
    st.stop()

st.success(f"✅ Loaded **{len(raw_df):,} rows** from `{uploaded_file.name}`")

with st.expander("👁 Preview raw data (first 20 rows)"):
    st.dataframe(raw_df.head(20), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSE BUTTON
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("")
run = st.button("🔍 Find Cannibalization Issues", type="primary", use_container_width=False)

if not run:
    st.markdown("""
    <div class="filter-note">
    ⚙️ Configure filters in the sidebar, then click <strong>Find Cannibalization Issues</strong> above.
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

with st.spinner("Analysing keyword cannibalization…"):
    filtered_df, audit = apply_filters(
        raw_df.copy(),
        pos_min=pos_min, pos_max=pos_max,
        min_impressions=min_impressions, min_clicks=min_clicks,
        filter_anchors=filter_anchors, filter_templates=filter_templates,
    )

    if filtered_df.empty:
        st.warning("No rows remain after applying filters. Try relaxing the position range or impression threshold.")
        st.stop()

    cannibs    = find_cannibalization(filtered_df, min_pages)
    query_sum  = build_query_summary(cannibs) if not cannibs.empty else pd.DataFrame()

if cannibs.empty:
    st.warning("No cannibalization issues found with the current filters. Try increasing Max Position or lowering Min Impressions.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="section-hdr">Analysis Results</div>', unsafe_allow_html=True)

# ── Filter audit strip ─────────────────────────────────────────────────────────
audit_parts = [f"**{audit['before']:,}** rows loaded"]
if audit['anchors_removed']:
    audit_parts.append(f"**{audit['anchors_removed']:,}** anchor-URL rows removed")
if audit['templates_removed']:
    audit_parts.append(f"**{audit['templates_removed']:,}** geo-template rows removed")
audit_parts.append(f"**{audit['after']:,}** rows analysed")

st.markdown(
    '<div class="filter-note">🔎 Filter log: ' + ' → '.join(audit_parts) + '</div>',
    unsafe_allow_html=True
)

# ── KPI cards ─────────────────────────────────────────────────────────────────
n_queries   = cannibs['query'].nunique()
n_pages     = len(cannibs)
total_impr  = int(cannibs['impressions'].sum())
total_clicks= int(cannibs['clicks'].sum())
avg_pages   = round(cannibs.groupby('query')['slug'].count().mean(), 1)
max_pages   = int(cannibs['competing_pages'].max())

query_sum['_sev'] = query_sum.apply(
    lambda r: severity(r['Best Average Position'], r['Impressions']), axis=1)
n_high   = len(query_sum[query_sum['_sev']=='High'])
n_medium = len(query_sum[query_sum['_sev']=='Medium'])
n_low    = len(query_sum[query_sum['_sev']=='Low'])

st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-card">
        <div class="kpi-label">Conflicting Queries</div>
        <div class="kpi-value">{n_queries:,}</div>
        <div class="kpi-sub">unique search terms</div>
    </div>
    <div class="kpi-card danger">
        <div class="kpi-label">High Severity</div>
        <div class="kpi-value">{n_high:,}</div>
        <div class="kpi-sub">pos ≤10 · impr ≥1K</div>
    </div>
    <div class="kpi-card accent">
        <div class="kpi-label">Medium Severity</div>
        <div class="kpi-value">{n_medium:,}</div>
        <div class="kpi-sub">pos ≤20 · impr ≥200</div>
    </div>
    <div class="kpi-card success">
        <div class="kpi-label">Low Severity</div>
        <div class="kpi-value">{n_low:,}</div>
        <div class="kpi-sub">lower priority</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Impressions at Stake</div>
        <div class="kpi-value">{total_impr:,}</div>
        <div class="kpi-sub">across all conflicts</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Clicks at Stake</div>
        <div class="kpi-value">{total_clicks:,}</div>
        <div class="kpi-sub">across all conflicts</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Avg URLs / Query</div>
        <div class="kpi-value">{avg_pages}</div>
        <div class="kpi-sub">max: {max_pages}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Query Summary",
    "🔍 Detail View",
    "🔴 High Severity",
    "💡 Recommendations",
])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1: Query Summary
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.markdown("#### One row per query — all competing slugs listed inline")

    display_qs = query_sum.copy()
    display_qs.insert(2, 'Severity',
        display_qs.apply(lambda r: severity(r['Best Average Position'], r['Impressions']), axis=1))

    if not show_full_urls:
        display_qs['Best Landing Page'] = display_qs['Best Landing Page'].str[:60]
        display_qs['All Landing Pages'] = display_qs['All Landing Pages'].str[:120]

    st.dataframe(display_qs.drop(columns=['_sev'], errors='ignore'),
                 use_container_width=True, hide_index=True)

    # Build detail export — referenced by Excel download button
    detail_export = cannibs.rename(columns={
        'query': 'Query', 'slug': 'Landing Page',
        'clicks': 'Url Clicks', 'impressions': 'Impressions',
        'ctr': 'URL CTR (%)', 'position': 'Average Position',
        'competing_pages': 'Competing Pages',
    })

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button("📥 Download CSV",
            data=to_csv(display_qs.drop(columns=['_sev'], errors='ignore')),
            file_name="cannibalization_query_summary.csv", mime="text/csv")
    with dl2:
        st.download_button("📥 Download Excel",
            data=to_excel({
                'Query Summary': display_qs.drop(columns=['_sev'], errors='ignore'),
                'Detail View': detail_export,
            }),
            file_name="cannibalization_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2: Detail View
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.markdown("#### Every query × URL slug pair — sortable and filterable")

    detail_display = cannibs.rename(columns={
        'query': 'Query', 'slug': 'Landing Page',
        'clicks': 'Url Clicks', 'impressions': 'Impressions',
        'ctr': 'URL CTR (%)', 'position': 'Average Position',
        'competing_pages': 'Competing Pages',
    }).copy()

    # Severity column
    detail_display.insert(2, 'Severity',
        detail_display.apply(lambda r: severity(r['Average Position'], r['Impressions']), axis=1))

    if not show_full_urls:
        detail_display['Landing Page'] = detail_display['Landing Page'].str[:70]

    st.dataframe(detail_display, use_container_width=True, hide_index=True)
    st.download_button("📥 Download Detail CSV",
        data=to_csv(detail_display),
        file_name="cannibalization_detail.csv", mime="text/csv")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3: High Severity
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    high_queries = query_sum[query_sum['_sev'] == 'High']['Query'].tolist()
    if not high_queries:
        st.info("No High Severity conflicts with the current filters. Try setting Max Position to 10 and Min Impressions to 500.")
    else:
        st.markdown(f"#### {len(high_queries)} queries · Best position ≤10 · Impressions ≥1,000")
        st.markdown("""
        <div class="filter-note">
        🚨 These are your most urgent fixes — you're already ranking on page 1 but splitting click potential across multiple URLs.
        Consolidating these will have the most direct impact on organic clicks.
        </div>
        """, unsafe_allow_html=True)

        high_detail = cannibs[cannibs['query'].isin(high_queries)].copy()
        high_detail_display = high_detail.rename(columns={
            'query': 'Query', 'slug': 'Landing Page',
            'clicks': 'Url Clicks', 'impressions': 'Impressions',
            'ctr': 'URL CTR (%)', 'position': 'Average Position',
            'competing_pages': 'Competing Pages',
        })

        # Expandable per query
        for q in high_queries[:30]:
            qdata = cannibs[cannibs['query'] == q].copy()
            # Score each page: impressions + clicks*10 — highest score = canonical winner
            qdata['_score'] = qdata['impressions'] + (qdata['clicks'] * 10)
            qdata_display   = qdata.sort_values('_score', ascending=False)
            best_pos        = qdata['position'].min()
            total_impr      = int(qdata['impressions'].sum())
            # The winning page is the one with most traffic authority
            best_slug       = qdata_display.iloc[0]['slug']
            weaker_slugs    = qdata_display.iloc[1:]['slug'].tolist()

            with st.expander(
                f"🔴  **{q}**  —  {len(qdata)} pages · pos {best_pos} · {total_impr:,} impressions"
            ):
                disp = qdata_display.rename(columns={
                    'slug': 'Landing Page', 'clicks': 'Url Clicks',
                    'impressions': 'Impressions', 'ctr': 'URL CTR (%)',
                    'position': 'Average Position', 'competing_pages': 'Competing Pages',
                })[['Landing Page', 'Url Clicks', 'Impressions', 'URL CTR (%)', 'Average Position', 'Competing Pages']]
                st.dataframe(disp, use_container_width=True, hide_index=True)

                weaker_preview = ', '.join(f'`{s}`' for s in weaker_slugs[:2])
                if len(weaker_slugs) > 2:
                    weaker_preview += f' +{len(weaker_slugs)-2} more'
                st.markdown(
                    f"**Suggested action:** Consolidate {weaker_preview} **into** "
                    f"`{best_slug}` (highest traffic authority) · use `rel=canonical` "
                    f"or 301 redirect on the weaker pages · then strengthen internal "
                    f"links to `{best_slug}`."
                )

        # ── Download buttons ──────────────────────────────────────────────
        dl_c1, dl_c2 = st.columns(2)
        with dl_c1:
            st.download_button("📥 Download CSV",
                data=to_csv(high_detail_display),
                file_name="cannibalization_high_severity.csv", mime="text/csv")
        with dl_c2:
            try:
                docx_bytes = generate_high_severity_docx(cannibs, query_sum)
                st.download_button(
                    "📄 Download Word Report (.docx)",
                    data=docx_bytes,
                    file_name="high_severity_cannibalization_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.warning(f"Word export unavailable: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 4: Recommendations
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    st.markdown("#### How to fix keyword cannibalization")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        <div class="rec-card">
        <h4>1. Consolidate (Merge)</h4>
        <p>Merge the weaker competing pages into the strongest one. Combine their content, then 301 redirect the old URLs to the winner. Best when pages cover the same intent.</p>
        </div>
        <div class="rec-card">
        <h4>2. Canonical Tags</h4>
        <p>Add <code>rel="canonical"</code> on weaker pages pointing to the primary URL. Fast to implement and signals to Google which page should rank.</p>
        </div>
        <div class="rec-card">
        <h4>3. Differentiate Content</h4>
        <p>If pages serve slightly different intents (informational vs transactional), update them to clearly target different keywords. Avoid keyword overlap in titles and H1s.</p>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="rec-card">
        <h4>4. Internal Linking</h4>
        <p>Use internal links to signal which page is the authority. Link to the primary page using exact-match anchor text from related posts.</p>
        </div>
        <div class="rec-card">
        <h4>5. 301 Redirect Weaker Pages</h4>
        <p>For pages with near-zero clicks and impressions, a direct 301 redirect to the primary page consolidates all link equity with no content work needed.</p>
        </div>
        <div class="rec-card">
        <h4>6. Update Title Tags & Meta</h4>
        <p>Ensure no two pages share the same or near-duplicate title tags. Differentiate each page's primary keyword in the title to reduce signal overlap.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Priority matrix for this dataset")

    priority_df = query_sum[['Query', 'Competing Pages', 'Impressions',
                              'Url Clicks', 'Best Average Position', '_sev']]\
                      .rename(columns={'_sev': 'Severity'})\
                      .sort_values('Impressions', ascending=False)\
                      .head(50)

    priority_df['Recommended Action'] = priority_df.apply(lambda r: (
        'Consolidate / 301 redirect'     if r['Severity'] == 'High'   else
        'Add canonicals / differentiate' if r['Severity'] == 'Medium' else
        'Monitor / internal linking'
    ), axis=1)

    st.dataframe(priority_df, use_container_width=True, hide_index=True)
    st.download_button("📥 Download Priority Matrix CSV",
        data=to_csv(priority_df),
        file_name="cannibalization_priority_matrix.csv", mime="text/csv")


# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center; font-size:0.78rem; color:rgba(255,255,255,.4); padding:8px 0 20px;">
    Built for Edstellar SEO team · Based on the Keyword Cannibalization Finder by 
    <a href="https://www.leefoot.com" target="_blank" style="color:#4A9FD5;">Lee Foot</a> ·
    Smart geo-template filtering for corporate training page series
</div>
""", unsafe_allow_html=True)
