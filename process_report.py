import pandas as pd
import os
import sys
import re
import html as _html
import io
from datetime import datetime

_LOGO_B64 = "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"

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalize(name):
    return re.sub(r'[\s_\-]+', '', str(name)).lower()

def find_client_config(filename, config_df):
    base = normalize(os.path.basename(filename))
    for _, row in config_df.iterrows():
        if normalize(row['Client Name']) in base:
            return row
    raise ValueError(
        f"No client in client_config.xlsx matched '{os.path.basename(filename)}'.\n"
        f"Clients available: {list(config_df['Client Name'])}"
    )

def find_col_by_key(csv_columns, config_row, key):
    val = config_row.get(key, '')
    if pd.notna(val) and str(val).strip():
        search = str(val).strip().lower()
        for c in csv_columns:
            if search in c.lower():
                return c
    return None

def find_conv_cols(csv_columns, config_row):
    cols = []
    for key in ['Conversion Column 1', 'Conversion Column 2',
                'Conversion Column 3', 'Conversion Column 4',
                'Conversion Column 5']:
        col = find_col_by_key(csv_columns, config_row, key)
        if col and col not in cols:
            cols.append(col)
    return cols

def short_name(col):
    return re.sub(r'\s*\[.*$', '', col).strip()

CHANNEL_KEYWORDS = ['CTV', 'Video', 'Display', 'Audio']

def extract_channel(campaign_name):
    upper = str(campaign_name).upper()
    for ch in CHANNEL_KEYWORDS:
        if ch.upper() in upper:
            return ch
    return str(campaign_name)

def calc_metrics(imp, clk, spnd, conv, pcv):
    ctr  = clk  / imp  * 100  if imp  else 0
    cpc  = spnd / clk         if clk  else 0
    cpm  = spnd / imp  * 1000 if imp  else 0
    cvr  = conv / imp  * 100  if imp  else 0
    comp = pcv  / imp  * 100  if imp  else 0
    return ctr, cpc, cpm, cvr, comp

def roas(rev, spnd):
    return rev / spnd if spnd else 0

def _extract_dates(filename):
    return re.findall(r'\d{4}-\d{2}-\d{2}', os.path.basename(filename))

def extract_date_range(filename):
    dates = _extract_dates(filename)
    if len(dates) >= 2:
        fmt = lambda d: datetime.strptime(d, '%Y-%m-%d').strftime('%-d %B %Y')
        return f"{fmt(dates[0])} – {fmt(dates[1])}"
    return ''

def extract_report_month(filename):
    dates = _extract_dates(filename)
    if dates:
        return datetime.strptime(dates[0], '%Y-%m-%d').strftime('%B %Y')
    return ''

def extract_adgroup_label(ag_full):
    """Return the last underscore-segment of a Trade Desk Ad Group name.
    e.g. 'IO9711_The Meat Box_Omnichannel_Audio_Rova' → 'Rova'
         'IO9711_..._Display_Females 35-64'           → 'Females 35-64'
    """
    return str(ag_full).split('_')[-1].strip()

def _prev_month_label(report_month):
    try:
        dt = datetime.strptime(report_month, '%B %Y')
        prev = dt.replace(year=dt.year - 1, month=12) if dt.month == 1 \
               else dt.replace(month=dt.month - 1)
        return prev.strftime('%B %Y')
    except Exception:
        return 'Previous Month'

def _calc_upsell(client_history, report_month, curr_spnd, curr_conv):
    if client_history is None or client_history.empty or curr_conv <= 0 or curr_spnd <= 0:
        return None
    try:
        curr_dt = datetime.strptime(report_month, '%B %Y')

        def _nth_prev(n):
            m, y = curr_dt.month - n, curr_dt.year
            while m <= 0:
                m += 12; y -= 1
            return datetime(y, m, 1).strftime('%B %Y')

        prev_months = [_nth_prev(i) for i in range(1, 4)]

        for m in prev_months[:2]:
            rows = client_history[client_history['Month'] == m]
            if not rows.empty:
                val = rows.iloc[0].get('Upsell_Triggered', '')
                if val is True or str(val).upper() == 'TRUE':
                    return None

        benchmark_cpas = []
        for m in prev_months:
            rows = client_history[client_history['Month'] == m]
            if not rows.empty:
                r = rows.iloc[0]
                s = float(r.get('Spend', 0) or 0)
                c = float(r.get('Conversions', 0) or 0)
                if s > 0 and c > 0:
                    benchmark_cpas.append(s / c)

        if len(benchmark_cpas) < 2:
            return None

        avg_cpa  = sum(benchmark_cpas) / len(benchmark_cpas)
        curr_cpa = curr_spnd / curr_conv

        if avg_cpa <= 0:
            return None

        improvement_pct = (avg_cpa - curr_cpa) / avg_cpa * 100
        if improvement_pct < 15.0:
            return None

        return {
            'curr_cpa':        curr_cpa,
            'avg_cpa':         avg_cpa,
            'improvement_pct': improvement_pct,
            'budget_rec':      curr_spnd * 0.20,
        }
    except Exception:
        return None

def _h(s):   return _html.escape(str(s))
def _n(v):   return f"{v:,.0f}"
def _m(v):   return f"${v:,.2f}"
def _p(v):   return f"{v:.2f}%"
def _rx(v):  return f"{v:.2f}x"

_KPI_CYCLE = ['#EF426F', '#5BC2E7']

# If total campaign conversions fall below this threshold, swap CPA for
# Cost Per Visit (CPV = Spend ÷ Site Traffic) wherever CPA would appear.
CONV_THRESHOLD = 20

def _delta_badge(curr_val, prev_val, invert=False):
    if not prev_val:
        return ''
    pct = (curr_val - prev_val) / abs(prev_val) * 100
    if abs(pct) < 0.05:
        return '<span style="color:#6B7280;font-size:12px;margin-left:5px;font-weight:600">no change</span>'
    is_good = pct > 0 if not invert else pct < 0
    color = '#5BC2E7' if is_good else '#EF426F'
    arrow = '▲' if pct > 0 else '▼'
    return f'<span style="color:{color};font-size:12px;margin-left:5px;font-weight:600">{arrow}&nbsp;{abs(pct):.1f}%</span>'

def _avg_badge(value, benchmark):
    color = '#5BC2E7' if value >= benchmark else '#EF426F'
    arrow = '▲' if value >= benchmark else '▼'
    return f'<span style="color:{color};font-size:13px;margin-left:4px">{arrow}</span>'

def _kpi_card(label, value, color='#EF426F', trend=None, invert=False):
    trend_html = ''
    if trend is not None:
        if abs(trend) < 0.05:
            trend_html = '<div class="kpi-trend" style="color:#6B7280">no change vs last month</div>'
        else:
            is_good = trend > 0 if not invert else trend < 0
            t_color = '#5BC2E7' if is_good else '#EF426F'
            arrow = '▲' if trend > 0 else '▼'
            trend_html = f'<div class="kpi-trend" style="color:{t_color}">{arrow} {abs(trend):.1f}% vs last month</div>'
    return f'''
        <div class="kpi-card" style="border:2px solid {color}">
          <div class="kpi-value">{value}</div>
          <div class="kpi-label">{label}</div>
          {trend_html}
        </div>'''

def _th(cells, right_from=1):
    ths = ''
    for i, c in enumerate(cells):
        cls = ' class="right"' if i >= right_from else ''
        ths += f'<th{cls}>{_h(c)}</th>'
    return f'<tr>{ths}</tr>'

def _td_row(cells, right_from=1):
    tds = ''
    for i, c in enumerate(cells):
        align = ' class="right"' if i >= right_from else ''
        safe = c if (isinstance(c, str) and c.startswith('<')) else _h(c)
        tds += f'<td{align}>{safe}</td>'
    return f'<tr>{tds}</tr>'


_BB_COLORS = ['#5BC2E7', '#D946EF', '#7C3AED', '#EF426F', '#9BA3AF']

_BAR_COLOR_KEY = (
    '<span class="bar-key-item"><span class="bar-key-dot" style="background:#22D3EE"></span>Above average</span>'
    '<span class="bar-key-sep">&middot;</span>'
    '<span class="bar-key-item"><span class="bar-key-dot" style="background:#FBBF24"></span>Within 15% of average</span>'
    '<span class="bar-key-sep">&middot;</span>'
    '<span class="bar-key-item"><span class="bar-key-dot" style="background:#EF426F"></span>Below average</span>'
)

# Inline SVG icons (Lucide-style)
_ICON_LIGHTBULB = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
    ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M15 14c.2-1 .7-1.7 1.5-2.5 1-.9 1.5-2.2 1.5-3.5A6 6 0 0 0 6 8c0 1 .2 2.2 1.5 3.5'
    '.7.7 1.3 1.5 1.5 2.5"/><path d="M9 18h6"/><path d="M10 22h4"/></svg>'
)
_ICON_TARGET = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
    ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/>'
    '<circle cx="12" cy="12" r="2"/></svg>'
)
_ICON_TREND_UP = (
    '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
    ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/>'
    '<polyline points="17 6 23 6 23 12"/></svg>'
)
_ICON_COG = (
    '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
    ' stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="3"/>'
    '<path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06'
    'a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09'
    'A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83'
    'l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09'
    'A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83'
    'l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09'
    'a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83'
    'l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09'
    'a1.65 1.65 0 0 0-1.51 1z"/></svg>'
)

def _creative_type_dot(name):
    """Coloured dot based on creative file type."""
    n = str(name).lower()
    if any(n.endswith(e) for e in ('.mp4', '.mov', '.avi', '.webm')):
        c, t = '#5BC2E7', 'Video'
    elif n.endswith('.gif'):
        c, t = '#D946EF', 'Animated GIF'
    elif any(n.endswith(e) for e in ('.jpg', '.jpeg', '.png', '.webp', '.svg')):
        c, t = '#9BA3AF', 'Static Image'
    else:
        c, t = '#7C3AED', 'Rich/Other'
    return f'<span class="type-dot" style="background:{c}" title="{t}"></span>'

def _metric_card_html(label, value, trend_html='', subtext=''):
    sub   = f'<div class="mc-sub">{subtext}</div>' if subtext else ''
    trend = f'<div class="mc-trend">{trend_html}</div>' if trend_html else ''
    return (f'<div class="metric-card">'
            f'<div class="mc-label">{label}</div>'
            f'<div class="mc-value">{value}</div>'
            f'{trend}{sub}</div>')

def _stacked_budget_bar_html(segments):
    """segments = [(label, pct, color), ...]  — tall bar with labels inside each segment."""
    if not segments:
        return ''
    bars = ''
    for l, p, c in segments:
        inner = (f'<span class="sbb-inner-label">{_h(l)} {p:.0f}%</span>'
                 if p >= 11 else '')
        bars += (f'<div class="sbb-seg" style="width:{p:.1f}%;background:{c}"'
                 f' title="{_h(l)}: {p:.0f}%">{inner}</div>')
    legend_items = ''.join(
        f'<span class="sbb-legend-item">'
        f'<span class="sbb-legend-dot" style="background:{c}"></span>'
        f'{_h(l)} <strong>{p:.0f}%</strong>'
        f'</span>'
        for l, p, c in segments
    )
    legend = f'<div class="sbb-legend">{legend_items}</div>'
    return f'<div class="sbb-track">{bars}</div>{legend}'

def _index_bar_html(value_str, val, avg_val, max_val, is_inverse=False):
    if max_val <= 0:
        return f'<span>{value_str}</span>'
    pct = min(val / max_val * 100, 100)
    color = '#22D3EE'
    avg_marker = ''
    if avg_val and avg_val > 0:
        avg_pct  = min(avg_val / max_val * 100, 100)
        is_good  = val <= avg_val if is_inverse else val >= avg_val
        is_warn  = (not is_good) and (val >= avg_val * 0.85 if not is_inverse else val <= avg_val * 1.15)
        color    = '#22D3EE' if is_good else ('#FBBF24' if is_warn else '#EF426F')
        # Marker extends above and below the bar (-4px top, 16px height, 2px wide)
        avg_marker = (f'<div style="position:absolute;left:{avg_pct:.1f}%;top:-4px;'
                      f'height:16px;width:2px;background:#fff;z-index:10;'
                      f'box-shadow:0 0 3px rgba(0,0,0,0.5);border-radius:1px"></div>')
    return (f'<div class="ibar-cell">'
            f'<span class="ibar-val" style="color:{color}">{value_str}</span>'
            f'<div class="ibar-track">'
            f'<div class="ibar-fill" style="width:{pct:.1f}%;background:{color}"></div>'
            f'{avg_marker}</div></div>')

def _insight_box_html(title, body, variant='cyan'):
    if variant == 'rose':
        color, bg, border, icon = '#EF426F', 'rgba(239,68,68,0.10)', 'rgba(239,68,68,0.30)', _ICON_TARGET
    elif variant == 'green':
        color, bg, border, icon = '#10B981', 'rgba(16,185,129,0.10)', 'rgba(16,185,129,0.30)', _ICON_TREND_UP
    else:
        color, bg, border, icon = '#22D3EE', 'rgba(6,182,212,0.10)', 'rgba(6,182,212,0.30)', _ICON_LIGHTBULB
    return (f'<div class="insight-v2" style="background:{bg};border-color:{border}">'
            f'<div class="iv2-bar" style="background:{color}"></div>'
            f'<div class="iv2-inner">'
            f'<div class="iv2-header">'
            f'<span class="iv2-icon" style="color:{color}">{icon}</span>'
            f'<h4 class="iv2-title" style="color:{color}">{title}</h4>'
            f'</div>'
            f'<p class="iv2-body">{body}</p>'
            f'</div></div>')

def _mom_sparkline_svg(months, convs, curr_month, curr_conv, height=140, label='Conversions'):
    """Returns SVG sparkline string. Height is dynamic so the chart fills its container.
    Returns '' if fewer than 2 data points total."""
    all_months = list(months) + [curr_month]
    all_convs  = list(convs)  + [float(curr_conv)]
    if len(all_months) < 2:
        return ''
    if len(all_months) > 6:
        all_months = all_months[-6:]
        all_convs  = all_convs[-6:]
    n = len(all_months)
    W = 280
    H = max(height, 80)
    PAD_L, PAD_R, PAD_T, PAD_B = 34, 6, 10, 22
    area_w   = W - PAD_L - PAD_R
    area_h   = H - PAD_T - PAD_B
    bar_slot = area_w / n
    bar_w    = bar_slot * 0.58
    max_conv = max(all_convs) or 1

    def _fmt(v):
        if v >= 10000: return f'{v/1000:.0f}K'
        if v >= 1000:  return f'{v/1000:.1f}K'
        return str(int(round(v)))

    # Faint horizontal grid lines + y-axis labels at 0%, 50%, 100%
    grid = ''
    for frac in (0.0, 0.5, 1.0):
        y   = PAD_T + area_h - frac * area_h
        val = frac * max_conv
        grid += (
            f'<line x1="{PAD_L}" y1="{y:.1f}" x2="{W - PAD_R}" y2="{y:.1f}"'
            f' stroke="rgba(255,255,255,0.07)" stroke-width="1"/>'
        )
        grid += (
            f'<text x="{PAD_L - 4}" y="{y + 3.5:.1f}" text-anchor="end"'
            f' font-size="10" fill="#4B5563" font-family="sans-serif">{_fmt(val)}</text>'
        )

    bars = ''
    labels = ''
    for i, (m, cv) in enumerate(zip(all_months, all_convs)):
        bx = PAD_L + i * bar_slot + (bar_slot - bar_w) / 2
        bh = max((cv / max_conv) * area_h, 2.0)
        by = PAD_T + area_h - bh
        bars += (
            f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bar_w:.1f}"'
            f' height="{bh:.1f}" rx="2" fill="#5BC2E7"/>'
        )
        cx  = PAD_L + i * bar_slot + bar_slot / 2
        lbl = m.split()[0][:3] if m else ''
        labels += (
            f'<text x="{cx:.1f}" y="{H - 5}" text-anchor="middle"'
            f' font-size="11" fill="#4B5563" font-family="sans-serif">{_h(lbl)}</text>'
        )

    return (
        f'<svg viewBox="0 0 {W} {H}" preserveAspectRatio="none"'
        f' style="width:100%;height:100%;display:block;min-height:80px">'
        f'{grid}{bars}{labels}</svg>'
    )


def _h_th(label, right=False, raw=False):
    """Generate a <th> cell; raw=True skips escaping (allows HTML sub-labels)."""
    align = ' class="right"' if right else ''
    inner = label if raw else _h(label)
    return f'<th{align}>{inner}</th>'


# ── Ad Group Performance helpers ──────────────────────────────────────────────
_VISIT_RATE_CHANNELS_AGP = {'CTV', 'Audio', 'Video'}
MIN_AGP_IMP = 1_000  # hide ad groups with fewer impressions than this

def _agp_bar_row(label, val_str, val, avg_val, is_inverse=False):
    """One row in the ad-group performance chart.
    avg_val always maps to 50% bar width so the avg line is centred.
    is_inverse=True  → lower value is better (CPA).
    is_inverse=False → higher value is better (Visit Rate).
    """
    bar_pct = min((val / avg_val) * 50.0, 100.0) if avg_val > 0 else 50.0
    is_good = (val <= avg_val) if is_inverse else (val >= avg_val)
    is_warn = (not is_good) and (val >= avg_val * 0.85 if not is_inverse else val <= avg_val * 1.15)
    color   = '#22D3EE' if is_good else ('#FBBF24' if is_warn else '#EF426F')
    return (
        f'<div class="agp-row">'
        f'<div class="agp-row-left">'
        f'<span class="agp-ag-name">{_h(label)}</span>'
        f'<span class="agp-val" style="color:{color}">{val_str}</span>'
        f'</div>'
        f'<div class="agp-track">'
        f'<div class="agp-fill" style="width:{bar_pct:.1f}%;background:{color}"></div>'
        f'<div class="agp-avg-line"></div>'
        f'</div>'
        f'</div>'
    )

def _agp_card_html(chan, chan_rows, use_cpa):
    """Build one channel card. Returns '' if fewer than 2 valid rows."""
    is_inverse = use_cpa  # CPA: lower better; Visit Rate: higher better

    if use_cpa:
        valid = chan_rows[(chan_rows['conv'] > 0) & (chan_rows['imp'] >= MIN_AGP_IMP)].copy()
        if len(valid) < 2:
            return ''
        valid['_val']     = valid['spnd'] / valid['conv']
        valid['_val_str'] = valid['_val'].apply(_m)
        avg_val    = valid['spnd'].sum() / valid['conv'].sum()
        avg_str    = _m(avg_val)
        metric_lbl = 'Cost per Conv (CPA)'
        valid = valid.sort_values('_val', ascending=True)   # lowest CPA first = best
    else:
        valid = chan_rows[(chan_rows['imp'] >= MIN_AGP_IMP) & (chan_rows['st'] > 0)].copy()
        if len(valid) < 2:
            return ''
        valid['_val']     = valid['st'] / valid['imp'] * 100
        valid['_val_str'] = valid['_val'].apply(_p)
        avg_val    = valid['st'].sum() / valid['imp'].sum() * 100
        avg_str    = _p(avg_val)
        metric_lbl = 'Visit Rate'
        valid = valid.sort_values('_val', ascending=False)  # highest VR first = best

    valid = valid.head(4)

    rows_html = ''.join(
        _agp_bar_row(r['_ag_label'], r['_val_str'], r['_val'], avg_val, is_inverse=is_inverse)
        for _, r in valid.iterrows()
    )

    # Only highlight performers that are genuinely above average (cyan = positive)
    above_avg = valid[valid['_val'] <= avg_val] if is_inverse else valid[valid['_val'] >= avg_val]
    top_labels = above_avg.iloc[:2]['_ag_label'].tolist()
    if top_labels:
        label_html = ' &amp; '.join(f'<strong>{_h(l)}</strong>' for l in top_labels)
        footer_html = (
            f'<div class="agp-card-footer">'
            f'<span class="agp-footer-icon">{_ICON_COG}</span>'
            f'<span class="agp-footer-text">Budget optimised toward {label_html}</span>'
            f'</div>'
        )
    else:
        footer_html = ''

    return (
        f'<div class="agp-card">'
        f'<div class="agp-card-header">'
        f'<span class="agp-chan-name">{_h(chan)}</span>'
        f'<div class="agp-metric-info">'
        f'<span class="agp-metric-name">{_h(metric_lbl)}</span>'
        f'<span class="agp-avg-chip">'
        f'<span class="agp-avg-tick"></span> Blended Avg {_h(avg_str)}'
        f'</span>'
        f'</div>'
        f'</div>'
        f'<div class="agp-rows">{rows_html}</div>'
        f'{footer_html}'
        f'</div>'
    )


def _agp_remarketing_hub_html(remark_by_chan, display_total_conv, prosp_by_chan=None):
    """Build the Remarketing Hub summary card from {chan: rows_df}.
    Shows one aggregated row per channel with a % vs prospecting-average badge.
    Returns '' if no remarketing data exists."""
    if not remark_by_chan:
        return ''

    RM_COLOR = '#60A5FA'   # blue-400 — calm, less intense than the previous purple
    RM_MUTED = '#93C5FD'   # blue-300
    rows_html = ''

    for chan, rows in remark_by_chan.items():
        use_cpa   = False
        prosp     = (prosp_by_chan or {}).get(chan, pd.DataFrame())

        if use_cpa:
            total_spnd = rows['spnd'].sum()
            total_conv = rows['conv'].sum()
            val        = total_spnd / total_conv if total_conv > 0 else None
            val_str    = _m(val) if val is not None else 'N/A'
            metric_lbl = 'Cost per Conversion (CPA)'
            p_spnd     = prosp['spnd'].sum() if not prosp.empty else 0
            p_conv     = prosp['conv'].sum() if not prosp.empty else 0
            prosp_avg  = p_spnd / p_conv if p_conv > 0 else None
        else:
            total_imp  = rows['imp'].sum()
            total_st   = rows['st'].sum()
            val        = total_st / total_imp * 100 if total_imp > 0 else None
            val_str    = _p(val) if val is not None else 'N/A'
            metric_lbl = 'Visit Rate'
            p_imp      = prosp['imp'].sum() if not prosp.empty else 0
            p_st       = prosp['st'].sum()  if not prosp.empty else 0
            prosp_avg  = p_st / p_imp * 100 if p_imp > 0 else None

        # % vs prospecting-average badge
        vs_html = ''
        if val is not None and prosp_avg and prosp_avg > 0:
            pct      = (val - prosp_avg) / prosp_avg * 100
            is_good  = (pct < 0) if use_cpa else (pct > 0)   # lower CPA = good, higher VR = good
            vs_color = '#22D3EE' if is_good else '#EF426F'
            arrow    = '▲' if pct > 0 else '▼'
            label    = 'below' if pct < 0 else 'above'
            vs_html  = (f'<span class="agp-rm-vs" style="color:{vs_color}">'
                        f'{arrow} {abs(pct):.0f}% {label} {_h(chan)} avg</span>')

        rows_html += (
            f'<div class="agp-rm-row">'
            f'<div class="agp-rm-left">'
            f'<span class="agp-rm-chan">{_h(chan)}</span>'
            f'<span class="agp-rm-lbl">{_h(metric_lbl)}</span>'
            f'</div>'
            f'<div class="agp-rm-right">'
            f'<span class="agp-rm-val" style="color:{RM_COLOR}">{val_str}</span>'
            f'{vs_html}'
            f'</div>'
            f'</div>'
        )

    return (
        f'<div class="agp-card agp-card-rm">'
        f'<div class="agp-card-header" style="border-bottom-color:rgba(96,165,250,0.25)">'
        f'<span class="agp-chan-name" style="color:{RM_COLOR}">Remarketing</span>'
        f'<div class="agp-metric-info">'
        f'<span class="agp-metric-name" style="color:{RM_MUTED}">Bottom-of-Funnel Audiences</span>'
        f'</div>'
        f'</div>'
        f'<div class="agp-rows agp-rm-rows">{rows_html}</div>'
        f'</div>'
    )


def generate_html(csv_path, client_name, conv_label, has_revenue,
                  totals, grp_chan, grp_cre, grp_site, prev_data=None, upsell_data=None,
                  client_history=None, grp_adgroup=None):
    imp, clk, spnd, conv, st, pcv, rev, uh = (*totals, 0) if len(totals) == 7 else totals
    _, cpc, cpm, cvr, _ = calc_metrics(imp, clk, spnd, conv, pcv)
    report_month = extract_report_month(csv_path)
    date_range   = extract_date_range(csv_path)

    NA     = '<span class="na">N/A</span>'
    NA_mom = '<span class="na">N/A</span>'

    if prev_data:
        pi    = float(prev_data.get('Impressions', 0) or 0)
        pk    = float(prev_data.get('Clicks', 0) or 0)
        ps    = float(prev_data.get('Spend', 0) or 0)
        pc    = float(prev_data.get('Conversions', 0) or 0)
        pr    = float(prev_data.get('Revenue', 0) or 0)
        p_st  = float(prev_data.get('Site Traffic', 0) or 0)
        prev_cpm  = ps / pi * 1000 if pi else 0
        prev_roas = roas(pr, ps)
    else:
        pi = pk = ps = pc = pr = p_st = prev_cpm = prev_roas = 0

    def _t(curr, prev_val):
        if not prev_data or not prev_val:
            return None
        return (curr - prev_val) / abs(prev_val) * 100

    def _trend_line(curr, prev_val, invert=False):
        t = _t(curr, prev_val)
        if t is None or abs(t) < 0.05:
            return ''
        is_good = t > 0 if not invert else t < 0
        color   = '#5BC2E7' if is_good else '#EF426F'
        arrow   = '&#9650;' if t > 0 else '&#9660;'
        return f'<div class="trend-line" style="color:{color}">{arrow} {abs(t):.1f}% vs last month</div>'

    cpa_kpi      = spnd / conv if conv else 0
    prev_cpa_kpi = ps   / pc   if pc   else 0

    # Low-conversion fallback: if total conversions < CONV_THRESHOLD, swap CPA
    # for Cost Per Visit (CPV = Spend ÷ Site Traffic) in the overview card,
    # MoM table, and sparkline.
    low_conv = (conv < CONV_THRESHOLD)
    cpv      = spnd / st if st > 0 else 0
    prev_cpv = ps   / p_st if p_st > 0 else 0

    # ── MoM table ─────────────────────────────────────────────────────────────
    def _mc(main, curr, prev, invert=False):
        badge = _delta_badge(curr, prev, invert=invert) if prev is not None else ''
        return f'<span>{main}{badge}</span>'

    history_rows_html = ''
    mom_insight_text  = ''
    has_mom           = False
    mom_table         = ''
    sp_months, sp_convs, sp_cpas, sp_sts = [], [], [], []

    if client_history is not None and not client_history.empty:
        def _month_dt(m):
            try:
                return datetime.strptime(m, '%B %Y')
            except Exception:
                return datetime.min

        hist = client_history.copy()
        hist = hist[hist['Month'] != report_month]
        hist['_dt'] = hist['Month'].apply(_month_dt)
        hist = hist.sort_values('_dt')

        if not hist.empty and conv > 0 and spnd > 0:
            first_r = hist.iloc[0]
            first_s = float(first_r.get('Spend', 0) or 0)
            first_c = float(first_r.get('Conversions', 0) or 0)
            if first_s > 0 and first_c > 0:
                first_cpa    = first_s / first_c
                curr_cpa_ins = spnd / conv
                improvement  = (first_cpa - curr_cpa_ins) / first_cpa * 100
                if improvement >= 5:
                    mom_insight_text = f'Since launch, your cost-per-conversion has improved by <strong>{improvement:.0f}%</strong>, reflecting an increasingly optimised path to conversion.'

        prev_spend = prev_cpa_hist = prev_st_hist = prev_conv_hist = prev_cpv_hist = None
        prev_rev_hist = None
        for _, r in hist.iterrows():
            s     = float(r.get('Spend', 0) or 0)
            c     = float(r.get('Conversions', 0) or 0)
            st_h  = float(r.get('Site Traffic', 0) or 0)
            rev_h = float(r.get('Revenue', 0) or 0) if has_revenue else 0
            if low_conv:
                cpv_h = s / st_h if st_h > 0 else None
                last_col_str = (
                    _mc(_m(cpv_h), cpv_h, prev_cpv_hist, invert=True)
                    if cpv_h is not None else NA_mom
                )
                prev_cpv_hist = cpv_h
            else:
                cpa_h = s / c if c else None
                last_col_str = (
                    _mc(_m(cpa_h), cpa_h, prev_cpa_hist, invert=True)
                    if cpa_h is not None else NA_mom
                )
                prev_cpa_hist = cpa_h
            row_cells = [
                str(r['Month']),
                _mc(_m(s), s, prev_spend),
                _mc(_n(st_h), st_h, prev_st_hist),
                _mc(_n(c), c, prev_conv_hist),
                last_col_str,
            ]
            if has_revenue:
                row_cells.append(_mc(_m(rev_h), rev_h, prev_rev_hist))
            history_rows_html += _td_row(row_cells)
            prev_spend, prev_st_hist, prev_conv_hist = s, st_h, c
            if has_revenue:
                prev_rev_hist = rev_h
            sp_months.append(str(r['Month']))
            sp_convs.append(c)
            sp_cpas.append(s / c if c > 0 else 0)
            sp_sts.append(st_h)

    if history_rows_html or prev_data:
        has_mom = True
        p_spend_badge = float(prev_data['Spend']) if prev_data else None
        p_st_badge    = float(prev_data['Site Traffic']) if prev_data else None
        p_conv_badge  = float(prev_data['Conversions']) if prev_data else None
        p_rev_badge   = (float(prev_data.get('Revenue', 0) or 0)
                         if prev_data and float(prev_data.get('Revenue', 0) or 0) else None)
        if low_conv:
            curr_cpv_val = spnd / st if st > 0 else None
            p_cpv_badge  = (float(prev_data['Spend']) / float(prev_data['Site Traffic'])
                            if prev_data and float(prev_data.get('Site Traffic', 0)) else None)
            curr_last_col = (
                _mc(_m(curr_cpv_val), curr_cpv_val, p_cpv_badge, invert=True)
                if curr_cpv_val is not None else NA_mom
            )
            mom_head_cols = ['Month', 'Spend', 'Site Traffic', 'Conversions', 'Cost Per Visit']
        else:
            curr_cpa      = spnd / conv if conv else None
            p_cpa_badge   = (float(prev_data['Spend']) / float(prev_data['Conversions'])
                             if prev_data and float(prev_data.get('Conversions', 0)) else None)
            curr_last_col = (
                _mc(_m(curr_cpa), curr_cpa, p_cpa_badge, invert=True)
                if curr_cpa is not None else NA_mom
            )
            mom_head_cols = ['Month', 'Spend', 'Site Traffic', 'Conversions', 'CPA']
        if has_revenue:
            mom_head_cols.append('Revenue')
        mom_head = _th(mom_head_cols, right_from=1)
        curr_row_cells = [
            f'<strong>{report_month}</strong>',
            _mc(_m(spnd), spnd, p_spend_badge),
            _mc(_n(st), st, p_st_badge),
            _mc(_n(conv), conv, p_conv_badge),
            curr_last_col,
        ]
        if has_revenue:
            curr_row_cells.append(_mc(_m(rev), rev, p_rev_badge))
        curr_row = _td_row(curr_row_cells)
        mom_table = f'''<div class="table-wrap">
      <table>
        <thead>{mom_head}</thead>
        <tbody>{history_rows_html}{curr_row}</tbody>
      </table>
    </div>'''

    # ── Upsell block ──────────────────────────────────────────────────────────
    upsell_html = ''
    if upsell_data:
        curr_cpa_fmt = f"${upsell_data['curr_cpa']:,.2f}"
        avg_cpa_fmt  = f"${upsell_data['avg_cpa']:,.2f}"
        impr_fmt     = f"{upsell_data['improvement_pct']:.0f}"
        budget_fmt   = f"${upsell_data['budget_rec']:,.0f}"
        upsell_body  = (f'Your current cost-per-conversion (<strong>{curr_cpa_fmt}</strong>) is trending '
                        f'<strong>{impr_fmt}%</strong> below your 3-month benchmark (<strong>{avg_cpa_fmt}</strong>). '
                        f'We recommend a budget increase of <strong>{budget_fmt}</strong> to maximise volume and '
                        f'capture more market share while your acquisition costs are performing so strongly.')
        upsell_html  = f'''<div class="upsell-wrap">
      <img class="upsell-logo" src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">
      {_insight_box_html('Optimisation Opportunity: Efficiency Momentum', upsell_body, variant='green')}
    </div>'''

    NO_CLICKS      = {'CTV', 'Audio'}
    NO_COMP        = {'Display'}
    _CTR_BENCH_CH  = {'Display', 'Video'}
    _COMP_HIGH_CH  = {'CTV', 'Audio'}
    _COMP_VIDEO_CH = {'Video'}

    # ── Channel Performance table ─────────────────────────────────────────────
    out_hdrs = ['Channel', 'Spend %', 'Impressions', 'Attributed Site Traffic',
                'Conversions', 'Conversion Rate', 'Cost per Conv (CPA)']
    out_head = _th(out_hdrs)
    out_body = ''
    for _, r in grp_chan.iterrows():
        _, _, _, cvr_, _ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        ch       = str(r['_chan'])
        spnd_pct = _p(r.spnd / spnd * 100) if spnd else _p(0)
        cpa      = _m(r.spnd / r.conv) if r.conv else NA
        out_body += _td_row([ch, spnd_pct, _n(r.imp), _n(r.st), _n(r.conv), _p(cvr_), cpa])

    # ── Channel Engagement table ──────────────────────────────────────────────
    eng_hdrs = ['Channel', 'Impressions', 'Clicks', 'CTR', 'eCPC', 'CPM', 'Completion Rate']
    eng_head = _th(eng_hdrs)
    eng_body = ''
    for _, r in grp_chan.iterrows():
        ctr_, cpc_, cpm_, _, comp_ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        ch = str(r['_chan'])
        if ch in NO_CLICKS:
            ctr_cell = NA
        elif ch in _CTR_BENCH_CH:
            ctr_cell = f'<span>{_p(ctr_)}{_avg_badge(ctr_, 0.15)}</span>'
        else:
            ctr_cell = _p(ctr_)
        if ch in NO_COMP:
            comp_cell = NA
        elif ch in _COMP_HIGH_CH:
            comp_cell = f'<span>{_p(comp_)}{_avg_badge(comp_, 95.0)}</span>'
        elif ch in _COMP_VIDEO_CH:
            comp_cell = f'<span>{_p(comp_)}{_avg_badge(comp_, 50.0)}</span>'
        else:
            comp_cell = _p(comp_)
        eng_body += _td_row([
            ch, _n(r.imp),
            NA if ch in NO_CLICKS else _n(r.clk),
            ctr_cell,
            NA if ch in NO_CLICKS else _m(cpc_),
            _m(cpm_),
            comp_cell,
        ])

    engagement_insight_html = ''
    best_channel  = None
    best_variance = 0.0
    for _, r in grp_chan.iterrows():
        ch = str(r['_chan'])
        ctr_i, _, _, _, comp_i = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        candidates = []
        if ch in {'Display', 'Video'} and r.clk > 0:
            candidates.append((ctr_i - 0.15) / 0.15 * 100)
        if ch in {'CTV', 'Audio'} and r.imp > 0:
            candidates.append((comp_i - 95.0) / 95.0 * 100)
        if ch == 'Video' and r.imp > 0:
            candidates.append((comp_i - 50.0) / 50.0 * 100)
        for v in candidates:
            if v > best_variance:
                best_variance = v
                best_channel  = ch
    if best_channel:
        engagement_insight_html = _insight_box_html(
            'Performance Lead',
            f'<strong>{_h(best_channel)}</strong> is outperforming industry benchmarks by <strong>{best_variance:.0f}%</strong>, driving premium audience engagement.'
        )

    # ── Awareness table (Video, CTV, Audio) ──────────────────────────────────
    AWARE_CH = {'Video', 'CTV', 'Audio'}
    aware_rows = grp_chan[grp_chan['_chan'].isin(AWARE_CH)]
    aw_head = _th(['Channel', 'Impressions', 'Reach', 'Frequency', 'CPM', 'Completion Rate'])
    aw_body = ''
    aw_best_ch, aw_best_var = None, 0.0
    for _, r in aware_rows.iterrows():
        _, _, cpm_, _, comp_ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        ch   = str(r['_chan'])
        reach_val = int(r.uh) if r.uh > 0 else None
        freq_val  = r.imp / r.uh if r.uh > 0 else None
        reach_cell = _n(reach_val) if reach_val else NA
        freq_cell  = f'{freq_val:.1f}x' if freq_val else NA
        if ch in {'CTV', 'Audio'}:
            comp_cell = f'<span>{_p(comp_)}{_avg_badge(comp_, 95.0)}</span>'
            var = (comp_ - 95.0) / 95.0 * 100 if r.imp > 0 else 0
        else:
            comp_cell = f'<span>{_p(comp_)}{_avg_badge(comp_, 50.0)}</span>'
            var = (comp_ - 50.0) / 50.0 * 100 if r.imp > 0 else 0
        if var > aw_best_var:
            aw_best_var, aw_best_ch = var, ch
        aw_body += _td_row([ch, _n(r.imp), reach_cell, freq_cell, _m(cpm_), comp_cell])
    aw_framing = '<div class="context-box">These channels are built for <strong>reach and recall</strong> — not clicks. The goal is to get in front of as many people as possible, as often as needed, at an efficient cost per thousand impressions.</div>'
    aw_lead = _insight_box_html(
        'Performance Lead',
        f'<strong>{_h(aw_best_ch)}</strong> is outperforming industry benchmarks by <strong>{aw_best_var:.0f}%</strong>, driving premium audience engagement.'
    ) if aw_best_ch else ''
    awareness_insight_html = aw_lead + aw_framing

    # ── Conversion table (Display) ────────────────────────────────────────────
    CONV_CH = {'Display'}
    conv_rows = grp_chan[grp_chan['_chan'].isin(CONV_CH)]
    cv_head = _th(['Channel', 'Impressions', 'Clicks', 'CTR', 'eCPC',
                   'Site Traffic', 'Visit Rate', 'Conversions', 'CPA'])
    cv_body = ''
    for _, r in conv_rows.iterrows():
        ctr_, cpc_, _, _, _ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        ch         = str(r['_chan'])
        visit_rate = _p(r.st / r.imp * 100) if r.imp > 0 else NA
        cpa_cell   = _m(r.spnd / r.conv) if r.conv else NA
        cv_body += _td_row([ch, _n(r.imp), _n(r.clk), _p(ctr_), _m(cpc_),
                             _n(r.st), visit_rate, _n(r.conv), cpa_cell])

    # ── Top 10 Creatives table ────────────────────────────────────────────────
    top10_cre  = grp_cre.nlargest(10, 'st')
    cre_vrs    = [r['st'] / r['imp'] if r['imp'] > 0 and r['st'] <= r['imp'] else None
                  for _, r in top10_cre.iterrows()]
    _cre_vrs_valid = [v for v in cre_vrs if v is not None]
    avg_cre_vr = sum(_cre_vrs_valid) / len(_cre_vrs_valid) if _cre_vrs_valid else 0
    max_cre_vr = max(_cre_vrs_valid) if _cre_vrs_valid else 0
    # Header for Visit Rate column includes avg sub-label with white marker tick
    vr_avg_label = (_p(avg_cre_vr * 100) if avg_cre_vr > 0 else '')
    vr_th_label  = (f'Visit Rate'
                    f'<div class="th-sub"><span class="th-avg-tick"></span> Avg {vr_avg_label}</div>'
                    if vr_avg_label else 'Visit Rate')
    cre_head = (
        f'<tr>{_h_th("Creative")}'
        f'{_h_th("CTR",right=True)}{_h_th("eCPC",right=True)}'
        f'{_h_th("Completion Rate",right=True)}{_h_th("Attributed Site Traffic",right=True)}'
        f'{_h_th(vr_th_label,right=True,raw=True)}</tr>'
    )
    cre_body  = ''
    for idx, (_, r) in enumerate(top10_cre.iterrows()):
        ctr_, cpc_, cpm_, cvr_, comp_ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        comp_v   = NA if comp_ == 0 else _p(comp_)
        vr       = cre_vrs[idx]
        if vr is None:
            vr_bar = '<span class="na">No delivery this month</span>'
        else:
            vr_bar = _index_bar_html(_p(vr * 100), vr * 100, avg_cre_vr * 100, max_cre_vr * 100)
        cre_name = (f'<span class="name-cell">'
                    f'{_creative_type_dot(r["Creative"])}'
                    f'<span class="name-text">{_h(str(r["Creative"]))}</span>'
                    f'</span>')
        cre_body += _td_row([cre_name, _p(ctr_), _m(cpc_), comp_v, _n(r.st), vr_bar])

    creative_insight_html = ''
    if not top10_cre.empty and imp > 0 and st > 0:
        campaign_vr = st / imp
        grp_cre_vr  = top10_cre.copy()
        grp_cre_vr['_vr'] = grp_cre_vr.apply(
            lambda r: r['st'] / r['imp'] if r['imp'] > 0 and r['st'] <= r['imp'] else 0, axis=1)
        top_vr_row = grp_cre_vr.loc[grp_cre_vr['_vr'].idxmax()]
        if campaign_vr > 0:
            variance_pct = (top_vr_row['_vr'] - campaign_vr) / campaign_vr * 100
            if variance_pct > 0:
                cre_name = _h(str(top_vr_row['Creative']))
                creative_insight_html = _insight_box_html(
                    'Traffic Efficiency',
                    f'<strong>&#8220;{cre_name}&#8221;</strong> was your most effective asset at driving web traffic this month, achieving a Visit Rate <strong>{variance_pct:.0f}%</strong> above the campaign average.'
                )

    # ── Top 10 Sites table ────────────────────────────────────────────────────
    site_body = ''
    site_vrs     = [r.st / r.imp * 100 if r.imp > 0 else 0 for _, r in grp_site.iterrows()]
    site_vrs_pos = [v for v in site_vrs if v > 0]
    avg_site_vr  = sum(site_vrs_pos) / len(site_vrs_pos) if site_vrs_pos else 0
    max_site_vr  = max(site_vrs_pos) if site_vrs_pos else 0
    vr_avg_label = _p(avg_site_vr) if avg_site_vr > 0 else ''
    vr_th_label  = (f'Visit Rate'
                    f'<div class="th-sub"><span class="th-avg-tick"></span> Blended Avg {vr_avg_label}</div>'
                    if vr_avg_label else 'Visit Rate')
    site_head = (
        f'<tr>{_h_th("Publisher Site")}'
        f'{_h_th("CPM",right=True)}{_h_th("CTR %",right=True)}'
        f'{_h_th(vr_th_label,right=True,raw=True)}</tr>'
    )
    for idx, (_, r) in enumerate(grp_site.iterrows()):
        cpm_    = r.spnd / r.imp * 1000 if r.imp else 0
        vr_val  = site_vrs[idx]
        vr_bar  = _index_bar_html(_p(vr_val), vr_val, avg_site_vr, max_site_vr) if r.imp > 0 else NA
        site_name = (f'<span class="name-cell">'
                     f'<span class="site-dot"></span>'
                     f'<span class="name-text">{_h(str(r["Site"]))}</span>'
                     f'</span>')
        ctr_val = r.clk / r.imp * 100 if r.imp > 0 else 0
        site_body += _td_row([site_name, _m(cpm_), _p(ctr_val), vr_bar])

    site_insight_html = ''
    if not grp_site.empty and imp > 0 and st > 0:
        campaign_vr = st / imp
        grp_site_vr = grp_site.copy()
        grp_site_vr['_vr'] = grp_site_vr.apply(
            lambda r: r['st'] / r['imp'] if r['imp'] > 0 else 0, axis=1)
        top_site_row = grp_site_vr.loc[grp_site_vr['_vr'].idxmax()]
        if campaign_vr > 0 and top_site_row['_vr'] > 0:
            site_variance_pct = (top_site_row['_vr'] - campaign_vr) / campaign_vr * 100
            site_name         = _h(str(top_site_row['Site']))
            site_insight_html = _insight_box_html(
                'Key Environments',
                f'<strong>{site_name}</strong> converted impressions to site visits at a rate <strong>{site_variance_pct:.0f}%</strong> above the campaign average — your most efficient placement this month. We have optimised the campaign to direct more of your investment here to maximise overall performance.'
            )

    # ── Glossary ──────────────────────────────────────────────────────────────
    roas_term = '''
      <div class="glossary-item">
        <span class="glossary-term">ROAS (Return on Ad Spend)</span> — The amount of revenue generated for every dollar spent on advertising.
      </div>''' if has_revenue else ''

    glossary_items = f'''
      <div class="glossary-item">
        <span class="glossary-term">Impressions</span> — The total number of times your ad was shown or played.
      </div>
      <div class="glossary-item">
        <span class="glossary-term">CPM (Cost Per Mille)</span> — The cost for every 1,000 ad impressions served.
      </div>
      <div class="glossary-item">
        <span class="glossary-term">CTR (Click-Through Rate)</span> — The percentage of people who saw your ad and chose to click it.
      </div>
      <div class="glossary-item">
        <span class="glossary-term">eCPC (Effective Cost Per Click)</span> — The average cost paid for each individual click.
      </div>
      <div class="glossary-item">
        <span class="glossary-term">Conversions</span> — Specific actions (e.g., Cart stage) completed by users after ad exposure.
      </div>
      <div class="glossary-item">
        <span class="glossary-term">Conversion Rate</span> — The percentage of impressions that successfully resulted in a conversion.
      </div>{roas_term}
      <div class="glossary-item">
        <span class="glossary-term">Attributed Site Traffic</span> — The number of users who visited your website after being exposed to the campaign.
      </div>'''

    # ── Ad Group Performance slide ────────────────────────────────────────────
    agp_cards_html = ''
    agp_card_count = 0
    if grp_adgroup is not None and not grp_adgroup.empty:
        _chan_order   = ['CTV', 'Audio', 'Video', 'Display']
        chans_present = grp_adgroup['_chan'].unique().tolist()
        sorted_chans  = [c for c in _chan_order if c in chans_present]
        sorted_chans += [c for c in chans_present if c not in _chan_order]

        # Separate remarketing rows from prospecting for every channel
        remark_by_chan = {}
        prosp_by_chan  = {}
        for chan in sorted_chans:
            rows    = grp_adgroup[grp_adgroup['_chan'] == chan].copy()
            rm_mask = rows['_ag_label'].str.lower().str.contains('remarketing', na=False)
            if rm_mask.any():
                remark_by_chan[chan] = rows[rm_mask]
            prosp_by_chan[chan] = rows[~rm_mask]

        # CPA threshold uses prospecting-only Display conv so remarketing doesn't
        # inflate the gate and skew prospecting charts
        disp_prosp      = prosp_by_chan.get('Display', pd.DataFrame(columns=['conv']))
        prosp_disp_conv = float(disp_prosp['conv'].sum()) if not disp_prosp.empty else 0
        # Full Display conv (incl. remarketing) — used only for the remarketing hub
        display_total_conv = float(
            grp_adgroup[grp_adgroup['_chan'] == 'Display']['conv'].sum()
        ) if 'Display' in grp_adgroup['_chan'].values else 0

        for chan in sorted_chans:
            use_cpa = False
            card    = _agp_card_html(chan, prosp_by_chan[chan], use_cpa)
            if card:
                agp_cards_html += card
                agp_card_count += 1

        rm_hub = _agp_remarketing_hub_html(remark_by_chan, display_total_conv, prosp_by_chan)
        if rm_hub:
            agp_cards_html += rm_hub
            agp_card_count += 1

    # ── Shared logo tags ──────────────────────────────────────────────────────
    foot_logo  = f'<img class="foot-logo"  src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">'
    cover_logo = f'<img class="cover-logo" src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">'

    # ── Overview (KPI) slide content ──────────────────────────────────────────
    conv_note = f'<div class="ov-conv-note">Conversions tracked: <strong>{_h(conv_label)}</strong></div>'

    def _mc_trend(curr, prev, invert=False):
        if not prev:
            return ''
        pct = (curr - prev) / abs(prev) * 100
        if abs(pct) < 0.05:
            return '<span style="color:#6B7280;font-size:13px">no change vs last month</span>'
        is_good = pct < 0 if invert else pct > 0
        color   = '#5BC2E7' if is_good else '#EF426F'
        arrow   = '&#9650;' if pct > 0 else '&#9660;'
        return f'<span style="color:{color};font-size:13px;font-weight:600">{arrow} {abs(pct):.1f}% vs last month</span>'

    mc_invest  = _metric_card_html('Investment',   f'${spnd:,.0f}', _mc_trend(spnd, ps))
    mc_traffic = _metric_card_html('Site Traffic', _n(st),          _mc_trend(st, p_st))
    mc_conv    = _metric_card_html('Conversions',  _n(conv),        _mc_trend(conv, pc), conv_label)
    if low_conv:
        mc_cpa = _metric_card_html('Cost Per Visit', _m(cpv) if st else NA,
                                   _mc_trend(cpv, prev_cpv, invert=True))
    else:
        mc_cpa = _metric_card_html('CPA', _m(cpa_kpi) if conv else NA,
                                   _mc_trend(cpa_kpi, prev_cpa_kpi, invert=True))

    budget_segs = []
    for i, (_, r) in enumerate(grp_chan.sort_values('spnd', ascending=False).iterrows()):
        if r.spnd <= 0:
            continue
        pct = r.spnd / spnd * 100 if spnd else 0
        budget_segs.append((str(r['_chan']), pct, _BB_COLORS[i % len(_BB_COLORS)]))

    budget_panel = f'''<div class="ov-lower-panel ov-budget-col">
      <div class="ov-panel-label">Budget Split &amp; Impact</div>
      {_stacked_budget_bar_html(budget_segs)}
    </div>'''

    if has_revenue:
        curr_roas    = roas(rev, spnd)
        roas_trend_h = _mc_trend(curr_roas, prev_roas)
        rev_trend_h  = _mc_trend(rev, pr)
        rev_gen_panel = f'''<div class="ov-lower-panel ov-revgen-col">
      <div class="ov-panel-label">Revenue Gen</div>
      <div class="ov-panel-big">{_m(rev)}</div>
      {f'<div class="ov-panel-trend">{rev_trend_h}</div>' if rev_trend_h else ""}
      <div class="ov-panel-roas"><span class="ov-panel-key">ROAS</span> <strong>{_rx(curr_roas)}</strong>{(" &nbsp;" + roas_trend_h) if roas_trend_h else ""}</div>
    </div>'''
        metric_grid = f'<div class="ov-metric-grid">{mc_invest}{mc_traffic}{mc_conv}{mc_cpa}</div>'
    else:
        if low_conv:
            cpv_trend_h = _mc_trend(cpv, prev_cpv, invert=True)
            rev_gen_panel = f'''<div class="ov-lower-panel ov-revgen-col">
      <div class="ov-panel-label">Cost Per Visit</div>
      <div class="ov-panel-big">{_m(cpv) if st else NA}</div>
      {f'<div class="ov-panel-trend">{cpv_trend_h}</div>' if cpv_trend_h else ""}
    </div>'''
        else:
            cpa_trend_h = _mc_trend(cpa_kpi, prev_cpa_kpi, invert=True)
            rev_gen_panel = f'''<div class="ov-lower-panel ov-revgen-col">
      <div class="ov-panel-label">Cost per Conversion</div>
      <div class="ov-panel-big">{_m(cpa_kpi) if conv else NA}</div>
      {f'<div class="ov-panel-trend">{cpa_trend_h}</div>' if cpa_trend_h else ""}
    </div>'''
        metric_grid = f'<div class="ov-metric-grid three">{mc_invest}{mc_traffic}{mc_conv}</div>'

    lower_row = f'<div class="ov-lower-row">{budget_panel}{rev_gen_panel}</div>'

    # Small metric key shown below conv_note
    _ov_key_items = []
    if low_conv:
        _ov_key_items.append('<strong>CPV</strong> Cost Per Visit &mdash; Spend &div; Site Traffic')
    else:
        _ov_key_items.append('<strong>CPA</strong> Cost per Acquisition &mdash; Spend &div; Conversions')
    if has_revenue:
        _ov_key_items.append('<strong>ROAS</strong> Return on Ad Spend &mdash; Revenue &div; Spend')
    ov_key = ('<div class="ov-key">'
              + ' <span class="ov-key-sep">&middot;</span> '.join(
                  f'<span class="ov-key-item">{t}</span>' for t in _ov_key_items)
              + '</div>') if _ov_key_items else ''

    kpi_inner = f'''{metric_grid}
  {lower_row}
  {conv_note}
  {ov_key}'''

    # ── MoM slide content ─────────────────────────────────────────────────────
    if mom_insight_text:
        takeaway = _insight_box_html('Key Takeaway', mom_insight_text)
    else:
        conv_chg = _t(conv, pc)
        st_chg   = _t(st, p_st)
        lines    = []
        if conv_chg is not None:
            arrow = '&#9650;' if conv_chg >= 0 else '&#9660;'
            col   = '#5BC2E7' if conv_chg >= 0 else '#EF426F'
            lines.append(f'Conversions <span style="color:{col};font-weight:700">{arrow} {abs(conv_chg):.0f}%</span> month-on-month.')
        if st_chg is not None:
            arrow = '&#9650;' if st_chg >= 0 else '&#9660;'
            col   = '#5BC2E7' if st_chg >= 0 else '#EF426F'
            lines.append(f'Site Traffic <span style="color:{col};font-weight:700">{arrow} {abs(st_chg):.0f}%</span> month-on-month.')
        body = ' '.join(lines) if lines else 'Historical data is accumulating as the campaign continues.'
        takeaway = _insight_box_html('Month Summary', body)

    if low_conv:
        spark_svg       = _mom_sparkline_svg(sp_months, sp_sts, report_month, float(st), label='Site Traffic')
        sparkline_title = 'Site Traffic — Monthly Trend'
    else:
        spark_svg       = _mom_sparkline_svg(sp_months, sp_convs, report_month, float(conv))
        sparkline_title = 'Conversions — Monthly Trend'
    sparkline_html = ''
    if spark_svg:
        sparkline_html = f'''<div class="sparkline-wrap">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.12em;color:#5BC2E7;margin-bottom:10px">{sparkline_title}</div>
      {spark_svg}
    </div>'''

    right_col = (
        f'<div class="mom-right">{sparkline_html}{takeaway}</div>'
        if sparkline_html else f'<div class="mom-right">{takeaway}</div>'
    )

    mom_inner = f'''<div class="mom-split">
    <div class="mom-left">{mom_table}</div>
    {right_col}
  </div>'''

    # ── Table slide helper ────────────────────────────────────────────────────
    def _tslide(title, table_html, extra=''):
        nrows = table_html.count('<tr') - 1  # subtract header row
        td_pad = max(6, min(32, 80 // max(nrows, 1)))
        sparse_style = f' style="--td-pad:{td_pad}px"'
        return f'''  <h2 class="slide-title slide-title-left">{title}</h2>
  <div class="slide-main table-main">
    <div class="table-wrap"{sparse_style}>{table_html}</div>
    {extra}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''

    # ── Assemble slides ───────────────────────────────────────────────────────
    slides = []

    # Cover
    dl = f'<div class="cover-date">{_h(date_range)}</div>' if date_range else ''
    slides.append(('cover', f'''  <div class="cover-inner">
    {cover_logo}
    <div class="cover-rule"></div>
    <div class="cover-client">{_h(client_name)}</div>
    <div class="cover-report-title">Omnichannel Report</div>
    <div class="cover-month">{_h(report_month)}</div>
    {dl}
    <div class="cover-hint">Use arrow keys or buttons to navigate</div>
    <div class="cover-confidential">Confidential — prepared by MediaWorks</div>
  </div>'''))

    # Overview (KPI)
    slides.append(('overview', f'''  <h2 class="slide-title slide-title-left">Campaign Overview</h2>
  <div class="slide-main overview-main">
    {kpi_inner}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    # MoM (conditional)
    if has_mom:
        slides.append(('mom', f'''  <h2 class="slide-title slide-title-left">Month on Month Results</h2>
  <div class="slide-main mom-main">
    {mom_inner}
    {upsell_html}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    # ── Build awareness / display content blocks ─────────────────────────────
    # Minimal grey info boxes — sit under their respective table
    _aw_chans = sorted(aware_rows['_chan'].unique().tolist()) if not aware_rows.empty else []
    def _fmt_chan_list(chans):
        if not chans:
            return 'Awareness Channels'
        if len(chans) == 1:
            return chans[0]
        return ', '.join(chans[:-1]) + ' &amp; ' + chans[-1]
    aw_ch_label = _fmt_chan_list(_aw_chans)
    total_aware_conv = int(aware_rows['conv'].sum()) if not aware_rows.empty else 0
    _aw_conv_sentence = (
        f' While the objective for these channels is to drive awareness, they drove'
        f' <strong>{_n(total_aware_conv)}</strong> conversions in {report_month}.'
    ) if total_aware_conv > 0 else ''
    aw_info = (
        f'<div class="context-box"><strong style="color:#C4B5D4">Awareness Channels</strong>'
        f' — {aw_ch_label} are built for <strong>reach and recall</strong>, not immediate'
        f' conversions. The goal is to put your brand in front of as many of the right people as'
        f' possible, as efficiently as possible.{_aw_conv_sentence}</div>'
    )
    disp_info = (
        '<div class="context-box"><strong style="color:#C4B5D4">Display Advertising</strong>'
        ' — Display is built for <strong>direct response</strong>, driving site visits and'
        ' conversions. The visit rate measures how effectively ad exposure translates into'
        ' real website traffic.</div>'
    )

    # Display performance insight: share of total campaign site traffic
    disp_st      = float(conv_rows['st'].sum()) if not conv_rows.empty else 0.0
    disp_st_pct  = disp_st / st * 100 if st > 0 else 0.0
    disp_perf_insight = (
        _insight_box_html(
            'Display Performance',
            f'Display delivered <strong>{_n(int(disp_st))}</strong> site visits this month'
            f' — <strong>{disp_st_pct:.0f}%</strong> of total campaign traffic,'
            f' making it the primary driver of direct response in the mix.',
            variant='cyan'
        ) if disp_st > 0 and disp_st_pct > 0 else ''
    )

    # Combined benchmark note spans all metrics shown across both tables
    _bench_parts = [
        '<span style="color:#5BC2E7;font-weight:700">&#9650;</span> Above benchmark',
        '<span style="color:#EF426F;font-weight:700">&#9660;</span> Below benchmark',
        '<strong>Industry Average Benchmarks:</strong>',
        'Completion (CTV &amp; Audio) <strong>95%</strong>',
        'Completion (Video) <strong>50%</strong>',
    ]
    if not conv_rows.empty:
        _bench_parts.append('CTR (Display) <strong>0.15%</strong>')
    combined_bench = (
        '<div class="bench-note">'
        + ' &nbsp;&#183;&nbsp; '.join(_bench_parts)
        + '</div>'
    )

    aw_table_html = f'<table><thead>{aw_head}</thead><tbody>{aw_body}</tbody></table>'
    cv_table_html = f'<table><thead>{cv_head}</thead><tbody>{cv_body}</tbody></table>'

    if not aware_rows.empty and not conv_rows.empty:
        # ── Combined channel performance slide ───────────────────────────────
        # 4-cell grid: left col = tables+info; right col = performance insights
        # Each row of the grid keeps its left section in line with its right insight.
        slides.append(('chperf', f'''  <h2 class="slide-title slide-title-left">Channel Performance</h2>
  <div class="slide-main ch-perf-main">
    <div class="ch-perf-body">
      <div class="ch-perf-section">
        <div class="ch-perf-col-label">Awareness — {aw_ch_label}</div>
        <div class="table-wrap">{aw_table_html}</div>
        {aw_info}
      </div>
      <div class="ch-perf-insight-cell">
        {aw_lead}
      </div>
      <div class="ch-perf-section">
        <div class="ch-perf-col-label">Display — Conversion &amp; Traffic</div>
        <div class="table-wrap">{cv_table_html}</div>
        {disp_info}
      </div>
      <div class="ch-perf-insight-cell">
        {disp_perf_insight}
      </div>
    </div>
    {combined_bench}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    elif not aware_rows.empty:
        # Awareness only
        slides.append(('table', _tslide('Awareness Performance',
            aw_table_html, awareness_insight_html + aw_info + combined_bench)))

    elif not conv_rows.empty:
        # Display only
        slides.append(('table', _tslide('Display Conversion Performance',
            cv_table_html, disp_info + disp_perf_insight + combined_bench)))

    # Shared inline colour key used on AGP, Sites and Creatives slides
    _vr_key_html = (
        '<div class="inline-key">'
        + _BAR_COLOR_KEY
        + '<span class="bar-key-sep">&middot;</span>'
        + '<span class="bar-key-note">Visit Rate: Site Visits &div; Impressions</span>'
        + '</div>'
    )

    # Ad Group Performance — one card per channel with ≥2 valid ad groups
    if agp_card_count >= 2:
        slides.append(('agp', f'''  <h2 class="slide-title slide-title-left">Platform and Audience Performance</h2>
  <div class="slide-main agp-main">
    <div class="agp-grid" data-count="{agp_card_count}">
      {agp_cards_html}
    </div>
    {_vr_key_html}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    def _optim_note(kind, top_names):
        names_html = ' &amp; '.join(f'<strong>{_h(n)}</strong>' for n in top_names)
        return (
            f'<div class="tbl-optim-note">'
            f'<span class="agp-footer-icon">{_ICON_COG}</span>'
            f'<span class="tbl-optim-text">Budget optimised towards top performing {kind} such as {names_html}</span>'
            f'</div>'
        )

    # Top Sites — only sites with above-average visit rate (matches cyan rows in the table)
    if not grp_site.empty:
        _site_perf = grp_site[(grp_site['st'] > 0) & (grp_site['imp'] > 0)].copy()
        _site_perf['_vr'] = _site_perf['st'] / _site_perf['imp']
        _site_avg_vr = _site_perf['st'].sum() / _site_perf['imp'].sum()
        _above_avg_sites = _site_perf[_site_perf['_vr'] >= _site_avg_vr].nlargest(2, '_vr')
        top_site_names = [str(r['Site']) for _, r in _above_avg_sites.iterrows()]
    else:
        top_site_names = []
    site_optim = _optim_note('sites', top_site_names) if top_site_names else ''
    slides.append(('table', _tslide('Top 10 Sites by Impressions',
        f'<table><thead>{site_head}</thead><tbody>{site_body}</tbody></table>',
        site_insight_html + site_optim + _vr_key_html)))

    # Top Creatives — only creatives with above-average visit rate (matches cyan rows in the table)
    if not top10_cre.empty:
        _cre_perf = top10_cre[(top10_cre['imp'] > 0) & (top10_cre['st'] > 0)].copy()
        _cre_perf['_vr'] = _cre_perf['st'] / _cre_perf['imp']
        _cre_avg_vr = _cre_perf['st'].sum() / _cre_perf['imp'].sum()
        _above_avg_cre = _cre_perf[_cre_perf['_vr'] >= _cre_avg_vr].nlargest(2, '_vr')
        top_cre_names = [str(r['Creative']) for _, r in _above_avg_cre.iterrows()]
    else:
        top_cre_names = []
    cre_optim = _optim_note('creatives', top_cre_names) if top_cre_names else ''
    slides.append(('table', _tslide('Top 10 Creatives by Site Traffic',
        f'<table><thead>{cre_head}</thead><tbody>{cre_body}</tbody></table>',
        creative_insight_html + cre_optim + _vr_key_html)))

    # Glossary
    gloss_attr = '<div class="glossary-attribution"><strong style="color:#C4B5D4">Note on Attribution:</strong> Both Conversions and Site Traffic are measured as View-Through + Click-Through — users who clicked an ad immediately, plus those who saw an ad and visited later within the attribution window.</div>'
    slides.append(('table', f'''  <h2 class="slide-title slide-title-left">Glossary of Terms</h2>
  <div class="slide-main">
    <div class="glossary-grid">{glossary_items}</div>
    {gloss_attr}
    <div class="foot-credit">Confidential — prepared for {_h(client_name)}</div>
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    total = len(slides)
    slide_divs = []
    for i, (kind, content) in enumerate(slides):
        active = ' active' if i == 0 else ''
        slide_divs.append(f'<div class="slide slide-{kind}{active}" id="s{i}">\n{content}\n</div>')
    slides_html = '\n'.join(slide_divs)

    dot_list = []
    for i in range(total):
        cls = 'dot dot-active' if i == 0 else 'dot'
        dot_list.append(f'<button class="{cls}" data-i="{i}" aria-label="Slide {i+1}"></button>')
    dots_html = ''.join(dot_list)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_h(client_name)} \u2014 {_h(report_month)}</title>
{_barlow_font_css()}
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    height: 100%; overflow: hidden;
    font-family: 'Barlow Semi Condensed', sans-serif;
    background: #111; color: #fff; font-size: 16px; line-height: 1.5;
  }}

  .deck {{ position: absolute; width: 1440px; height: 810px; overflow: hidden; transform-origin: top left; }}
  .slide {{ position: absolute; inset: 0; display: none; flex-direction: column; overflow: hidden; }}
  .slide.active {{ display: flex; }}

  /* ── Cover (unchanged) ─────────────────────────────────────────────────── */
  .slide-cover {{ background: #1A1A1B; overflow-y: auto; }}
  .slide-cover::before {{
    content: ''; position: absolute; top: -200px; right: -200px;
    width: 560px; height: 560px; border-radius: 50%; pointer-events: none;
    background: radial-gradient(circle, rgba(103,30,117,0.5) 0%, transparent 65%);
  }}
  .slide-cover::after {{
    content: ''; position: absolute; bottom: -120px; left: -120px;
    width: 420px; height: 420px; border-radius: 50%; pointer-events: none;
    background: radial-gradient(circle, rgba(239,66,111,0.22) 0%, transparent 65%);
  }}
  .cover-inner {{
    position: relative; z-index: 1; min-height: 100%;
    padding: 64px 108px 80px;
    display: flex; flex-direction: column; justify-content: center; align-items: flex-start;
  }}
  .cover-logo {{
    height: 34px; width: auto; align-self: flex-start;
    filter: brightness(0) invert(1); opacity: 0.88; margin-bottom: 56px;
  }}
  .cover-rule {{ width: 56px; height: 4px; background: #EF426F; border-radius: 2px; margin-bottom: 28px; }}
  .cover-client {{ font-size: 60px; font-weight: 700; letter-spacing: -0.6px; line-height: 1.05; margin-bottom: 10px; }}
  .cover-report-title {{ font-size: 22px; font-weight: 300; color: #9BA3AF; margin-bottom: 18px; letter-spacing: 0.03em; }}
  .cover-month {{ font-size: 34px; font-weight: 700; color: #EF426F; letter-spacing: -0.2px; margin-bottom: 8px; }}
  .cover-date {{ font-size: 15px; color: #6B7280; margin-bottom: 72px; }}
  .cover-hint {{ font-size: 14px; color: #4B5563; margin-bottom: 8px; letter-spacing: 0.05em; }}
  .cover-confidential {{ font-size: 13px; color: #374151; }}

  /* ── Slide chrome ──────────────────────────────────────────────────────── */
  .slide-title {{
    font-size: 30px; font-weight: 700; letter-spacing: -0.2px;
    padding: 28px 72px 18px; flex-shrink: 0; text-align: center;
  }}
  .slide-title-left {{ text-align: left; }}
  .slide-main {{ flex: 1; min-height: 0; overflow-y: auto; padding: 0 72px 12px; }}
  .slide-main.table-main {{ display: flex; flex-direction: column; justify-content: flex-start; }}
  .slide-foot {{
    height: 40px; min-height: 40px; flex-shrink: 0;
    display: flex; align-items: center; justify-content: flex-end;
    padding: 0 72px; border-top: 1px solid rgba(255,255,255,0.05);
  }}
  .inline-key {{
    flex-shrink: 0; margin-top: 10px;
    display: flex; align-items: center; gap: 0; flex-wrap: nowrap;
    font-size: 12px; color: #9BA3AF;
  }}
  .bar-key-item {{ display: flex; align-items: center; gap: 6px; white-space: nowrap; }}
  .bar-key-dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }}
  .bar-key-sep {{ margin: 0 12px; color: #6B7280; }}
  .bar-key-note {{ color: #6B7280; white-space: nowrap; }}
  .foot-logo {{ height: 20px; width: auto; filter: brightness(0) invert(1); opacity: 0.4; }}
  .foot-credit {{ margin-top: 20px; font-size: 13px; color: #374151; text-align: center; }}

  /* ── MetricCards ───────────────────────────────────────────────────────── */
  .ov-metric-grid {{
    display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 16px;
  }}
  .ov-metric-grid.three {{ grid-template-columns: repeat(3, 1fr); }}
  .metric-card {{
    background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px; padding: 18px 22px;
  }}
  .mc-label {{
    font-size: 14px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.12em; color: #9BA3AF; margin-bottom: 8px;
  }}
  .mc-value {{
    font-size: 48px; font-weight: 700; color: #fff;
    font-variant-numeric: tabular-nums; line-height: 1.1; margin-bottom: 6px;
  }}
  .mc-trend {{ font-size: 13px; }}
  .mc-sub {{ font-size: 13px; color: #6B7280; margin-top: 2px; }}

  /* ── Overview lower row ─────────────────────────────────────────────────── */
  .overview-main {{
    display: flex; flex-direction: column; justify-content: center;
    padding-top: 4px; padding-bottom: 4px;
  }}
  .ov-lower-row {{
    display: grid; grid-template-columns: 2fr 1fr; gap: 16px; margin-bottom: 14px;
  }}
  .ov-lower-panel {{
    background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px; padding: 20px 24px;
  }}
  .ov-budget-col {{ display: flex; flex-direction: column; gap: 14px; }}
  .ov-revgen-col {{ display: flex; flex-direction: column; justify-content: center; gap: 6px; }}
  .ov-panel-label {{
    font-size: 14px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.14em; color: #9BA3AF;
  }}
  .ov-panel-big {{
    font-size: 56px; font-weight: 700; color: #fff;
    font-variant-numeric: tabular-nums; line-height: 1.05;
  }}
  .ov-panel-trend {{ font-size: 14px; font-weight: 600; }}
  .ov-panel-roas {{ font-size: 15px; color: #9BA3AF; margin-top: 6px; border-top: 1px solid rgba(255,255,255,0.07); padding-top: 10px; }}
  .ov-panel-roas strong {{ color: #fff; font-size: 18px; font-weight: 700; }}
  .ov-panel-key {{ font-size: 13px; text-transform: uppercase; letter-spacing: 0.08em; }}
  .ov-conv-note {{
    font-size: 14px; color: #6B7280; display: inline-block;
    padding: 6px 12px; background: rgba(255,255,255,0.03);
    border-left: 3px solid #5BC2E7; border-radius: 0 6px 6px 0;
  }}
  .ov-conv-note strong {{ color: #fff; font-weight: 600; }}
  .trend-line {{ font-size: 13px; font-weight: 600; letter-spacing: 0.04em; margin-top: 4px; }}
  .ov-key {{
    margin-top: 10px; display: flex; align-items: center; flex-wrap: wrap;
    font-size: 12px; color: #4B5563; gap: 0;
  }}
  .ov-key-item {{ color: #4B5563; }}
  .ov-key-item strong {{ color: #6B7280; font-weight: 700; }}
  .ov-key-sep {{ margin: 0 10px; color: #374151; }}

  /* ── Stacked budget bar ─────────────────────────────────────────────────── */
  .sbb-track {{
    display: flex; height: 48px; border-radius: 10px; overflow: hidden; gap: 3px;
  }}
  .sbb-seg {{
    height: 100%; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center;
    padding: 0 8px; overflow: hidden;
  }}
  .sbb-inner-label {{
    font-size: 14px; font-weight: 700; color: #fff;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }}
  .sbb-legend {{
    display: flex; flex-wrap: wrap; gap: 6px 14px; margin-top: 8px;
  }}
  .sbb-legend-item {{
    display: flex; align-items: center; gap: 5px;
    font-size: 12px; color: rgba(255,255,255,0.7);
    white-space: nowrap;
  }}
  .sbb-legend-dot {{
    width: 10px; height: 10px; border-radius: 3px; flex-shrink: 0;
  }}
  .sbb-legend-item strong {{ color: #fff; }}

  /* ── MoM slide ──────────────────────────────────────────────────────────── */
  .mom-main {{ display: flex; flex-direction: column; justify-content: center; padding-top: 4px; }}
  .mom-split {{ display: grid; grid-template-columns: 58fr 38fr; gap: 28px; align-items: stretch; }}
  .mom-left {{ display: flex; flex-direction: column; }}
  .mom-left .table-wrap {{ border-radius: 14px; flex: 1; }}
  .mom-right {{ display: flex; flex-direction: column; gap: 14px; }}
  .sparkline-wrap {{
    background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.07);
    border-radius: 14px; padding: 16px 18px; flex: 1;
    display: flex; flex-direction: column; min-height: 120px;
  }}
  .sparkline-wrap svg {{ flex: 1; min-height: 0; }}

  /* ── Combined Channel Performance slide ────────────────────────────────── */
  .ch-perf-main {{
    display: flex; flex-direction: column; gap: 10px; padding-top: 2px;
  }}
  /* 4-cell grid: 2 rows × 2 cols (left 2/3 table sections, right 1/3 insights)
     align-items:start lets each row shrink to its content height               */
  .ch-perf-body {{
    display: grid;
    grid-template-columns: 2fr 1fr;
    grid-template-rows: auto auto;
    gap: 14px 20px;
    align-items: start;
    flex: 1; min-height: 0;
  }}
  .ch-perf-section {{
    display: flex; flex-direction: column; gap: 8px;
  }}
  .ch-perf-col-label {{
    font-size: 12px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.13em; color: #9BA3AF; flex-shrink: 0;
    padding-bottom: 8px; border-bottom: 2px solid rgba(255,255,255,0.08);
  }}
  /* right-column insight cells — offset to align with the table, not the label */
  .ch-perf-insight-cell {{
    display: flex; flex-direction: column; gap: 10px;
    padding-top: 28px;
  }}
  .ch-perf-insight-cell .insight-v2 {{ margin-top: 0; }}
  /* grey info boxes inside sections */
  .ch-perf-section .context-box {{ margin-top: 0; }}
  .ch-perf-bench {{ flex-shrink: 0; }}

  /* ── IndexBar ───────────────────────────────────────────────────────────── */
  .ibar-cell {{ display: flex; align-items: center; gap: 10px; min-width: 130px; }}
  .ibar-val {{ font-weight: 700; font-size: 14px; min-width: 52px; text-align: right; flex-shrink: 0; font-variant-numeric: tabular-nums; }}
  .ibar-track {{
    flex: 1; height: 8px; background: #1f2937;
    border-radius: 9999px; position: relative; min-width: 60px;
  }}
  .ibar-fill {{ height: 100%; border-radius: 9999px; }}

  /* ── InsightBox v2 ──────────────────────────────────────────────────────── */
  .insight-v2 {{
    position: relative; padding: 16px 20px 16px 26px;
    border-radius: 12px; border: 1px solid; margin-top: 14px; overflow: hidden;
  }}
  .iv2-bar {{
    position: absolute; left: 0; top: 0; bottom: 0;
    width: 6px; border-radius: 12px 0 0 12px;
  }}
  .iv2-inner {{ padding-left: 0; }}
  .iv2-header {{ display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }}
  .iv2-icon {{ display: flex; flex-shrink: 0; }}
  .iv2-title {{
    font-size: 14px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.08em; line-height: 1;
  }}
  .iv2-body {{ font-size: 14px; color: #D1D5DB; line-height: 1.65; }}
  .iv2-body strong {{ color: #fff; }}

  /* ── Table: name cells with dot indicators ──────────────────────────────── */
  .name-cell {{ display: flex; align-items: center; gap: 8px; }}
  .name-text {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 220px; }}
  .type-dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }}
  .site-dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; background: #6B7280; }}

  /* ── Table: IndexBar column sub-label in header ─────────────────────────── */
  .th-sub {{
    font-size: 12px; color: #6B7280; text-transform: none; letter-spacing: normal;
    font-weight: 500; margin-top: 4px; display: flex; align-items: center;
    justify-content: flex-end; gap: 5px;
  }}
  .th-avg-tick {{
    display: inline-block; width: 2px; height: 10px;
    background: #fff; border-radius: 1px; flex-shrink: 0;
  }}

  /* ── Ad Group Performance slide ─────────────────────────────────────────── */
  /* agp-main uses a 2-row CSS grid so the colour key always gets its natural
     height (auto) and the card grid fills the remaining 1fr — no clipping. */
  .agp-main {{
    flex: 1; min-height: 0; overflow: hidden;
    display: grid; grid-template-rows: 1fr auto;
  }}
  .agp-grid {{
    display: grid; gap: 14px; min-height: 0; overflow: hidden; align-items: start;
  }}
  .agp-grid[data-count="1"] {{ grid-template-columns: 1fr; grid-template-rows: 1fr; max-width: 520px; margin: 0 auto; width: 100%; }}
  /* count=2: cap width so cards aren't excessively wide */
  .agp-grid[data-count="2"] {{ grid-template-columns: repeat(2, 1fr); grid-template-rows: 1fr; max-width: 960px; margin: 0 auto; width: 100%; }}
  .agp-grid[data-count="3"] {{ grid-template-columns: repeat(3, 1fr); grid-template-rows: 1fr; }}
  /* Multi-row grids: align-self:start so they don't stretch to fill agp-main */
  .agp-grid[data-count="4"] {{ grid-template-columns: repeat(2, 1fr); grid-template-rows: repeat(2, auto); align-self: start; }}
  /* 4 channel cards (top row) + remarketing hub centred under middle two */
  .agp-grid[data-count="5"] {{ grid-template-columns: repeat(4, 1fr); grid-template-rows: auto auto; align-self: start; }}
  .agp-grid[data-count="5"] .agp-card-rm {{ grid-column: 2 / 4; }}
  .agp-card {{
    background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px; overflow: hidden;
    display: flex; flex-direction: column; min-height: 0;
  }}
  .agp-card-header {{
    padding: 12px 20px; border-bottom: 1px solid rgba(255,255,255,0.07);
    display: flex; align-items: flex-start; justify-content: space-between;
    flex-shrink: 0;
  }}
  .agp-chan-name {{
    font-size: 14px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.12em; color: #fff; line-height: 1.3;
  }}
  .agp-metric-info {{ text-align: right; }}
  .agp-metric-name {{
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.1em; color: #9BA3AF; display: block;
  }}
  .agp-avg-chip {{
    font-size: 13px; color: #9BA3AF; margin-top: 3px;
    display: flex; align-items: center; justify-content: flex-end; gap: 5px;
  }}
  .agp-avg-tick {{
    display: inline-block; width: 2px; height: 10px;
    background: #fff; border-radius: 1px; flex-shrink: 0;
  }}
  /* rows section stretches to fill remaining card height, rows share space evenly */
  .agp-rows {{
    flex: 1; min-height: 0; overflow: hidden;
    display: flex; flex-direction: column;
  }}
  .agp-row {{
    flex: 1; min-height: 55px; max-height: 80px;
    display: flex; align-items: center; gap: 18px;
    padding: 0 20px; border-bottom: 1px solid rgba(255,255,255,0.05);
  }}
  .agp-row:last-child {{ border-bottom: none; }}
  .agp-row-left {{ width: 88px; flex-shrink: 0; }}
  .agp-ag-name {{
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.07em; color: #6B7280; display: block; margin-bottom: 3px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }}
  .agp-val {{
    font-size: 17px; font-weight: 700; font-variant-numeric: tabular-nums;
    display: block; line-height: 1;
  }}
  .agp-track {{
    flex: 1; height: 12px; background: #1f2937;
    border-radius: 6px; position: relative; min-width: 80px;
  }}
  .agp-fill {{ height: 100%; border-radius: 6px; }}
  .agp-avg-line {{
    position: absolute; left: 50%; top: -5px;
    height: 22px; width: 2px; background: #fff;
    border-radius: 1px; transform: translateX(-50%);
    box-shadow: 0 0 4px rgba(0,0,0,0.6);
  }}
  .agp-card-footer {{
    flex-shrink: 0;
    padding: 7px 20px 8px;
    font-size: 13px; color: #D1D5DB; line-height: 1.35;
    border-top: 1px solid rgba(251,191,36,0.25);
    background: rgba(251,191,36,0.07);
    display: flex; align-items: flex-start; gap: 8px;
  }}
  .agp-footer-text {{
    flex: 1; min-width: 0;
    overflow: hidden;
    display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical;
  }}
  .agp-card-footer strong {{ color: #fff; font-weight: 700; }}
  .agp-footer-icon {{ color: #FBBF24; flex-shrink: 0; display: flex; padding-top: 1px; }}
  .tbl-optim-note {{
    margin-top: 12px;
    padding: 8px 18px;
    font-size: 13px; color: #D1D5DB; line-height: 1.35;
    border: 1px solid rgba(251,191,36,0.25);
    border-left: 3px solid #FBBF24;
    background: rgba(251,191,36,0.07);
    border-radius: 0 8px 8px 0;
    display: flex; align-items: flex-start; gap: 8px;
  }}
  .tbl-optim-text {{
    flex: 1; min-width: 0;
    overflow: hidden;
    display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical;
  }}
  .tbl-optim-note strong {{ color: #fff; font-weight: 700; }}

  /* ── AGP metric key bar ────────────────────────────────────────────────── */
  .agp-key {{
    flex-shrink: 0; padding: 8px 0 2px;
    display: flex; align-items: center; flex-wrap: wrap; gap: 0;
    font-size: 12px; color: #4B5563;
  }}
  .agp-key-item {{ display: flex; align-items: center; gap: 5px; }}
  .agp-key-item strong {{ color: #6B7280; font-weight: 700; }}
  .agp-key-note {{ color: #374151; }}
  .agp-key-sep {{ margin: 0 14px; color: #374151; }}
  .agp-key-dot {{
    width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0;
    opacity: 0.7;
  }}

  /* ── Remarketing Hub card ──────────────────────────────────────────────── */
  .agp-card-rm {{
    background: rgba(96,165,250,0.06);
    border-color: rgba(96,165,250,0.30);
    align-self: start; /* size to content; don't stretch when it has few rows */
  }}
  /* Hub rows: no bar tracks — just left label + right value/badge */
  .agp-rm-rows {{ justify-content: center; }}
  .agp-rm-row {{
    height: 58px; flex-shrink: 0;
    display: flex; align-items: center; justify-content: space-between;
    padding: 0 24px; border-bottom: 1px solid rgba(255,255,255,0.05);
  }}
  .agp-rm-row:last-child {{ border-bottom: none; }}
  .agp-rm-left {{ display: flex; flex-direction: column; gap: 3px; }}
  .agp-rm-chan {{
    font-size: 14px; font-weight: 700; color: #fff; line-height: 1;
  }}
  .agp-rm-lbl {{
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.09em; color: #6B7280;
  }}
  .agp-rm-right {{
    display: flex; flex-direction: column; align-items: flex-end; gap: 4px;
  }}
  .agp-rm-val {{
    font-size: 22px; font-weight: 700; font-variant-numeric: tabular-nums;
    line-height: 1;
  }}
  .agp-rm-vs {{
    font-size: 12px; font-weight: 600; letter-spacing: 0.03em; line-height: 1;
  }}

  /* ── Upsell ─────────────────────────────────────────────────────────────── */
  .upsell-wrap {{ position: relative; margin-top: 16px; }}
  .upsell-logo {{
    position: absolute; top: 18px; right: 18px; height: 20px; width: auto;
    filter: brightness(0) invert(1); opacity: 0.7; z-index: 1;
  }}
  .upsell-wrap .insight-v2 {{ margin-top: 0; padding-right: 60px; }}

  /* ── Tables ─────────────────────────────────────────────────────────────── */
  .table-wrap {{
    background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px; overflow: hidden; overflow-x: auto;
  }}
  table {{ width: 100%; border-collapse: collapse; font-size: 15px; min-width: 600px; }}
  thead tr {{ background: rgba(255,255,255,0.06); }}
  th {{
    color: #9BA3AF; padding: 12px 16px; font-weight: 600; font-size: 13px;
    text-transform: uppercase; letter-spacing: 0.08em; min-width: 80px;
    vertical-align: bottom; line-height: 1.3; text-align: right;
  }}
  th:first-child {{ min-width: 160px; text-align: left; }}
  td {{
    padding: var(--td-pad, 13px) 16px; border-bottom: 1px solid rgba(255,255,255,0.05);
    color: #fff; vertical-align: middle; text-align: left;
  }}
  td:first-child {{ word-break: break-word; overflow-wrap: anywhere; max-width: 240px; line-height: 1.4; }}
  td.right {{ text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }}
  tr:last-child td {{ border-bottom: none; }}
  tbody tr:hover td {{ background: rgba(255,255,255,0.04); }}

  /* ── Other components ───────────────────────────────────────────────────── */
  .na {{ color: #4B5563; font-size: 13px; }}
  .context-box {{
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 10px; padding: 12px 18px;
    margin-top: 14px; font-size: 14px; color: #9BA3AF; line-height: 1.6;
  }}
  .context-box strong {{ color: #C4B5D4; }}
  .bench-note {{
    display: inline-block; margin-top: 10px; padding: 7px 12px;
    background: rgba(255,255,255,0.03); border-left: 3px solid #5BC2E7;
    border-radius: 0 6px 6px 0; font-size: 13px; color: #9BA3AF;
  }}
  .bench-note strong {{ color: #fff; font-weight: 600; }}
  .glossary-grid {{ display: grid; grid-template-columns: repeat(2,1fr); gap: 8px 28px; margin-bottom: 14px; }}
  .glossary-item {{
    font-size: 14px; color: #9BA3AF; line-height: 1.5; padding: 8px 12px;
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-left: 3px solid rgba(103,30,117,0.8); border-radius: 0 8px 8px 0;
  }}
  .glossary-term {{ font-weight: 700; color: #C4B5D4; }}
  .glossary-attribution {{
    font-size: 14px; color: #9BA3AF; padding: 10px 14px;
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07); border-radius: 8px; line-height: 1.6;
  }}

  /* ── Navigation ─────────────────────────────────────────────────────────── */
  .nav-arrow {{
    position: absolute; top: 50%; transform: translateY(-50%); z-index: 200;
    background: rgba(37,37,38,0.88); border: 1px solid rgba(255,255,255,0.1);
    color: #fff; width: 44px; height: 44px; border-radius: 50%; font-size: 18px;
    cursor: pointer; display: flex; align-items: center; justify-content: center;
    transition: background 0.15s, border-color 0.15s;
    -webkit-tap-highlight-color: transparent; user-select: none;
  }}
  .nav-arrow:hover {{ background: #671E75; border-color: #671E75; }}
  .nav-prev {{ left: 14px; }}
  .nav-next {{ right: 14px; }}
  .nav-bar {{
    position: absolute; bottom: 0; left: 0; right: 0; height: 48px; z-index: 200;
    display: flex; align-items: center; justify-content: center; gap: 10px;
    background: rgba(26,26,27,0.92); backdrop-filter: blur(10px);
    border-top: 1px solid rgba(255,255,255,0.05);
  }}
  .dot {{
    width: 7px; height: 7px; border-radius: 50%; background: rgba(255,255,255,0.18);
    border: none; cursor: pointer; padding: 0; transition: background 0.2s, transform 0.2s;
  }}
  .dot.dot-active {{ background: #EF426F; transform: scale(1.45); }}
  .dot:hover:not(.dot-active) {{ background: rgba(255,255,255,0.42); }}
  .slide-counter {{
    position: absolute; bottom: 15px; right: 22px; z-index: 201;
    font-size: 13px; color: #374151; font-variant-numeric: tabular-nums;
    letter-spacing: 0.05em; pointer-events: none;
  }}

  @media print {{
    @page {{ size: 15in 8.4375in; margin: 0; }}
    html, body {{ height: auto; overflow: visible; background: #111;
      margin: 0; padding: 0;
      -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .deck {{ position: static; width: 1440px; height: auto;
      transform: none !important; left: 0 !important; top: 0 !important; }}
    .slide {{ position: static; display: flex !important; flex-direction: column;
      width: 1440px; height: 810px; overflow: hidden;
      page-break-after: always; break-after: page;
      page-break-inside: avoid; break-inside: avoid; }}
    .slide-cover {{ flex: 1; min-height: 0; position: relative; overflow: hidden; }}
    .cover-inner {{ flex: 1; min-height: 0; }}
    .slide-main {{ overflow: hidden; flex: 1; min-height: 0; }}
    .slide-main:not(.table-main) {{ display: flex; flex-direction: column; justify-content: center; }}
    .nav-arrow, .nav-bar, .slide-counter {{ display: none !important; }}

    /* AGP slide: WeasyPrint can't fully resolve nested CSS Grid 1fr rows
       or flex-grow inside content-sized parents. Pin row heights so the
       card body has real content height; let cards size to that content. */
    .agp-main {{
      display: flex !important;
      flex-direction: column !important;
      overflow: hidden !important;
    }}
    .agp-grid {{
      flex: 0 0 auto !important;
      width: 100% !important;
      align-self: stretch !important;
      overflow: hidden !important;
      min-height: 0 !important;
    }}
    .agp-rows {{
      flex: 0 0 auto !important;
      overflow: hidden !important;
    }}
    /* Single-row card layouts (1-3 cards): compact rows ~ HTML appearance */
    .agp-grid[data-count="1"] .agp-row,
    .agp-grid[data-count="2"] .agp-row,
    .agp-grid[data-count="3"] .agp-row,
    .agp-grid[data-count="1"] .agp-rm-row,
    .agp-grid[data-count="2"] .agp-rm-row,
    .agp-grid[data-count="3"] .agp-rm-row {{
      flex: 0 0 auto !important;
      height: 100px !important;
    }}
    /* Two-row card layouts (4, 5 cards): shorter rows so cards fit vertically */
    .agp-grid[data-count="4"] .agp-row,
    .agp-grid[data-count="5"] .agp-row,
    .agp-grid[data-count="4"] .agp-rm-row,
    .agp-grid[data-count="5"] .agp-rm-row {{
      flex: 0 0 auto !important;
      height: 65px !important;
    }}
  }}
</style>
</head>
<body>
<div class="deck" id="deck">
{slides_html}
<button class="nav-arrow nav-prev" id="btnPrev" aria-label="Previous">&#8592;</button>
<button class="nav-arrow nav-next" id="btnNext" aria-label="Next">&#8594;</button>
<div class="nav-bar">{dots_html}</div>
<div class="slide-counter" id="counter">1 / {total}</div>
</div>
<script>
(function() {{
  var BASE_W = 1440, BASE_H = 810;
  var deck = document.getElementById('deck');
  function scaleDeck() {{
    var scale = Math.min(window.innerWidth / BASE_W, window.innerHeight / BASE_H);
    var ox = (window.innerWidth  - BASE_W * scale) / 2;
    var oy = (window.innerHeight - BASE_H * scale) / 2;
    deck.style.transform = 'scale(' + scale + ')';
    deck.style.left = ox + 'px';
    deck.style.top  = oy + 'px';
  }}
  scaleDeck();
  window.addEventListener('resize', scaleDeck);

  var slides  = document.querySelectorAll('.slide');
  var dots    = document.querySelectorAll('.dot');
  var counter = document.getElementById('counter');
  var cur     = 0;
  function go(n) {{
    slides[cur].classList.remove('active');
    dots[cur].classList.remove('dot-active');
    cur = ((n % slides.length) + slides.length) % slides.length;
    slides[cur].classList.add('active');
    dots[cur].classList.add('dot-active');
    counter.textContent = (cur + 1) + ' / ' + slides.length;
  }}
  document.getElementById('btnPrev').addEventListener('click', function() {{ go(cur - 1); }});
  document.getElementById('btnNext').addEventListener('click', function() {{ go(cur + 1); }});
  dots.forEach(function(d, i) {{ d.addEventListener('click', function() {{ go(i); }}); }});
  document.addEventListener('keydown', function(e) {{
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown')  go(cur + 1);
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') go(cur - 1);
  }});
  var touchX = null;
  document.addEventListener('touchstart', function(e) {{ touchX = e.touches[0].clientX; }}, {{passive: true}});
  document.addEventListener('touchend', function(e) {{
    if (touchX === null) return;
    var dx = e.changedTouches[0].clientX - touchX;
    if (Math.abs(dx) > 40) go(dx < 0 ? cur + 1 : cur - 1);
    touchX = null;
  }}, {{passive: true}});
}})();
</script>
</body>
</html>'''


def _barlow_font_css() -> str:
    import base64, pathlib
    fonts_dir = pathlib.Path(__file__).parent / 'fonts'
    rules = []
    for weight in (300, 400, 600, 700):
        path = fonts_dir / f'barlow_{weight}.woff2'
        b64 = base64.b64encode(path.read_bytes()).decode()
        rules.append(
            f"@font-face{{font-family:'Barlow Semi Condensed';font-style:normal;"
            f"font-weight:{weight};font-display:swap;"
            f"src:url('data:font/woff2;base64,{b64}') format('woff2');}}"
        )
    return '<style>' + ''.join(rules) + '</style>'


def _diagnose_cpu():
    # TEMP DIAGNOSTIC (round 5): --no-zygote and --single-process both
    # failed to move the crash address at all (identical address under gdb,
    # which disables ASLR - meaning the crash point is independent of
    # process-model flags entirely). That rules out multi-process
    # coordination as the cause. Leading remaining theory: Chromium/V8
    # requires a CPU instruction set this host doesn't have, and aborts via
    # the same silent trap-based crash mechanism. Check /proc/cpuinfo
    # directly instead of guessing more flags.
    import sys
    try:
        with open('/proc/cpuinfo') as f:
            cpuinfo = f.read()
        first_flags_line = next((l for l in cpuinfo.splitlines() if l.startswith('flags')), '')
        model_line = next((l for l in cpuinfo.splitlines() if l.startswith('model name')), '')
        present_flags = set(first_flags_line.split(':', 1)[-1].split())
        watch = ['sse4_2', 'popcnt', 'avx', 'avx2', 'bmi1', 'bmi2', 'fma', 'lzcnt', 'movbe', 'f16c']
        status = {flag: (flag in present_flags) for flag in watch}
        print(f"[diag5] {model_line.strip()}", file=sys.stdout, flush=True)
        print(f"[diag5] cpu feature check: {status}", file=sys.stdout, flush=True)
    except Exception as e:
        print(f"[diag5] could not read /proc/cpuinfo: {e!r}", file=sys.stdout, flush=True)


def make_driver():
    import shutil
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    flags = [
        '--headless', '--no-sandbox', '--disable-dev-shm-usage',
        '--disable-gpu', '--window-size=1440,810', '--no-zygote',
        '--single-process',
    ]
    options = Options()
    for flag in flags:
        options.add_argument(flag)
    chromium = shutil.which('chromium') or shutil.which('chromium-browser')
    if chromium:
        options.binary_location = chromium
        _diagnose_cpu()
    chromedriver = shutil.which('chromedriver')
    service = Service(executable_path=chromedriver) if chromedriver else Service()
    return webdriver.Chrome(service=service, options=options)


def html_to_pdf(html_str: str, driver=None) -> bytes:
    import os, base64, tempfile
    close_after = driver is None
    if driver is None:
        driver = make_driver()
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode='w', suffix='.html', delete=False, dir='/tmp'
        ) as f:
            f.write(html_str)
            tmp_path = f.name
        driver.get(f'file://{tmp_path}')
        result = driver.execute_cdp_cmd('Page.printToPDF', {
            'printBackground': True,
            'preferCSSPageSize': True,
        })
        return base64.b64decode(result['data'])
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        if close_after:
            driver.quit()


# ── Core processor ────────────────────────────────────────────────────────────
def process_csv(csv_bytes, csv_filename, config_df, prev_data=None, client_history=None):
    """
    Returns (client_name, html_str, totals_dict).
    Raises ValueError/FileNotFoundError on bad input.
    """
    config_row  = find_client_config(csv_filename, config_df)
    client_name = str(config_row['Client Name'])

    df = pd.read_csv(io.BytesIO(csv_bytes))
    conv_cols = find_conv_cols(df.columns, config_row)
    if not conv_cols:
        raise ValueError(
            f"No conversion columns found for '{client_name}'. "
            "Check Conversion Column 1–5 in client_config.xlsx."
        )

    st_col      = find_col_by_key(df.columns, config_row, 'Site Traffic')
    rev_col     = find_col_by_key(df.columns, config_row, 'Revenue')
    has_revenue = rev_col is not None
    conv_label  = ' + '.join(short_name(c) for c in conv_cols)

    numeric_cols = (
        ['Impressions', 'Clicks', 'Player Completed Views', 'Advertiser Cost (Adv Currency)']
        + conv_cols
        + ([st_col]  if st_col  else [])
        + ([rev_col] if rev_col else [])
    )
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['_conv'] = df[conv_cols].sum(axis=1)
    df['_st']   = df[st_col]  if st_col  else 0
    df['_rev']  = df[rev_col] if rev_col else 0
    df['_chan']  = df['Campaign'].apply(extract_channel)

    imp  = df['Impressions'].sum()
    clk  = df['Clicks'].sum()
    spnd = df['Advertiser Cost (Adv Currency)'].sum()
    conv = df['_conv'].sum()
    st   = df['_st'].sum()
    pcv  = df['Player Completed Views'].sum()
    rev  = df['_rev'].sum()
    uh   = df['Unique Households'].sum() if 'Unique Households' in df.columns else 0

    _uh_agg = {'uh': ('Unique Households', 'sum')} if 'Unique Households' in df.columns else {}
    grp_chan = df.groupby('_chan').agg(**{
        'imp':  ('Impressions', 'sum'),
        'clk':  ('Clicks', 'sum'),
        'spnd': ('Advertiser Cost (Adv Currency)', 'sum'),
        'conv': ('_conv', 'sum'),
        'st':   ('_st', 'sum'),
        'pcv':  ('Player Completed Views', 'sum'),
        'rev':  ('_rev', 'sum'),
        **_uh_agg,
    }).reset_index()
    if 'uh' not in grp_chan.columns:
        grp_chan['uh'] = 0

    grp_cre = df.groupby('Creative').agg(
        imp=('Impressions', 'sum'), clk=('Clicks', 'sum'),
        spnd=('Advertiser Cost (Adv Currency)', 'sum'),
        conv=('_conv', 'sum'), st=('_st', 'sum'),
        pcv=('Player Completed Views', 'sum'), rev=('_rev', 'sum')
    ).reset_index().sort_values('conv', ascending=False)

    _SITE_CHANNELS = {'Display', 'Video'}
    df_sites = df[
        df['_chan'].isin(_SITE_CHANNELS) &
        ~df['Site'].str.contains(r'\[tail aggregate\]', case=False, na=False)
    ]
    grp_site = df_sites.groupby('Site').agg(
        imp=('Impressions', 'sum'),
        spnd=('Advertiser Cost (Adv Currency)', 'sum'),
        clk=('Clicks', 'sum'),
        st=('_st', 'sum'),
        rev=('_rev', 'sum')
    ).reset_index().sort_values('imp', ascending=False).head(10)

    grp_adgroup = None
    if 'Ad Group' in df.columns:
        df['_ag_label'] = df['Ad Group'].apply(extract_adgroup_label)
        grp_adgroup = df.groupby(['_chan', '_ag_label']).agg(
            imp=('Impressions',                       'sum'),
            spnd=('Advertiser Cost (Adv Currency)',   'sum'),
            conv=('_conv',                            'sum'),
            st=('_st',                                'sum'),
            pcv=('Player Completed Views',            'sum'),
        ).reset_index()

    upsell_data = _calc_upsell(client_history, extract_report_month(csv_filename),
                               float(spnd), float(conv))

    html_str = generate_html(
        csv_filename, client_name, conv_label, has_revenue,
        (imp, clk, spnd, conv, st, pcv, rev, uh),
        grp_chan, grp_cre, grp_site, prev_data, upsell_data, client_history,
        grp_adgroup=grp_adgroup,
    )

    totals_dict = {
        'Client':            client_name,
        'Month':             extract_report_month(csv_filename),
        'Impressions':       float(imp),
        'Clicks':            float(clk),
        'Spend':             float(spnd),
        'Conversions':       float(conv),
        'Revenue':           float(rev),
        'Site Traffic':      float(st),
        'Upsell_Triggered':  'TRUE' if upsell_data else 'FALSE',
    }
    return client_name, html_str, totals_dict


# ── CLI entry point ────────────────────────────────────────────────────────────
def main(csv_path):
    csv_path = os.path.abspath(csv_path)
    folder   = os.path.dirname(csv_path)

    config_path = os.path.join(folder, 'client_config.xlsx')
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"client_config.xlsx not found.\nExpected: {config_path}")

    config_df = pd.read_excel(config_path, sheet_name='Config')
    with open(csv_path, 'rb') as f:
        csv_bytes = f.read()

    client_name, html_str, _ = process_csv(
        csv_bytes, os.path.basename(csv_path), config_df
    )
    print(f"Matched client: {client_name}")

    html_path = os.path.join(folder, f'{client_name}_report.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_str)
    print(f"HTML saved  : {html_path}")


def _run_tests():
    """Generate two test reports to ~/Desktop/Reporting/:
    - high_conv_test_report.html  → uses real CSV, CPA should appear everywhere
    - low_conv_test_report.html   → same CSV but conversions zeroed to 5 total,
                                    Cost Per Visit should appear instead of CPA
    """
    import io as _io

    CONFIG_PATH = '/Users/samhurdley/Desktop/Reporting/client_config.xlsx'
    CSV_PATH    = ('/Users/samhurdley/Desktop/Reporting/'
                   'Meatbox _ LastMonth _ Sam - Omnichannel CSV Report'
                   '_2026-03-01_2026-04-01_190090886.csv')
    OUT_DIR     = '/Users/samhurdley/Desktop/Reporting'

    config_df = pd.read_excel(CONFIG_PATH, sheet_name='Config')

    # ── Test 1: High-conv (real data, CPA) ────────────────────────────────────
    print('Running HIGH-CONV test…')
    with open(CSV_PATH, 'rb') as f:
        csv_bytes_real = f.read()

    # Dummy history for high-conv: historical CPAs higher than current so upsell fires.
    # Real data: ~1617 conv, $7500 spend → CPA ≈ $4.64
    # History CPAs: $6.36 + $6.19 → avg ≈ $6.28 → improvement ≈ 26% (≥15% threshold)
    high_history = pd.DataFrame([
        {'Client': 'Meatbox', 'Month': 'January 2026',  'Impressions': 4_800_000, 'Clicks': 9_200,
         'Spend': 7_000, 'Conversions': 1_100, 'Revenue': 0, 'Site Traffic': 29_000, 'Upsell_Triggered': 'FALSE'},
        {'Client': 'Meatbox', 'Month': 'February 2026', 'Impressions': 5_300_000, 'Clicks': 10_100,
         'Spend': 6_500, 'Conversions': 1_050, 'Revenue': 0, 'Site Traffic': 32_000, 'Upsell_Triggered': 'FALSE'},
    ])
    prev_data_high = {
        'Impressions': 5_300_000, 'Clicks': 10_100,
        'Spend': 6_500, 'Conversions': 1_050, 'Revenue': 0, 'Site Traffic': 32_000,
    }
    client_name, html_high, _ = process_csv(
        csv_bytes_real, os.path.basename(CSV_PATH), config_df,
        prev_data=prev_data_high, client_history=high_history,
    )
    out_high = os.path.join(OUT_DIR, 'high_conv_test_report.html')
    with open(out_high, 'w', encoding='utf-8') as f:
        f.write(html_high)
    assert 'Cost Per Visit' not in html_high.split('<div class="mc-label">')[1].split('</div>')[0], \
        'HIGH-CONV: overview card should say CPA, not Cost Per Visit'
    assert '>CPA<' in html_high, 'HIGH-CONV: MoM table should have CPA header'
    assert 'Site Traffic — Monthly Trend' not in html_high, \
        'HIGH-CONV: sparkline must not switch to Site Traffic when conversions are high'
    assert 'Optimisation Opportunity' in html_high, 'HIGH-CONV: upsell block should appear'
    print(f'  ✓ Saved: {out_high}')

    # ── Test 2: Low-conv (5 conversions on Display rows, CPV) ─────────────────
    print('Running LOW-CONV test…')
    raw_df    = pd.read_csv(_io.BytesIO(csv_bytes_real))
    config_r  = find_client_config(os.path.basename(CSV_PATH), config_df)
    conv_cols = find_conv_cols(raw_df.columns.tolist(), config_r)
    if conv_cols:
        for col in conv_cols:
            raw_df[col] = pd.to_numeric(raw_df[col], errors='coerce').fillna(0)
            raw_df[col] = 0.0
        display_mask = raw_df['Campaign'].str.contains('Display', case=False, na=False)
        display_idx  = raw_df[display_mask].index[:5]
        target_idx   = display_idx if len(display_idx) else raw_df.index[:5]
        raw_df.loc[target_idx, conv_cols[0]] = 1.0

    modified_csv = raw_df.to_csv(index=False).encode('utf-8')

    # Dummy history for low-conv: historical CPAs higher than current so upsell fires.
    # Current: 5 conv, $7500 spend → CPA = $1500
    # History: $2000 + $1700 → avg $1850 → improvement ≈ 19% (≥15% threshold)
    low_history = pd.DataFrame([
        {'Client': 'Meatbox', 'Month': 'January 2026',  'Impressions': 4_200_000, 'Clicks': 8_100,
         'Spend': 6_000, 'Conversions': 3, 'Revenue': 0, 'Site Traffic': 28_000, 'Upsell_Triggered': 'FALSE'},
        {'Client': 'Meatbox', 'Month': 'February 2026', 'Impressions': 5_100_000, 'Clicks': 9_400,
         'Spend': 6_800, 'Conversions': 4, 'Revenue': 0, 'Site Traffic': 31_500, 'Upsell_Triggered': 'FALSE'},
    ])
    prev_data_low = {
        'Impressions': 5_100_000, 'Clicks': 9_400,
        'Spend': 6_800, 'Conversions': 4, 'Revenue': 0, 'Site Traffic': 31_500,
    }
    client_name2, html_low, _ = process_csv(
        modified_csv, os.path.basename(CSV_PATH), config_df,
        prev_data=prev_data_low, client_history=low_history,
    )
    out_low = os.path.join(OUT_DIR, 'low_conv_test_report.html')
    with open(out_low, 'w', encoding='utf-8') as f:
        f.write(html_low)
    assert 'Cost Per Visit' in html_low, 'LOW-CONV: overview card should show Cost Per Visit'
    assert '>Cost Per Visit<' in html_low, 'LOW-CONV: MoM table header should say Cost Per Visit'
    # Sparkline only renders with ≥2 history months; confirm CPA sparkline is NOT active
    assert 'Conversions — Monthly Trend' not in html_low, \
        'LOW-CONV: sparkline must not say Conversions when in low-conv mode'
    print(f'  ✓ Saved: {out_low}')

    print('\nAll tests passed ✓')
    print(f'Open the files in {OUT_DIR} to review.')


if __name__ == '__main__':
    if len(sys.argv) == 2 and sys.argv[1] == '--test':
        try:
            _run_tests()
        except Exception as e:
            import traceback
            traceback.print_exc()
            sys.exit(1)
    elif len(sys.argv) != 2:
        print("Usage:  python process_report.py path/to/report.csv")
        print("        python process_report.py --test   (run high/low-conv tests)")
        sys.exit(1)
    else:
        try:
            main(sys.argv[1])
        except Exception as e:
            print(f"\nError: {e}")
            sys.exit(1)
