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
                'Conversion Column 3', 'Conversion Column 4']:
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

def _prev_month_label(report_month):
    try:
        dt = datetime.strptime(report_month, '%B %Y')
        prev = dt.replace(year=dt.year - 1, month=12) if dt.month == 1 \
               else dt.replace(month=dt.month - 1)
        return prev.strftime('%B %Y')
    except Exception:
        return 'Previous Month'

def _h(s):   return _html.escape(str(s))
def _n(v):   return f"{v:,.0f}"
def _m(v):   return f"${v:,.2f}"
def _p(v):   return f"{v:.2f}%"
def _rx(v):  return f"{v:.2f}x"

_KPI_CYCLE = ['#EF426F', '#5BC2E7']

def _delta_badge(curr_val, prev_val, invert=False):
    if not prev_val:
        return ''
    pct = (curr_val - prev_val) / abs(prev_val) * 100
    if abs(pct) < 0.05:
        return '<span style="color:#7A939C;font-size:10px;margin-left:5px;font-weight:600">no change</span>'
    is_good = pct > 0 if not invert else pct < 0
    color = '#5BC2E7' if is_good else '#EF426F'
    arrow = '▲' if pct > 0 else '▼'
    return f'<span style="color:{color};font-size:10px;margin-left:5px;font-weight:600">{arrow}&nbsp;{abs(pct):.1f}%</span>'

def _avg_badge(value, benchmark):
    color = '#5BC2E7' if value >= benchmark else '#EF426F'
    arrow = '▲' if value >= benchmark else '▼'
    return f'<span style="color:{color};font-size:11px;margin-left:4px">{arrow}</span>'

def _kpi_card(label, value, color='#EF426F', trend=None, invert=False):
    trend_html = ''
    if trend is not None:
        if abs(trend) < 0.05:
            trend_html = '<div class="kpi-trend" style="color:#7A939C">no change vs last month</div>'
        else:
            is_good = trend > 0 if not invert else trend < 0
            t_color = '#5BC2E7' if is_good else '#EF426F'
            arrow = '▲' if trend > 0 else '▼'
            trend_html = f'<div class="kpi-trend" style="color:{t_color}">{arrow} {abs(trend):.1f}% vs last month</div>'
    return f'''
        <div class="kpi-card" style="border-left:4px solid {color}">
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


def generate_html(csv_path, client_name, conv_label, has_revenue,
                  totals, grp_chan, grp_cre, grp_site, prev_data=None):
    imp, clk, spnd, conv, st, pcv, rev = totals
    _, cpc, cpm, cvr, _ = calc_metrics(imp, clk, spnd, conv, pcv)
    ctr_overall = clk / imp * 100 if imp else 0
    report_month = extract_report_month(csv_path)
    title = f"{client_name} Omnichannel Report — {report_month}"

    NA = '<span style="color:#7A939C;font-size:11px">N/A</span>'

    if prev_data:
        pi    = float(prev_data.get('Impressions', 0) or 0)
        pk    = float(prev_data.get('Clicks', 0) or 0)
        ps    = float(prev_data.get('Spend', 0) or 0)
        pc    = float(prev_data.get('Conversions', 0) or 0)
        pr    = float(prev_data.get('Revenue', 0) or 0)
        p_st  = float(prev_data.get('Site Traffic', 0) or 0)
        prev_ctr  = pk / pi * 100  if pi else 0
        prev_cpc  = ps / pk        if pk else 0
        prev_cpm  = ps / pi * 1000 if pi else 0
        prev_cvr  = pc / pi * 100  if pi else 0
        prev_roas = roas(pr, ps)
    else:
        pi = pk = ps = pc = pr = p_st = prev_ctr = prev_cpc = prev_cpm = prev_cvr = prev_roas = 0

    def _t(curr, prev_val):
        if not prev_data or not prev_val:
            return None
        return (curr - prev_val) / abs(prev_val) * 100

    kpi_defs = [
        ('Total Impressions',  _n(imp),         _t(imp, pi),               False),
        ('Total Clicks',       _n(clk),         _t(clk, pk),               False),
        ('Total Spend',        _m(spnd),        _t(spnd, ps),              False),
        ('CTR',                _p(ctr_overall), _t(ctr_overall, prev_ctr), False),
        ('eCPC',               _m(cpc),         _t(cpc, prev_cpc),         True),
        ('eCPM',               _m(cpm),         _t(cpm, prev_cpm),         True),
        (f'Conversions ({conv_label})', _n(conv), _t(conv, pc),            False),
        ('Conversion Rate',    _p(cvr),         _t(cvr, prev_cvr),         False),
    ]
    if has_revenue:
        kpi_defs += [
            ('Attributed Revenue', _m(rev),              _t(rev, pr),                    False),
            ('ROAS',               _rx(roas(rev, spnd)), _t(roas(rev, spnd), prev_roas), False),
        ]
    kpi_cards = ''.join(
        _kpi_card(lbl, val, _KPI_CYCLE[i % len(_KPI_CYCLE)], trend, inv)
        for i, (lbl, val, trend, inv) in enumerate(kpi_defs)
    )
    kpi_cols = 5 if has_revenue else 4

    # ── MoM comparison table ─────────────────────────────────────────────
    mom_section = ''
    if prev_data:
        prev_month_label = _prev_month_label(report_month)
        curr_cpa = spnd / conv if conv else None
        prev_cpa = ps / pc if pc else None
        NA_mom = '<span style="color:#7A939C;font-size:11px">N/A</span>'

        def _mc(main, curr, prev):
            badge = _delta_badge(curr, prev) if prev else ''
            return f'<span>{main}{badge}</span>'

        prev_cpa_str = _m(prev_cpa) if prev_cpa else NA_mom
        curr_cpa_str = (
            f'<span>{_m(curr_cpa)}{_delta_badge(curr_cpa, prev_cpa, invert=True)}</span>'
            if curr_cpa else NA_mom
        )
        mom_head = _th(['Month', 'Spend', 'Site Traffic', 'Conversions', 'CPA'], right_from=1)
        prev_row = _td_row([prev_month_label, _m(ps), _n(p_st), _n(pc), prev_cpa_str])
        curr_row = _td_row([
            report_month,
            _mc(_m(spnd), spnd, ps),
            _mc(_n(st), st, p_st),
            _mc(_n(conv), conv, pc),
            curr_cpa_str,
        ])
        mom_section = f'''
  <div class="section-label">Month-on-Month Performance</div>
  <div class="table-wrap">
    <table>
      <thead>{mom_head}</thead>
      <tbody>{prev_row}{curr_row}</tbody>
    </table>
  </div>'''

    NO_CLICKS = {'CTV', 'Audio'}
    NO_COMP   = {'Display'}

    # ── Table 1: Channel Performance (Outcomes) ─────────────────────────
    out_hdrs = ['Channel', 'Spend %', 'Impressions', 'Attributed Site Traffic',
                'Conversions', 'Conversion Rate', 'Cost per Conv (CPA)']
    if has_revenue:
        out_hdrs += ['Revenue', 'ROAS']
    out_head = _th(out_hdrs)
    out_body = ''
    for _, r in grp_chan.iterrows():
        _, _, _, cvr_, _ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        ch = str(r['_chan'])
        spnd_pct = _p(r.spnd / spnd * 100) if spnd else _p(0)
        cpa = _m(r.spnd / r.conv) if r.conv else NA
        cells = [ch, spnd_pct, _n(r.imp), _n(r.st), _n(r.conv), _p(cvr_), cpa]
        if has_revenue:
            cells += [_m(r.rev), _rx(roas(r.rev, r.spnd))]
        out_body += _td_row(cells)

    # ── Table 2: Channel Engagement (Efficiency) ──────────────────────────
    _CTR_BENCH_CH  = {'Display', 'Video'}
    _COMP_HIGH_CH  = {'CTV', 'Audio'}
    _COMP_VIDEO_CH = {'Video'}

    eng_hdrs = ['Channel', 'Impressions', 'Clicks', 'CTR', 'eCPC', 'eCPM', 'Completion Rate']
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

        cells = [
            ch, _n(r.imp),
            NA if ch in NO_CLICKS else _n(r.clk),
            ctr_cell,
            NA if ch in NO_CLICKS else _m(cpc_),
            _m(cpm_),
            comp_cell,
        ]
        eng_body += _td_row(cells)

    top10_cre = grp_cre.nlargest(10, 'st')
    cre_hdrs = ['Creative', 'Impressions', 'Spend', 'CTR', 'eCPC',
                'Conversions', 'Conversion Rate', 'Completion Rate', 'Attributed Site Traffic']
    if has_revenue:
        cre_hdrs += ['Attributed Revenue', 'ROAS']
    cre_head = _th(cre_hdrs)
    cre_body = ''
    for i, (_, r) in enumerate(top10_cre.iterrows()):
        ctr_, cpc_, cpm_, cvr_, comp_ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        comp_v = NA if comp_ == 0 else _p(comp_)
        cells = [r['Creative'], _n(r.imp), _m(r.spnd), _p(ctr_), _m(cpc_),
                 _n(r.conv), _p(cvr_), comp_v, _n(r.st)]
        if has_revenue:
            cells += [_m(r.rev), _rx(roas(r.rev, r.spnd))]
        cre_body += _td_row(cells)

    site_hdrs = ['Site', 'Impressions', 'Spend', 'CPM', 'Attributed Site Traffic']
    if has_revenue:
        site_hdrs += ['Attributed Revenue', 'ROAS']
    site_head = _th(site_hdrs)
    site_body = ''
    for _, r in grp_site.iterrows():
        cpm_ = r.spnd / r.imp * 1000 if r.imp else 0
        cells = [r['Site'], _n(r.imp), _m(r.spnd), _m(cpm_), _n(r.st)]
        if has_revenue:
            cells += [_m(r.rev), _rx(roas(r.rev, r.spnd))]
        site_body += _td_row(cells)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_h(title)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Semi+Condensed:wght@400;600;700&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Barlow Semi Condensed', sans-serif;
    background: #22326E;
    color: #FFFFFF;
    font-size: 15px;
    line-height: 1.5;
  }}

  .report-header {{
    background: #22326E;
    border-bottom: 3px solid #EF426F;
    padding: 36px 48px 28px;
    display: flex;
    align-items: center;
    gap: 36px;
  }}
  .header-logo {{
    height: 44px;
    width: auto;
    flex-shrink: 0;
    filter: brightness(0) invert(1);
  }}
  .header-text h1 {{
    font-size: 24px;
    font-weight: 700;
    color: #FFFFFF;
    letter-spacing: -0.2px;
    line-height: 1.2;
  }}
  .header-text .subtitle {{
    font-size: 13px;
    color: #7A939C;
    margin-top: 5px;
  }}

  .content {{
    max-width: 1280px;
    margin: 0 auto;
    padding: 40px 48px 72px;
  }}

  .section-label {{
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: #5BC2E7;
    margin: 48px 0 16px;
  }}
  .section-label:first-child {{ margin-top: 0; }}

  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat({kpi_cols}, 1fr);
    gap: 16px;
  }}
  .kpi-card {{
    background: #425563;
    border-radius: 10px;
    padding: 24px 22px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
  }}
  .kpi-value {{
    font-size: 28px;
    font-weight: 700;
    color: #FFFFFF;
    font-variant-numeric: tabular-nums;
    line-height: 1.1;
  }}
  .kpi-label {{
    font-size: 11px;
    color: #7A939C;
    margin-top: 8px;
    text-transform: uppercase;
    letter-spacing: 0.09em;
    font-weight: 600;
  }}
  .kpi-trend {{
    font-size: 11px;
    margin-top: 6px;
    font-weight: 600;
    letter-spacing: 0.05em;
  }}

  .table-wrap {{
    background: #425563;
    border-radius: 10px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    overflow-x: auto;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 12.5px;
    min-width: 600px;
  }}
  thead tr {{
    background: #2c3e53;
    border-top: 2px solid #EF426F;
  }}
  th {{
    color: #FFFFFF;
    padding: 12px 14px;
    font-weight: 600;
    font-size: 10.5px;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    white-space: normal;
    min-width: 80px;
    vertical-align: bottom;
    line-height: 1.3;
    text-align: right;
  }}
  th:first-child {{
    min-width: 160px;
    text-align: left;
  }}
  td {{
    padding: 10px 14px;
    border-bottom: 1px solid rgba(255,255,255,0.07);
    color: #FFFFFF;
    vertical-align: middle;
    text-align: left;
  }}
  td:first-child {{
    word-break: break-word;
    overflow-wrap: anywhere;
    max-width: 260px;
    line-height: 1.4;
  }}
  td.right {{
    text-align: right;
    white-space: nowrap;
    font-variant-numeric: tabular-nums;
  }}
  tr:last-child td {{ border-bottom: none; }}
  tbody tr:nth-child(even) td {{ background: rgba(255,255,255,0.04); }}
  tbody tr.highlight td {{ background: rgba(239,66,111,0.18); font-weight: 600; }}
  tbody tr:hover td {{ background: rgba(91,194,231,0.12); }}

  .conv-note {{
    display: inline-block;
    margin-top: 16px;
    padding: 9px 16px;
    background: rgba(255,255,255,0.06);
    border-left: 3px solid #5BC2E7;
    border-radius: 0 6px 6px 0;
    font-size: 12px;
    color: #7A939C;
  }}
  .conv-note strong {{
    color: #FFFFFF;
    font-weight: 600;
  }}

  .footer {{
    text-align: center;
    font-size: 11px;
    color: #7A939C;
    margin-top: 56px;
    padding-top: 20px;
    border-top: 1px solid rgba(255,255,255,0.1);
  }}

  @media print {{
    body {{ background: #fff; color: #1a1a1a; }}
    .report-header {{
      background: #22326E; border-bottom-color: #EF426F;
      -webkit-print-color-adjust: exact; print-color-adjust: exact;
    }}
    .header-text h1 {{ color: #fff; }}
    .header-text .subtitle {{ color: #c0cfe0; }}
    .header-logo {{ filter: brightness(0) invert(1); }}
    .kpi-card {{
      background: #f2f4f8; box-shadow: none; border: 1px solid #dde2ef;
      -webkit-print-color-adjust: exact; print-color-adjust: exact;
    }}
    .kpi-value {{ color: #1a1a1a; }}
    .kpi-label {{ color: #6b7a99; }}
    .table-wrap {{ background: #fff; box-shadow: none; border: 1px solid #dde2ef;
      overflow-x: visible; page-break-inside: avoid; }}
    thead tr {{
      background: #22326E; border-top-color: #EF426F;
      -webkit-print-color-adjust: exact; print-color-adjust: exact;
    }}
    td {{ color: #1a1a1a; border-bottom-color: #eef0f5; }}
    td.right {{ color: #1a1a1a; }}
    tbody tr:nth-child(even) td {{ background: #f7f8fc; }}
    tbody tr.highlight td {{ background: #fef2f5; }}
    tbody tr:hover td {{ background: transparent; }}
    .conv-note {{ background: #f2f4f8; color: #6b7a99; border-left-color: #5BC2E7; }}
    .conv-note strong {{ color: #1a1a1a; }}
    .footer {{ color: #9aa3bc; border-top-color: #e4e7f0; }}
    .section-label {{ color: #22326E; }}
    .kpi-grid {{ grid-template-columns: repeat({kpi_cols}, 1fr); }}
    table {{ min-width: unset; font-size: 9px; }}
    th, td {{ padding: 6px 8px; }}
    td:first-child {{ max-width: 180px; }}
  }}
</style>
</head>
<body>

<div class="report-header">
  <img class="header-logo" src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">
  <div class="header-text">
    <h1>{_h(title)}</h1>
    <div class="subtitle">{_h(extract_date_range(csv_path))}</div>
  </div>
</div>

<div class="content">

  <div class="section-label">Key Performance Indicators</div>
  <div class="kpi-grid">
    {kpi_cards}
  </div>
  <div class="conv-note">Conversions tracked: <strong>{_h(conv_label)}</strong></div>
{mom_section}
  <div class="section-label">Channel Performance — Outcomes</div>
  <div class="table-wrap">
    <table>
      <thead>{out_head}</thead>
      <tbody>{out_body}</tbody>
    </table>
  </div>

  <div class="section-label">Channel Engagement — Efficiency</div>
  <div class="table-wrap">
    <table>
      <thead>{eng_head}</thead>
      <tbody>{eng_body}</tbody>
    </table>
  </div>
  <div class="conv-note">
    <span style="color:#5BC2E7;font-weight:700">▲</span> = Above industry average &nbsp;·&nbsp;
    <span style="color:#EF426F;font-weight:700">▼</span> = Below industry average<br>
    CTR (Display, Video): <strong>0.15%</strong> &nbsp;·&nbsp;
    Completion Rate (CTV, Audio): <strong>95%</strong> &nbsp;·&nbsp;
    Completion Rate (Video): <strong>50%</strong>
  </div>

  <div class="section-label">Top 10 Creatives by Attributed Site Traffic</div>
  <div class="table-wrap">
    <table>
      <thead>{cre_head}</thead>
      <tbody>{cre_body}</tbody>
    </table>
  </div>

  <div class="section-label">Top 10 Sites by Attributed Site Traffic</div>
  <div class="table-wrap">
    <table>
      <thead>{site_head}</thead>
      <tbody>{site_body}</tbody>
    </table>
  </div>

  <div class="footer">Confidential — prepared for {_h(client_name)}</div>

</div>
</body>
</html>'''


# ── Core processor ────────────────────────────────────────────────────────────
def process_csv(csv_bytes, csv_filename, config_df, prev_data=None):
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
            "Check Conversion Column 1–4 in client_config.xlsx."
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

    grp_chan = df.groupby('_chan').agg(
        imp=('Impressions', 'sum'), clk=('Clicks', 'sum'),
        spnd=('Advertiser Cost (Adv Currency)', 'sum'),
        conv=('_conv', 'sum'), st=('_st', 'sum'),
        pcv=('Player Completed Views', 'sum'), rev=('_rev', 'sum')
    ).reset_index()

    grp_cre = df.groupby('Creative').agg(
        imp=('Impressions', 'sum'), clk=('Clicks', 'sum'),
        spnd=('Advertiser Cost (Adv Currency)', 'sum'),
        conv=('_conv', 'sum'), st=('_st', 'sum'),
        pcv=('Player Completed Views', 'sum'), rev=('_rev', 'sum')
    ).reset_index().sort_values('conv', ascending=False)

    df_sites = df[~df['Site'].str.contains(r'\[tail aggregate\]', case=False, na=False)]
    grp_site = df_sites.groupby('Site').agg(
        imp=('Impressions', 'sum'),
        spnd=('Advertiser Cost (Adv Currency)', 'sum'),
        st=('_st', 'sum'),
        rev=('_rev', 'sum')
    ).reset_index().sort_values('st', ascending=False).head(10)

    html_str = generate_html(
        csv_filename, client_name, conv_label, has_revenue,
        (imp, clk, spnd, conv, st, pcv, rev),
        grp_chan, grp_cre, grp_site, prev_data
    )

    totals_dict = {
        'Client':       client_name,
        'Month':        extract_report_month(csv_filename),
        'Impressions':  float(imp),
        'Clicks':       float(clk),
        'Spend':        float(spnd),
        'Conversions':  float(conv),
        'Revenue':      float(rev),
        'Site Traffic': float(st),
    }
    return client_name, html_str, totals_dict


# ── CLI entry point ────────────────────────────────────────────────────────────
def main(csv_path):
    csv_path = os.path.abspath(csv_path)
    folder   = os.path.dirname(csv_path)

    config_path = os.path.join(folder, 'client_config.xlsx')
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"client_config.xlsx not found.\nExpected: {config_path}")

    config_df = pd.read_excel(config_path)
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


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage:  python process_report.py path/to/report.csv")
        sys.exit(1)
    try:
        main(sys.argv[1])
    except Exception as e:
        print(f"\nError: {e}")
        sys.exit(1)
