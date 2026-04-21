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

        if not benchmark_cpas:
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

def _delta_badge(curr_val, prev_val, invert=False):
    if not prev_val:
        return ''
    pct = (curr_val - prev_val) / abs(prev_val) * 100
    if abs(pct) < 0.05:
        return '<span style="color:#6B7280;font-size:10px;margin-left:5px;font-weight:600">no change</span>'
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


def generate_html(csv_path, client_name, conv_label, has_revenue,
                  totals, grp_chan, grp_cre, grp_site, prev_data=None, upsell_data=None,
                  client_history=None):
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

    # ── MoM table ─────────────────────────────────────────────────────────────
    def _mc(main, curr, prev, invert=False):
        badge = _delta_badge(curr, prev, invert=invert) if prev is not None else ''
        return f'<span>{main}{badge}</span>'

    history_rows_html = ''
    mom_insight_text  = ''
    has_mom           = False
    mom_table         = ''

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

        prev_spend = prev_cpa_hist = prev_st_hist = prev_conv_hist = None
        for _, r in hist.iterrows():
            s    = float(r.get('Spend', 0) or 0)
            c    = float(r.get('Conversions', 0) or 0)
            st_h = float(r.get('Site Traffic', 0) or 0)
            cpa_h = s / c if c else None
            cpa_str = (
                _mc(_m(cpa_h), cpa_h, prev_cpa_hist, invert=True)
                if cpa_h is not None else NA_mom
            )
            history_rows_html += _td_row([
                str(r['Month']),
                _mc(_m(s), s, prev_spend),
                _mc(_n(st_h), st_h, prev_st_hist),
                _mc(_n(c), c, prev_conv_hist),
                cpa_str,
            ])
            prev_spend, prev_st_hist, prev_conv_hist, prev_cpa_hist = s, st_h, c, cpa_h

    if history_rows_html or prev_data:
        has_mom = True
        curr_cpa      = spnd / conv if conv else None
        p_spend_badge = float(prev_data['Spend']) if prev_data else None
        p_st_badge    = float(prev_data['Site Traffic']) if prev_data else None
        p_conv_badge  = float(prev_data['Conversions']) if prev_data else None
        p_cpa_badge   = (float(prev_data['Spend']) / float(prev_data['Conversions'])
                         if prev_data and float(prev_data.get('Conversions', 0)) else None)
        curr_cpa_str = (
            _mc(_m(curr_cpa), curr_cpa, p_cpa_badge, invert=True)
            if curr_cpa is not None else NA_mom
        )
        curr_row = _td_row([
            f'<strong>{report_month}</strong>',
            _mc(_m(spnd), spnd, p_spend_badge),
            _mc(_n(st), st, p_st_badge),
            _mc(_n(conv), conv, p_conv_badge),
            curr_cpa_str,
        ])
        mom_head  = _th(['Month', 'Spend', 'Site Traffic', 'Conversions', 'CPA'], right_from=1)
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
        upsell_html  = f'''<div class="upsell-block">
      <img class="upsell-logo" src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">
      <div class="upsell-headline">Optimisation Opportunity: Efficiency Momentum</div>
      <p class="upsell-body">Your current cost-per-conversion (<strong>{curr_cpa_fmt}</strong>) is trending <strong>{impr_fmt}%</strong> below your 3-month benchmark (<strong>{avg_cpa_fmt}</strong>). We recommend a budget increase of <strong>{budget_fmt}</strong> to scale reach while conditions remain favourable.</p>
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
        engagement_insight_html = f'<div class="insight-box">Performance Lead: <strong>{_h(best_channel)}</strong> is outperforming industry benchmarks by <strong>{best_variance:.0f}%</strong>, driving premium audience engagement.</div>'

    # ── Awareness table (Video, CTV, Audio) ──────────────────────────────────
    AWARE_CH = {'Video', 'CTV', 'Audio'}
    aware_rows = grp_chan[grp_chan['_chan'].isin(AWARE_CH)]
    aw_head = _th(['Channel', 'Impressions', 'Reach', 'Frequency', 'eCPM', 'Completion Rate'])
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
    aw_lead = f'<div class="insight-box">Performance Lead: <strong>{_h(aw_best_ch)}</strong> is outperforming industry benchmarks by <strong>{aw_best_var:.0f}%</strong>, driving premium audience engagement.</div>' if aw_best_ch else ''
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
    top10_cre = grp_cre.nlargest(10, 'st')
    cre_hdrs  = ['Creative', 'Impressions', 'CTR', 'eCPC',
                 'Conversions', 'Conversion Rate', 'Completion Rate', 'Attributed Site Traffic']
    cre_head  = _th(cre_hdrs)
    cre_body  = ''
    for _, r in top10_cre.iterrows():
        ctr_, cpc_, cpm_, cvr_, comp_ = calc_metrics(r.imp, r.clk, r.spnd, r.conv, r.pcv)
        comp_v = NA if comp_ == 0 else _p(comp_)
        cre_body += _td_row([r['Creative'], _n(r.imp), _p(ctr_), _m(cpc_),
                             _n(r.conv), _p(cvr_), comp_v, _n(r.st)])

    creative_insight_html = ''
    if not top10_cre.empty and imp > 0 and st > 0:
        campaign_vr = st / imp
        grp_cre_vr  = top10_cre.copy()
        grp_cre_vr['_vr'] = grp_cre_vr.apply(
            lambda r: r['st'] / r['imp'] if r['imp'] > 0 else 0, axis=1)
        top_vr_row = grp_cre_vr.loc[grp_cre_vr['_vr'].idxmax()]
        if campaign_vr > 0:
            variance_pct = (top_vr_row['_vr'] - campaign_vr) / campaign_vr * 100
            if variance_pct > 0:
                cre_name = _h(str(top_vr_row['Creative']))
                creative_insight_html = f'<div class="insight-box">Traffic Efficiency: <strong>&#8220;{cre_name}&#8221;</strong> was your most effective asset at driving web traffic this month, achieving a Visit Rate <strong>{variance_pct:.0f}%</strong> above the campaign average.</div>'

    # ── Top 10 Sites table ────────────────────────────────────────────────────
    site_hdrs = ['Site', 'Impressions', 'CPM', 'Attributed Site Traffic']
    site_head = _th(site_hdrs)
    site_body = ''
    for _, r in grp_site.iterrows():
        cpm_ = r.spnd / r.imp * 1000 if r.imp else 0
        site_body += _td_row([r['Site'], _n(r.imp), _m(cpm_), _n(r.st)])

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
            site_insight_html = f'<div class="insight-box">Key Environments: <strong>{site_name}</strong> converted impressions to site visits at a rate <strong>{site_variance_pct:.0f}%</strong> above the campaign average — your most efficient placement this month. We have optimised the campaign to direct more of your investment here to maximise overall performance.</div>'

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

    # ── Shared logo tags ──────────────────────────────────────────────────────
    foot_logo  = f'<img class="foot-logo"  src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">'
    cover_logo = f'<img class="cover-logo" src="data:image/png;base64,{_LOGO_B64}" alt="MediaWorks">'

    # ── Overview (KPI) slide content ──────────────────────────────────────────
    def _ov_row(label, value, trend_curr, trend_prev, invert=False, first=False):
        tl  = _trend_line(trend_curr, trend_prev, invert)
        div = '' if first else '<div class="ov-divider"></div>'
        return f'''{div}<div class="ov-row">
          <span class="ov-label">{label}</span>
          <div class="ov-num-wrap"><span class="ov-num">{value}</span>{tl}</div>
        </div>'''

    reach_row = _ov_row('Reach', _n(uh), 0, 0, False) if uh else ''

    budget_chips = ''
    for _, r in grp_chan.sort_values('spnd', ascending=False).iterrows():
        if r.spnd <= 0:
            continue
        pct = r.spnd / spnd * 100 if spnd else 0
        budget_chips += f'<span class="budget-chip"><span class="budget-chip-label">{_h(str(r["_chan"]))}</span><span class="budget-chip-pct">{pct:.0f}%</span></span>'
    budget_section = f'''<div class="ov-divider" style="margin-top:8px"></div>
      <div class="budget-split-row">
        <span class="ov-label" style="font-size:11px;text-transform:uppercase;letter-spacing:0.1em;color:#5BC2E7;font-weight:700">Budget Split</span>
        <div class="budget-chips">{budget_chips}</div>
      </div>''' if budget_chips else ''

    left_card = f'''<div class="ov-card">
      <div class="ov-card-title">Overview of the Month</div>
      {_ov_row('Investment',              f"${spnd:,.0f}", spnd, ps,   False, True)}
      {_ov_row('Total Impressions',       _n(imp),         imp,  pi,   False)}
      {reach_row}
      {_ov_row('Attributed Site Traffic', _n(st),          st,   p_st, False)}
      {budget_section}
    </div>'''

    def _perf_row(label, value, trend_curr, trend_prev, invert=False, first=False):
        tl  = _trend_line(trend_curr, trend_prev, invert)
        div = '' if first else '<div class="perf-divider"></div>'
        return f'''{div}<div class="perf-row">
          <span class="perf-label">{label}</span>
          <span class="perf-num">{value}</span>{tl}
        </div>'''

    right_perf = f'''<div class="ov-perf">
      <div class="ov-perf-title">Performance</div>
      {_perf_row(f'Total Conversions ({conv_label})', _n(conv),                       conv,    pc,          False, True)}
      {_perf_row('CPA',                               _m(cpa_kpi) if conv else NA,    cpa_kpi, prev_cpa_kpi, True)}'''
    if has_revenue:
        right_perf += _perf_row('Revenue', _m(rev),             rev,            pr,         False)
        right_perf += _perf_row('ROAS',    _rx(roas(rev, spnd)), roas(rev, spnd), prev_roas, False)
    right_perf += '\n    </div>'

    conv_note = f'<div class="ov-conv-note">Conversions tracked: <strong>{_h(conv_label)}</strong></div>'

    kpi_inner = f'''<div class="overview-split">
    {left_card}
    {right_perf}
  </div>
  {conv_note}
  {upsell_html}'''

    # ── MoM slide content ─────────────────────────────────────────────────────
    if mom_insight_text:
        takeaway = f'''<div class="takeaway-card">
      <div class="takeaway-title">Key Takeaway</div>
      <p class="takeaway-body">{mom_insight_text}</p>
    </div>'''
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
        takeaway = f'''<div class="takeaway-card">
      <div class="takeaway-title">Month Summary</div>
      <p class="takeaway-body">{body}</p>
    </div>'''

    mom_inner = f'''<div class="mom-split">
    <div class="mom-left">{mom_table}</div>
    {takeaway}
  </div>'''

    # ── Table slide helper ────────────────────────────────────────────────────
    def _tslide(title, table_html, extra=''):
        nrows = table_html.count('<tr') - 1  # subtract header row
        td_pad = max(15, min(40, 110 // max(nrows, 1)))
        sparse_style = f' style="--td-pad:{td_pad}px"' if nrows < 7 else ''
        return f'''  <h2 class="slide-title">{title}</h2>
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
    slides.append(('overview', f'''  <h2 class="slide-title">Campaign Overview</h2>
  <div class="slide-main overview-main">
    {kpi_inner}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    # MoM (conditional)
    if has_mom:
        slides.append(('mom', f'''  <h2 class="slide-title">Month on Month Results</h2>
  <div class="slide-main mom-main">
    {mom_inner}
  </div>
  <div class="slide-foot">{foot_logo}</div>'''))

    # Awareness (Video, CTV, Audio) — only if any of those channels ran
    if not aware_rows.empty:
        aw_note = '<div class="bench-note"><span style="color:#5BC2E7;font-weight:700">&#9650;</span> Above benchmark &nbsp;&#183;&nbsp; <span style="color:#EF426F;font-weight:700">&#9660;</span> Below benchmark &nbsp;&#183;&nbsp; <strong>Industry Average Benchmarks:</strong> Completion (CTV, Audio) <strong>95%</strong> &nbsp;&#183;&nbsp; Completion (Video) <strong>50%</strong></div>'
        slides.append(('table', _tslide('Awareness Performance',
            f'<table><thead>{aw_head}</thead><tbody>{aw_body}</tbody></table>',
            awareness_insight_html + aw_note)))

    # Conversion (Display) — only if Display ran
    if not conv_rows.empty:
        cv_insight = '<div class="context-box">Display is built for <strong>direct response</strong> — driving clicks, site visits, and conversions. The visit rate shows how effectively exposure is turning into real website traffic.</div>'
        slides.append(('table', _tslide('Display Conversion Performance',
            f'<table><thead>{cv_head}</thead><tbody>{cv_body}</tbody></table>',
            cv_insight)))

    # Top Creatives
    slides.append(('table', _tslide('Top 10 Creatives by Site Traffic',
        f'<table><thead>{cre_head}</thead><tbody>{cre_body}</tbody></table>',
        creative_insight_html)))

    # Top Sites
    slides.append(('table', _tslide('Top 10 Sites by Site Traffic',
        f'<table><thead>{site_head}</thead><tbody>{site_body}</tbody></table>',
        site_insight_html)))

    # Glossary
    gloss_attr = '<div class="glossary-attribution"><strong style="color:#C4B5D4">Note on Attribution:</strong> Both Conversions and Site Traffic are measured as View-Through + Click-Through — users who clicked an ad immediately, plus those who saw an ad and visited later within the attribution window.</div>'
    slides.append(('table', f'''  <h2 class="slide-title">Glossary of Terms</h2>
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
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Semi+Condensed:wght@300;400;600;700&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    height: 100%; overflow: hidden;
    font-family: 'Barlow Semi Condensed', sans-serif;
    background: #111; color: #fff; font-size: 15px; line-height: 1.5;
  }}

  .deck {{ position: absolute; width: 1440px; height: 810px; overflow: hidden; transform-origin: top left; }}
  .slide {{ position: absolute; inset: 0; display: none; flex-direction: column; overflow: hidden; }}
  .slide.active {{ display: flex; }}

  /* Cover */
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
  .cover-date {{ font-size: 14px; color: #6B7280; margin-bottom: 72px; }}
  .cover-hint {{ font-size: 12px; color: #4B5563; margin-bottom: 8px; letter-spacing: 0.05em; }}
  .cover-confidential {{ font-size: 11px; color: #374151; }}

  /* Slide chrome */
  .slide-title {{
    font-size: 40px; font-weight: 700; text-align: center; letter-spacing: -0.3px;
    padding: 36px 80px 24px; flex-shrink: 0;
  }}
  .slide-main {{ flex: 1; min-height: 0; overflow-y: auto; padding: 0 72px 12px; }}
  .slide-main.table-main {{ display: flex; flex-direction: column; justify-content: center; }}
  .slide-foot {{
    height: 40px; min-height: 40px; flex-shrink: 0;
    display: flex; align-items: center; justify-content: flex-end;
    padding: 0 72px; border-top: 1px solid rgba(255,255,255,0.05);
  }}
  .foot-logo {{ height: 20px; width: auto; filter: brightness(0) invert(1); opacity: 0.4; }}
  .foot-credit {{ margin-top: 20px; font-size: 11px; color: #374151; text-align: center; }}

  /* Overview / KPI slide */
  .overview-main {{
    display: flex; flex-direction: column; justify-content: center;
    padding-top: 8px; padding-bottom: 8px;
  }}
  .overview-split {{
    display: grid; grid-template-columns: 44fr 52fr; gap: 48px; align-items: center;
  }}
  .ov-card {{
    background: #252526; border-radius: 16px; padding: 40px 44px;
    display: flex; flex-direction: column;
  }}
  .ov-card-title {{
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.14em; color: #5BC2E7; margin-bottom: 24px;
  }}
  .ov-divider {{ height: 1px; background: rgba(255,255,255,0.07); }}
  .ov-row {{
    display: flex; align-items: center; justify-content: space-between;
    gap: 16px; padding: 20px 0;
  }}
  .ov-label {{ font-size: 15px; color: #9BA3AF; max-width: 150px; line-height: 1.3; }}
  .ov-num-wrap {{ text-align: right; }}
  .ov-num {{ font-size: 42px; font-weight: 700; font-variant-numeric: tabular-nums; line-height: 1.0; display: block; }}
  .trend-line {{ font-size: 11px; font-weight: 600; letter-spacing: 0.04em; margin-top: 4px; }}
  .ov-perf {{ display: flex; flex-direction: column; padding: 4px 0; }}
  .ov-perf-title {{
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.14em; color: #5BC2E7; margin-bottom: 20px;
  }}
  .perf-divider {{ height: 1px; background: rgba(255,255,255,0.07); margin: 4px 0; }}
  .perf-row {{ padding: 20px 0; }}
  .perf-label {{ font-size: 15px; color: #9BA3AF; display: block; margin-bottom: 6px; }}
  .perf-num {{
    font-size: 72px; font-weight: 700; color: #EF426F;
    font-variant-numeric: tabular-nums; line-height: 1.0; display: block; letter-spacing: -1px;
  }}
  .ov-conv-note {{
    margin-top: 18px; font-size: 12px; color: #6B7280; display: inline-block;
    padding: 7px 14px; background: rgba(255,255,255,0.03);
    border-left: 3px solid #5BC2E7; border-radius: 0 6px 6px 0;
  }}
  .ov-conv-note strong {{ color: #fff; font-weight: 600; }}
  .budget-split-row {{ display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 14px 0 4px; }}
  .budget-chips {{ display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }}
  .budget-chip {{ display: flex; align-items: center; gap: 5px; background: rgba(255,255,255,0.06); border-radius: 6px; padding: 4px 10px; }}
  .budget-chip-label {{ font-size: 11px; color: #9BA3AF; }}
  .budget-chip-pct {{ font-size: 13px; font-weight: 700; color: #fff; }}

  /* MoM slide */
  .mom-main {{ display: flex; flex-direction: column; justify-content: center; padding-top: 8px; }}
  .mom-split {{ display: grid; grid-template-columns: 58fr 38fr; gap: 32px; align-items: start; }}
  .mom-left .table-wrap {{ border-radius: 14px; }}
  .takeaway-card {{
    background: #252526; border-radius: 14px; padding: 36px 32px;
    display: flex; flex-direction: column; justify-content: center; align-self: stretch;
  }}
  .takeaway-title {{ font-size: 26px; font-weight: 700; margin-bottom: 20px; letter-spacing: -0.1px; }}
  .takeaway-body {{ font-size: 17px; color: #D1D5DB; line-height: 1.75; font-weight: 400; }}
  .takeaway-body strong {{ color: #EF426F; }}

  /* Tables */
  .table-wrap {{ background: #252526; border-radius: 12px; overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 15px; min-width: 600px; }}
  thead tr {{ background: #671E75; }}
  th {{
    color: #fff; padding: 15px 18px; font-weight: 600; font-size: 12px;
    text-transform: uppercase; letter-spacing: 0.07em; min-width: 80px;
    vertical-align: bottom; line-height: 1.3; text-align: right;
  }}
  th:first-child {{ min-width: 160px; text-align: left; }}
  td {{
    padding: var(--td-pad, 15px) 18px; border-bottom: 1px solid rgba(255,255,255,0.05);
    color: #fff; vertical-align: middle; text-align: left;
  }}
  td:first-child {{ word-break: break-word; overflow-wrap: anywhere; max-width: 240px; line-height: 1.4; }}
  td.right {{ text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }}
  tr:last-child td {{ border-bottom: none; }}
  tbody tr:nth-child(even) td {{ background: rgba(255,255,255,0.025); }}
  tbody tr:hover td {{ background: rgba(91,194,231,0.07); }}

  /* Components */
  .na {{ color: #4B5563; font-size: 11px; }}
  .insight-box {{
    background: rgba(91,194,231,0.08);
    border: 1px solid rgba(91,194,231,0.2);
    border-left: 5px solid #5BC2E7;
    border-radius: 0 10px 10px 0; padding: 18px 24px;
    margin-top: 16px; font-size: 15px; color: #fff; line-height: 1.65;
  }}
  .insight-box strong {{ color: #fff; }}
  .context-box {{
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 8px; padding: 14px 20px;
    margin-top: 16px; font-size: 13px; color: #9BA3AF; line-height: 1.6;
  }}
  .context-box strong {{ color: #C4B5D4; }}
  .bench-note {{
    display: inline-block; margin-top: 12px; padding: 8px 14px;
    background: rgba(255,255,255,0.04); border-left: 3px solid #5BC2E7;
    border-radius: 0 6px 6px 0; font-size: 12px; color: #9BA3AF;
  }}
  .bench-note strong {{ color: #fff; font-weight: 600; }}
  .upsell-block {{
    position: relative; background: #EF426F; border-radius: 12px;
    padding: 26px 30px; margin-top: 20px; color: #fff;
  }}
  .upsell-headline {{ font-size: 16px; font-weight: 700; margin-bottom: 10px; padding-right: 80px; }}
  .upsell-body {{ font-size: 13px; line-height: 1.7; max-width: 88%; margin: 0; }}
  .upsell-logo {{
    position: absolute; top: 26px; right: 26px; height: 22px; width: auto;
    filter: brightness(0) invert(1); opacity: 0.85;
  }}
  .glossary-grid {{ display: grid; grid-template-columns: repeat(2,1fr); gap: 10px 36px; margin-bottom: 16px; }}
  .glossary-item {{
    font-size: 12px; color: #9BA3AF; line-height: 1.5; padding: 8px 12px;
    background: #252526; border-left: 2px solid #671E75; border-radius: 0 4px 4px 0;
  }}
  .glossary-term {{ font-weight: 700; color: #C4B5D4; }}
  .glossary-attribution {{
    font-size: 12px; color: #9BA3AF; padding: 10px 14px;
    background: #252526; border-radius: 6px; line-height: 1.6;
  }}

  /* Navigation */
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
    font-size: 11px; color: #374151; font-variant-numeric: tabular-nums;
    letter-spacing: 0.05em; pointer-events: none;
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

    df_sites = df[~df['Site'].str.contains(r'\[tail aggregate\]', case=False, na=False)]
    grp_site = df_sites.groupby('Site').agg(
        imp=('Impressions', 'sum'),
        spnd=('Advertiser Cost (Adv Currency)', 'sum'),
        st=('_st', 'sum'),
        rev=('_rev', 'sum')
    ).reset_index().sort_values('st', ascending=False).head(10)

    upsell_data = _calc_upsell(client_history, extract_report_month(csv_filename),
                               float(spnd), float(conv))

    html_str = generate_html(
        csv_filename, client_name, conv_label, has_revenue,
        (imp, clk, spnd, conv, st, pcv, rev, uh),
        grp_chan, grp_cre, grp_site, prev_data, upsell_data, client_history
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
