#!/usr/bin/env python3
"""
Match Analysis Report Generator
Run with: python3 report_generator_updated.py
"""

import csv
import os
import sys
import base64
from pathlib import Path
from datetime import datetime

# ── HELPERS ───────────────────────────────────────────────────────────────────

def is_legal(r):
    return r.get('Legal Ball', '').strip() == 'Yes'

def is_Boundary(r):
    try:
        return int(r.get('Runs', 0) or 0) in [4, 6]
    except:
        return False

def is_Dot(r):
    try:
        return is_legal(r) and int(r.get('Runs', 0) or 0) == 0 and int(r.get('Extra Runs', 0) or 0) == 0
    except:
        return False

def tot_runs(r):
    try:
        return int(r.get('Runs', 0) or 0) + int(r.get('Extra Runs', 0) or 0)
    except:
        return 0

def o_fmt(balls):
    return f"{balls // 6}.{balls % 6}"

def calc_block(rows):
    L = sum(1 for r in rows if is_legal(r))
    runs = sum(tot_runs(r) for r in rows)
    wkts = sum(1 for r in rows if r.get('Wicket', '').strip())
    Dots = sum(1 for r in rows if is_Dot(r))
    bounds = sum(1 for r in rows if is_Boundary(r))
    overs = L / 6 if L > 0 else 0
    return {
        'runs': runs, 'wkts': wkts, 'L': L, 'Dots': Dots, 'bounds': bounds,
        'rr':      f"{runs/overs:.1f}" if overs > 0 else '0.0',
        'Dot_pct': f"{Dots/L*100:.1f}" if L > 0 else '0.0',
        'bnd_pct': f"{bounds/L*100:.1f}" if L > 0 else '0.0',
    }

def get_phase(rows, f, t):
    result = []
    for r in rows:
        try:
            o = int(r.get('Over', 0) or 0)
            if f <= o <= t:
                result.append(r)
        except:
            pass
    return result

TYPE_LABELS = {
    'RF':'RF - Right-arm fast','RFM':'RFM - Right-arm fast-medium',
    'RMF':'RMF - Right-arm medium-fast','RM':'RM - Right-arm medium',
    'LOB':'LOB - Left-arm orthodox spin','ROB':'ROB - Right-arm off-break',
    'LB':'LB - Leg-break','RLB':'RLB - Right-arm leg-break',
    'LFM':'LFM - Left-arm fast-medium','LF':'LF - Left-arm fast',
    'SLA':'SLA - Slow left-arm orthodox',
}

def rr_col(v):
    try:
        n=float(v); return 'good' if n<6 else 'warn' if n<9 else 'bad'
    except: return ''

def Dot_col(v):
    try:
        n=float(v); return 'good' if n>=40 else 'warn' if n>=25 else 'bad'
    except: return ''

def bnd_col(v):
    try:
        n=float(v); return 'bad' if n>=20 else 'warn' if n>=10 else 'good'
    except: return ''

# Batting-specific colour logic (inverted from bowling)
def bat_rr_col(v):
    try:
        n=float(v); return 'good' if n>=9 else 'warn' if n>=6 else 'bad'
    except: return ''

def bat_dot_col(v):
    try:
        n=float(v); return 'bad' if n>=40 else 'warn' if n>=25 else 'good'
    except: return ''

def bat_bnd_col(v):
    try:
        n=float(v); return 'good' if n>=20 else 'warn' if n>=10 else 'bad'
    except: return ''

def phase_row_html(name, label, cls, b, batting=False):
    _rr  = bat_rr_col  if batting else rr_col
    _bnd = bat_bnd_col if batting else bnd_col
    _dot = bat_dot_col if batting else Dot_col
    return f"""<tr>
      <td class="{cls}"><div class="phase-name">{name}</div><div class="phase-overs">Overs {label}</div></td>
      <td class="r score">{b['runs']}</td><td class="r">{b['wkts']}</td>
      <td class="r {_rr(b['rr'])}">{b['rr']}</td>
      <td class="r {_bnd(b['bnd_pct'])}">{b['bnd_pct']}%</td>
      <td class="r {_dot(b['Dot_pct'])}">{b['Dot_pct']}%</td>
    </tr>"""

def bullet(html):
    return f'<div class="bullet"><div class="bDot"></div><div>{html}</div></div>'

# ── BATTING SECTION ───────────────────────────────────────────────────────────

def build_batting(bat_rows, phases, inn_label=''):
    ov = calc_block(bat_rows)
    try: max_over = max(int(r.get('Over',0) or 0) for r in bat_rows)
    except: max_over = 0

    title = f"Match Summary - Batting{' (' + inn_label + ')' if inn_label else ''}"

    sum_html = f"""<div class="sum-row">
      <div class="sum-card bat-card"><div class="sum-label">Total Runs</div><div class="sum-val">{ov['runs']}</div><div class="sum-sub">{ov['wkts']} Wickets &middot; {max_over} Overs</div></div>
      <div class="sum-card accent-card"><div class="sum-label">Run Rate</div><div class="sum-val">{ov['rr']}</div><div class="sum-sub">Per Over</div></div>
      <div class="sum-card bat-card"><div class="sum-label">Dot Ball %</div><div class="sum-val">{ov['Dot_pct']}%</div><div class="sum-sub">{ov['Dots']} Dots / {ov['L']} Legal Balls</div></div>
      <div class="sum-card bat-card"><div class="sum-label">Boundary %</div><div class="sum-val">{ov['bnd_pct']}%</div><div class="sum-sub">{ov['bounds']} Boundaries / {ov['L']} Legal Balls</div></div>
    </div>"""

    # Phase table (only if phases provided)
    phase_section = ''
    if phases:
        phase_html = ''
        for name,label,f,t,cls in phases:
            b = calc_block(get_phase(bat_rows,f,t))
            if b['L'] > 0: phase_html += phase_row_html(name,label,cls,b,batting=True)
        if phase_html:
            phase_section = f"""<div class="bat-sec-hdr">Phase Analysis</div>
      <div class="tbl-wrap"><table>
        <thead><tr class="bat-head"><th style="width:150px">Phase</th><th class="r">Runs</th><th class="r">Wickets</th><th class="r">Run Rate</th><th class="r">Boundary %</th><th class="r">Dot Ball %</th></tr></thead>
        <tbody>{phase_html}</tbody></table></div>"""

    # Bowling type table
    types = sorted(set(r.get('Bowler Type','').strip() for r in bat_rows if r.get('Bowler Type','').strip()))
    type_html = ''
    type_blocks = []
    for bt in types:
        b = calc_block([r for r in bat_rows if r.get('Bowler Type','').strip()==bt])
        if b['L'] == 0: continue
        type_blocks.append((bt,b))
        lbl = TYPE_LABELS.get(bt,bt)
        type_html += f"""<tr>
          <td class="name">{lbl}</td><td class="r">{o_fmt(b['L'])}</td>
          <td class="r">{b['wkts']}</td><td class="r score">{b['runs']}</td>
          <td class="r {bat_rr_col(b['rr'])}">{b['rr']}</td>
          <td class="r {bat_bnd_col(b['bnd_pct'])}">{b['bnd_pct']}%</td>
          <td class="r {bat_dot_col(b['Dot_pct'])}">{b['Dot_pct']}%</td>
        </tr>"""

    # Individual Batting
    batters,order,dismissals = {},{},{}
    for r in bat_rows:
        b = r.get('Batter','').strip()
        if not b: continue
        if b not in batters:
            batters[b]={'runs':0,'balls':0,'fours':0,'sixes':0,'Dots':0,'off_runs':0,'leg_runs':0,'hand':''}
            order[b]=len(order)
        if not batters[b]['hand'] and r.get('Batting Hand','').strip():
            batters[b]['hand'] = r.get('Batting Hand','').strip()
        if is_legal(r):
            batters[b]['balls']+=1
            if is_Dot(r): batters[b]['Dots']+=1
        try:
            runs=int(r.get('Runs',0) or 0)
            batters[b]['runs']+=runs
            if runs==4: batters[b]['fours']+=1
            if runs==6: batters[b]['sixes']+=1
            if runs > 0:
                fx = r.get('FieldX','').strip()
                if fx:
                    hand = batters[b]['hand'] or r.get('Batting Hand','').strip()
                    is_leg = (int(fx) >= 175) if hand != 'LHB' else (int(fx) < 175)
                    if is_leg:
                        batters[b]['leg_runs'] += runs
                    else:
                        batters[b]['off_runs'] += runs
        except: pass
        dis=r.get('Dismissed Batter','').strip()
        wkt=r.get('Wicket','').strip()
        if wkt and dis:
            wkt_clean = next(iter(dict.fromkeys(p.strip() for p in wkt.split(',') if p.strip())), wkt)
            dismissals[dis]={'how':wkt_clean,'bowler':r.get('Bowler','').strip()}

    # Find top scorer for row highlight
    top_scorer = max(order.keys(), key=lambda n: batters[n]['runs']) if order else None
    ind_html=''
    for name in sorted(order, key=lambda n: order[n]):
        s=batters[name]
        bounds=s['fours']+s['sixes']
        sr=f"{s['runs']/s['balls']*100:.0f}" if s['balls']>0 else '0'
        dp=f"{s['Dots']/s['balls']*100:.1f}" if s['balls']>0 else '0.0'
        bp=f"{bounds/s['balls']*100:.1f}" if s['balls']>0 else '0.0'
        dis=dismissals.get(name)
        dis_str=f'<span class="pill pill-r">{dis["how"]}</span> <span class="dis-bowler">b {dis["bowler"]}</span>' if dis else '<span class="pill pill-g">Not Out</span>'
        side_total = s['off_runs'] + s['leg_runs']
        off_pct = f"{s['off_runs']/side_total*100:.0f}%" if side_total > 0 else '-'
        leg_pct = f"{s['leg_runs']/side_total*100:.0f}%" if side_total > 0 else '-'
        row_cls = ' class="top-row"' if name == top_scorer and s['runs'] > 0 else ''
        ind_html+=f"""<tr{row_cls}>
          <td class="name">{name}</td>
          <td class="r score">{s['runs']} <span style="font-size:0.75em;opacity:0.7">({s['balls']}b)</span></td>
          <td class="r">{sr}</td><td class="r">{s['fours']}/{s['sixes']}</td>
          <td class="r {bat_bnd_col(bp)}">{bp}%</td>
          <td class="r {bat_dot_col(dp)}">{dp}%</td>
          <td class="r">{off_pct}</td><td class="r">{leg_pct}</td>
          <td>{dis_str}</td>
        </tr>"""

    # Dismissal tally
    from collections import Counter
    dis_counts = Counter(v['how'] for v in dismissals.values() if v.get('how'))
    if dis_counts:
        tally_items = ''.join(
            f'<div class="dis-tally-item"><span class="dis-tally-count">{cnt}</span><span class="dis-tally-label">{how}</span></div>'
            for how, cnt in sorted(dis_counts.items(), key=lambda x: -x[1])
        )
        dis_tally_html = f'<div class="dis-tally">{tally_items}</div>'
    else:
        dis_tally_html = ''

    # Partnerships
    pships = {}   # pnum -> dict
    pship_order = []
    for r in bat_rows:
        pnum = r.get('Partnership Number','').strip()
        if not pnum: continue
        if pnum not in pships:
            pships[pnum] = {'runs':0,'balls':0,'batter_runs':{},'wicket':'','bowler':'','first_ball':None}
            pship_order.append(pnum)
        p = pships[pnum]
        try:
            ib = int(r.get('Innings Ball',0) or 0)
            if p['first_ball'] is None or ib < p['first_ball']:
                p['first_ball'] = ib
        except: pass
        p['runs'] += tot_runs(r)
        if is_legal(r):
            p['balls'] += 1
        bname = r.get('Batter','').strip()
        if bname:
            try: p['batter_runs'][bname] = p['batter_runs'].get(bname, 0) + int(r.get('Runs',0) or 0)
            except: pass
        wkt = r.get('Wicket','').strip()
        if wkt:
            wkt_clean = next(iter(dict.fromkeys(x.strip() for x in wkt.split(',') if x.strip())), wkt)
            p['wicket'] = wkt_clean
            p['bowler'] = r.get('Bowler','').strip()

    pship_rows_html = ''
    for i, pnum in enumerate(sorted(pship_order, key=lambda x: pships[x]['first_ball'] or 0), 1):
        p = pships[pnum]
        batters_sorted = sorted(p['batter_runs'].items(), key=lambda x: -x[1])
        if len(batters_sorted) >= 2:
            pair = f"{batters_sorted[0][0]} ({batters_sorted[0][1]}) / {batters_sorted[1][0]} ({batters_sorted[1][1]})"
        elif len(batters_sorted) == 1:
            pair = f"{batters_sorted[0][0]} ({batters_sorted[0][1]})"
        else:
            pair = '-'
        rr = f"{p['runs']/(p['balls']/6):.1f}" if p['balls'] > 0 else '-'
        if p['wicket']:
            end_str = f'<span class="pill pill-r">{p["wicket"]}</span>'
            if p['bowler']: end_str += f' <span class="dis-bowler">b {p["bowler"]}</span>'
        else:
            end_str = '<span class="pill pill-g">Not Out</span>'
        top_p = max(pships.values(), key=lambda x: x['runs']) if pships else None
        row_cls = ' class="top-row"' if p is top_p and p['runs'] > 0 else ''
        pship_rows_html += f"""<tr{row_cls}>
          <td class="r" style="width:30px">{i}</td>
          <td class="name">{pair}</td>
          <td class="r score">{p['runs']}</td>
          <td class="r">{p['balls']}</td>
          <td class="r">{rr}</td>
          <td>{end_str}</td>
        </tr>"""
    partnerships_html = f"""<div class="tbl-wrap"><table>
      <thead><tr class="bat-head"><th class="r" style="width:30px">#</th><th style="width:280px">Partnership</th><th class="r">Runs</th><th class="r">Balls</th><th class="r">Run Rate</th><th>How Out</th></tr></thead>
      <tbody>{pship_rows_html}</tbody></table></div>""" if pship_rows_html else ''

    # Key Observations
    bullets_list=[]
    if phases:
        for name,label,f,t,cls in phases:
            b=calc_block(get_phase(bat_rows,f,t))
            if b['L']>0: bullets_list.append(f"<b>{name} ({label}):</b> {b['runs']}/{b['wkts']} - Run Rate {b['rr']}, Boundary {b['bnd_pct']}%, Dot Ball {b['Dot_pct']}%")
    qual=[(bt,b) for bt,b in type_blocks if b['L']>=6]
    if len(qual)>1:
        lo=min(qual,key=lambda x:float(x[1]['rr']))
        hi=max(qual,key=lambda x:float(x[1]['rr']))
        if lo[0]!=hi[0]:
            bullets_list.append(f"<b>Lower Run Rate vs {TYPE_LABELS.get(lo[0],lo[0])}:</b> {lo[1]['rr']} - Boundary {lo[1]['bnd_pct']}%, Dot {lo[1]['Dot_pct']}%")
            bullets_list.append(f"<b>Higher Run Rate vs {TYPE_LABELS.get(hi[0],hi[0])}:</b> {hi[1]['rr']} - Boundary {hi[1]['bnd_pct']}%, Dot {hi[1]['Dot_pct']}%")
    if order:
        top=max(order.keys(),key=lambda n:batters[n]['runs'])
        s=batters[top]; bounds=s['fours']+s['sixes']
        bp=f"{bounds/s['balls']*100:.1f}" if s['balls']>0 else '0.0'
        dp=f"{s['Dots']/s['balls']*100:.1f}" if s['balls']>0 else '0.0'
        sr_top=f"{s['runs']/s['balls']*100:.0f}" if s['balls']>0 else '0'
        bullets_list.append(('top', f"<b>Top Batter - {top}:</b> {s['runs']} ({s['balls']}b) SR {sr_top}, Boundary {bp}%, Dot {dp}%"))
    bullets_html=''.join(
        f'<div class="top-perf"><div class="bDot"></div><div>{b[1]}</div></div>' if isinstance(b, tuple) and b[0]=='top'
        else bullet(b)
        for b in bullets_list
    )

    return f"""<div class="innings-block">
      <div class="section-banner bat-banner">
        <div><div class="section-banner-title">{title}</div><div class="section-banner-sub">Batting Performance Analysis</div></div>
      </div>
      <div class="section-body bat-body">
        <div class="bat-sec-hdr">Match Summary</div>{sum_html}
        {phase_section}
        <div class="bat-sec-hdr">Batting vs Bowling Type</div>
        <div class="tbl-wrap"><table>
          <thead><tr class="bat-head"><th style="width:220px">Bowling Type</th><th class="r">Overs</th><th class="r">Wickets</th><th class="r">Runs</th><th class="r">Run Rate</th><th class="r">Boundary %</th><th class="r">Dot Ball %</th></tr></thead>
          <tbody>{type_html}</tbody></table></div>
        <div class="bat-sec-hdr">Individual Batting</div>
        <div class="tbl-wrap"><table>
          <thead><tr class="bat-head"><th style="width:160px">Batter</th><th class="r">Runs (b)</th><th class="r">SR</th><th class="r">4s/6s</th><th class="r">Boundary %</th><th class="r">Dot Ball %</th><th class="r">Off Side %</th><th class="r">Leg Side %</th><th>Dismissal</th></tr></thead>
          <tbody>{ind_html}</tbody></table></div>
        {dis_tally_html}
        <div class="bat-sec-hdr">Partnerships</div>
        {partnerships_html}
        <div class="bat-sec-hdr">Key Observations</div>
        <div class="bullets">{bullets_html}</div>
      </div>
    </div>"""

# ── BOWLING SECTION ───────────────────────────────────────────────────────────

def build_bowling(bowl_rows, phases, inn_label='', fmt=''):
    ov=calc_block(bowl_rows)
    try: max_over=max(int(r.get('Over',0) or 0) for r in bowl_rows)
    except: max_over=0

    title = f"Match Summary - Bowling{' (' + inn_label + ')' if inn_label else ''}"

    sum_html=f"""<div class="sum-row">
      <div class="sum-card bowl-card"><div class="sum-label">Runs Conceded</div><div class="sum-val">{ov['runs']}</div><div class="sum-sub">{ov['wkts']} Wickets &middot; {max_over} Overs</div></div>
      <div class="sum-card accent-card"><div class="sum-label">Run Rate</div><div class="sum-val">{ov['rr']}</div><div class="sum-sub">Per Over Conceded</div></div>
      <div class="sum-card bowl-card"><div class="sum-label">Dot Ball %</div><div class="sum-val">{ov['Dot_pct']}%</div><div class="sum-sub">{ov['Dots']} Dots / {ov['L']} Legal Balls</div></div>
      <div class="sum-card bowl-card"><div class="sum-label">Boundary %</div><div class="sum-val">{ov['bnd_pct']}%</div><div class="sum-sub">{ov['bounds']} Boundaries / {ov['L']} Legal Balls</div></div>
    </div>"""

    # Phase table (only if phases provided)
    phase_section = ''
    if phases:
        phase_html=''
        for name,label,f,t,cls in phases:
            b=calc_block(get_phase(bowl_rows,f,t))
            if b['L']>0: phase_html+=phase_row_html(name,label,cls,b)
        if phase_html:
            phase_section = f"""<div class="bowl-sec-hdr">Phase Analysis - When Bowling</div>
      <div class="tbl-wrap"><table>
        <thead><tr class="bowl-head"><th style="width:150px">Phase</th><th class="r">Runs Conceded</th><th class="r">Wickets</th><th class="r">Run Rate</th><th class="r">Boundary %</th><th class="r">Dot Ball %</th></tr></thead>
        <tbody>{phase_html}</tbody></table></div>"""

    bowlers,bowl_order={},[]
    for r in bowl_rows:
        b=r.get('Bowler','').strip()
        if not b: continue
        if b not in bowlers:
            bowlers[b]={'runs':0,'balls':0,'wkts':0,'Dots':0,'bounds':0,'wides':0}
            bowl_order.append(b)
        bowlers[b]['runs']+=tot_runs(r)
        if is_legal(r):
            bowlers[b]['balls']+=1
            if is_Dot(r): bowlers[b]['Dots']+=1
            if is_Boundary(r): bowlers[b]['bounds']+=1
        wkt=r.get('Wicket','').strip()
        if wkt and wkt.lower()!='run out': bowlers[b]['wkts']+=1
        if 'wide' in r.get('Extra','').lower(): bowlers[b]['wides']+=1

    # Find best bowler for highlight (most wickets; tiebreak by economy)
    qual_names = [n for n in bowl_order if bowlers[n]['balls'] >= 6]
    top_bowler = None
    if qual_names:
        top_bowler = max(qual_names, key=lambda n: (bowlers[n]['wkts'], -bowlers[n]['runs']/bowlers[n]['balls']))
    ind_html=''
    for name in bowl_order:
        s=bowlers[name]
        eco=f"{s['runs']/(s['balls']/6):.1f}" if s['balls']>0 else '-'
        dp=f"{s['Dots']/s['balls']*100:.1f}" if s['balls']>0 else '0.0'
        bp=f"{s['bounds']/s['balls']*100:.1f}" if s['balls']>0 else '0.0'
        wkt_html=f'<span class="pill pill-g">{s["wkts"]}</span>' if s['wkts']>0 else '0'
        row_cls = ' class="top-row"' if name == top_bowler else ''
        ind_html+=f"""<tr{row_cls}>
          <td class="name">{name}</td><td class="r">{o_fmt(s['balls'])}</td>
          <td class="r">{wkt_html}</td><td class="r score">{s['runs']}</td>
          <td class="r {rr_col(eco)}">{eco}</td>
          <td class="r {bnd_col(bp)}">{bp}%</td>
          <td class="r {Dot_col(dp)}">{dp}%</td>
          <td class="r">{s['wides']}</td>
        </tr>"""

    # ── LENGTH & STUMP ANALYSIS ───────────────────────────────────────────────
    LENGTH_ORDER = ['Full', 'Yorker', 'Length', 'BOL', 'Short', 'Full Toss', 'Bouncer']
    LENGTH_NORM  = {'Back Of Length': 'BOL'}
    legal_rows = [r for r in bowl_rows if is_legal(r)]
    total_legal = len(legal_rows)

    from collections import Counter
    team_len_counts = Counter(LENGTH_NORM.get(r.get('Length','').strip(), r.get('Length','').strip()) for r in legal_rows if r.get('Length','').strip())
    stump_total = sum(1 for r in legal_rows if r.get('Line','').strip() == 'Line')
    stump_pct = f"{stump_total/total_legal*100:.1f}" if total_legal > 0 else '0.0'

    # Team-level summary cards
    len_sum_html = f"""<div class="sum-row">
      <div class="sum-card bowl-card"><div class="sum-label">Stump-Line Balls</div><div class="sum-val">{stump_total}</div><div class="sum-sub">{stump_pct}% of Legal Balls</div></div>
    </div>"""

    # Per-bowler length/stump table
    bowl_len = {}
    for r in bowl_rows:
        b = r.get('Bowler','').strip()
        if not b or not is_legal(r): continue
        if b not in bowl_len:
            bowl_len[b] = {'total': 0, 'stump': 0, 'lengths': Counter()}
        bowl_len[b]['total'] += 1
        if r.get('Line','').strip() == 'Line':
            bowl_len[b]['stump'] += 1
        lng = r.get('Length','').strip()
        if lng:
            bowl_len[b]['lengths'][LENGTH_NORM.get(lng, lng)] += 1

    used_lengths = [l for l in LENGTH_ORDER if any(l in bd['lengths'] for bd in bowl_len.values())]
    len_head = ''.join(f'<th class="r">{ln}</th>' for ln in used_lengths)
    len_rows_bowler = ''
    for bname in bowl_order:
        if bname not in bowl_len: continue
        bd = bowl_len[bname]
        t = bd['total']
        if t == 0: continue
        sp = f"{bd['stump']/t*100:.0f}%" if t > 0 else '-'
        cells = ''.join(f'<td class="r">{bd["lengths"].get(ln,0)/t*100:.0f}%</td>' for ln in used_lengths)
        len_rows_bowler += f'<tr><td class="name">{bname}</td>{cells}<td class="r stump-hi">{sp}</td></tr>'
    len_bowler_table = f"""<div class="tbl-wrap"><table>
      <thead><tr class="bowl-head"><th style="width:160px">Bowler</th>{len_head}<th class="r">Stump %</th></tr></thead>
      <tbody>{len_rows_bowler}</tbody></table></div>""" if len_rows_bowler else ''

    # ── CONSECUTIVE DOT BALLS (red ball only) ────────────────────────────────
    consec_html = ''
    if fmt == 'red_ball':
        try:
            sorted_all = sorted(bowl_rows, key=lambda r: int(r.get('Innings Ball', 0) or 0))
        except:
            sorted_all = bowl_rows
        sequences, current = [], []
        for r in sorted_all:
            if tot_runs(r) == 0:
                current.append(r)
            else:
                if len(current) >= 18:
                    sequences.append(current[:])
                current = []
        if len(current) >= 18:
            sequences.append(current[:])
        if sequences:
            seq_items = ''
            for seq in sorted(sequences, key=len, reverse=True):
                s_ov = seq[0].get('Over',''); s_bl = seq[0].get('Ball','')
                e_ov = seq[-1].get('Over',''); e_bl = seq[-1].get('Ball','')
                bowlers_in = list(dict.fromkeys(r.get('Bowler','').strip() for r in seq if r.get('Bowler','').strip()))
                seq_items += f'<div class="consec-item"><span class="consec-count">{len(seq)}</span><span class="consec-detail">consecutive balls with no runs &mdash; Overs {s_ov}.{s_bl} to {e_ov}.{e_bl} &mdash; {", ".join(bowlers_in)}</span></div>'
            consec_html = f'<div class="bowl-sec-hdr">18+ Consecutive Balls Without Runs</div><div class="consec-wrap">{seq_items}</div>'

    bullets_list=[]
    if phases:
        for name,label,f,t,cls in phases:
            b=calc_block(get_phase(bowl_rows,f,t))
            if b['L']>0: bullets_list.append(f"<b>{name} ({label}):</b> {b['runs']} Conceded, {b['wkts']} Wickets - Run Rate {b['rr']}, Boundary {b['bnd_pct']}%, Dot {b['Dot_pct']}%")
    qual=[(n,bowlers[n]) for n in bowl_order if bowlers[n]['balls']>=6]
    if qual:
        best=min(qual,key=lambda x:x[1]['runs']/x[1]['balls'])
        worst=max(qual,key=lambda x:x[1]['runs']/x[1]['balls'])
        bs,ws=best[1],worst[1]
        bullets_list.append(('top', f"<b>Best Bowler - {best[0]}:</b> {o_fmt(bs['balls'])}, {bs['wkts']}w/{bs['runs']}r - Run Rate {bs['runs']/(bs['balls']/6):.1f}, Dot {bs['Dots']/bs['balls']*100:.1f}%"))
        if best[0]!=worst[0]:
            bullets_list.append(f"<b>Most Expensive - {worst[0]}:</b> {o_fmt(ws['balls'])}, {ws['wkts']}w/{ws['runs']}r - Run Rate {ws['runs']/(ws['balls']/6):.1f}, Boundary {ws['bounds']/ws['balls']*100:.1f}%")
        dl=max(qual,key=lambda x:x[1]['Dots']/x[1]['balls'])
        d=dl[1]
        bullets_list.append(f"<b>Dot Ball Leader - {dl[0]}:</b> {d['Dots']/d['balls']*100:.1f}% ({d['Dots']}/{d['balls']} Legal Balls)")
    bullets_html=''.join(
        f'<div class="top-perf"><div class="bDot"></div><div>{b[1]}</div></div>' if isinstance(b, tuple) and b[0]=='top'
        else bullet(b)
        for b in bullets_list
    )

    return f"""<div class="innings-block">
      <div class="section-banner bowl-banner">
        <div><div class="section-banner-title">{title}</div><div class="section-banner-sub">Bowling Performance Analysis</div></div>
      </div>
      <div class="section-body bowl-body">
        <div class="bowl-sec-hdr">Match Summary</div>{sum_html}
        {phase_section}
        <div class="bowl-sec-hdr">Individual Bowling</div>
        <div class="tbl-wrap"><table>
          <thead><tr class="bowl-head"><th style="width:160px">Bowler</th><th class="r">Overs</th><th class="r">Wickets</th><th class="r">Runs</th><th class="r">Run Rate</th><th class="r">Boundary %</th><th class="r">Dot Ball %</th><th class="r">Wides</th></tr></thead>
          <tbody>{ind_html}</tbody></table></div>
        <div class="bowl-sec-hdr">Length &amp; Stump-Line Analysis</div>
        {len_sum_html}
        <div class="bowl-sec-hdr">Length &amp; Stump Breakdown by Bowler</div>
        {len_bowler_table}
        {consec_html}
        <div class="bowl-sec-hdr">Key Observations</div>
        <div class="bullets bowl-bullets">{bullets_html}</div>
      </div>
    </div>"""

# ── DETECT FORMAT & BUILD HTML ────────────────────────────────────────────────

def detect_format(rows, team):
    """Detect T20, 50-over, or red ball based on innings structure."""
    innings_nums = sorted(set(int(r.get('Innings', 1)) for r in rows))
    num_innings = len(innings_nums)

    # Red ball: 3 or 4 innings total in match, or team bats in innings 1 AND 3
    team_bat_innings = sorted(set(
        int(r.get('Innings', 1)) for r in rows
        if r.get('Batting Team', '').strip() == team
    ))

    if len(team_bat_innings) >= 2:
        return 'red_ball', team_bat_innings

    # White ball: check max over to distinguish T20 vs 50-over
    bat_rows = [r for r in rows if r.get('Batting Team', '').strip() == team]
    try:
        max_over = max(int(r.get('Over', 0) or 0) for r in bat_rows)
    except:
        max_over = 20

    if max_over <= 20:
        return 't20', team_bat_innings
    else:
        return 'fifty_over', team_bat_innings

def build_html(rows, team):
    s = rows[0]
    competition = s.get('Competition', '').strip()
    match       = s.get('Match', '').strip()
    date        = s.get('Date', '').strip()
    venue       = s.get('Venue', '').strip()
    result      = s.get('Result', '').strip()
    won = team.lower().split()[-1] in result.lower()
    result_html = f'<span class="result-tag {"r-win" if won else "r-loss"}">{result}</span>' if result else ''
    meta = ' &middot; '.join(x for x in [match, date, venue] if x)

    fmt, team_bat_innings = detect_format(rows, team)

    # Scorecard table
    all_teams = sorted(set(r.get('Batting Team','').strip() for r in rows if r.get('Batting Team','').strip()))
    sc_rows = []
    for bat_team in all_teams:
        innings_nums = sorted(set(int(r.get('Innings',1)) for r in rows if r.get('Batting Team','').strip()==bat_team))
        for inn in innings_nums:
            inn_rows = [r for r in rows if r.get('Batting Team','').strip()==bat_team and int(r.get('Innings',1))==inn]
            runs = sum(tot_runs(r) for r in inn_rows)
            wkts = sum(1 for r in inn_rows if r.get('Wicket','').strip())
            legal = sum(1 for r in inn_rows if is_legal(r))
            balls_rem = legal % 6
            ov_str = f"{legal//6}.{balls_rem}" if balls_rem else f"{legal//6}"
            sc_rows.append(f"<tr><td class='sc-team'>{bat_team}</td><td class='sc-score'>{runs}/{wkts}</td><td class='sc-ov'>({ov_str} ov)</td></tr>")
    scorecard_html = f'<table class="sc-table">{"".join(sc_rows)}</table>' if sc_rows else ''

    # Generated timestamp
    generated = datetime.now().strftime('%-d %B %Y, %H:%M')

    # Logo — loaded from file at runtime (not embedded; keep badge file in .gitignore)
    logo_html = ''
    logo_path = Path(os.path.dirname(os.path.abspath(__file__))) / 'lccc_badge.png'  # add your badge file here
    if logo_path.exists():
        raw = logo_path.read_bytes()
        b64 = base64.b64encode(raw).decode('ascii')
        mime = 'image/webp' if raw[:4] == b'RIFF' else 'image/png'
        logo_html = f'<img src="data:{mime};base64,{b64}" class="hdr-logo" alt="Club Badge">'

    # Build phases based on format
    if fmt == 't20':
        phases = [('Powerplay','1-6',1,6,'pp-bar'),('Middle','7-16',7,16,'mid-bar'),('Death','17-20',17,20,'death-bar')]
        format_label = 'T20'
    elif fmt == 'fifty_over':
        phases = [('Powerplay','1-10',1,10,'pp-bar'),('Middle','11-40',11,40,'mid-bar'),('Death','41-50',41,50,'death-bar')]
        format_label = '50-over'
    else:
        phases = []  # No phases for red ball
        format_label = 'Red Ball'

    # Build content blocks
    content_html = ''

    if fmt == 'red_ball':
        # Red ball: all batting innings first, then all bowling innings
        all_innings = sorted(set(int(r.get('Innings', 1)) for r in rows))
        team_bowl_innings = [i for i in all_innings if i not in team_bat_innings]

        # All batting innings
        for idx_i, bat_inn in enumerate(team_bat_innings, 1):
            bat_rows = [r for r in rows if r.get('Batting Team','').strip()==team and int(r.get('Innings',1))==bat_inn]
            inn_label = '1st Innings' if idx_i == 1 else '2nd Innings'
            if idx_i > 1:
                content_html += '<div class="inn-sep"><div class="inn-sep-line"></div><div class="inn-sep-label">2ND BATTING INNINGS</div><div class="inn-sep-line"></div></div>'
            if bat_rows:
                content_html += build_batting(bat_rows, phases, inn_label)

        # Separator between batting and bowling
        content_html += '<div class="section-sep"><div class="section-sep-line"></div><div class="section-sep-label">BOWLING</div><div class="section-sep-line"></div></div>'

        # All bowling innings
        for idx_i, bowl_inn in enumerate(team_bowl_innings, 1):
            bowl_rows = [r for r in rows if r.get('Bowling Team','').strip()==team and int(r.get('Innings',1))==bowl_inn]
            inn_label = '1st Innings' if idx_i == 1 else '2nd Innings'
            if idx_i > 1:
                content_html += '<div class="inn-sep"><div class="inn-sep-line"></div><div class="inn-sep-label">2ND BOWLING INNINGS</div><div class="inn-sep-line"></div></div>'
            if bowl_rows:
                content_html += build_bowling(bowl_rows, phases, inn_label, fmt)

    else:
        # White ball: single innings each
        bat_rows  = [r for r in rows if r.get('Batting Team','').strip()==team]
        bowl_rows = [r for r in rows if r.get('Bowling Team','').strip()==team]
        content_html += build_batting(bat_rows, phases)
        content_html += build_bowling(bowl_rows, phases, fmt=fmt)

    css = """
:root{
  --green-dark:#00703C;
  --green-mid:#005C32;
  --green-light:#00a651;
  --bat-accent:#1a1a1a;
  --bowl-accent:#c8102e;
  --bat-bg:#fdf9ee;
  --bowl-bg:#fff5f6;
  --bat-border:#c9a227;
  --bowl-border:#c8102e;
  --gold:#c9a227;
  --heading:#1a1a1a;
  --red:#c8102e;
  --amber:#b7791f;
  --bg:#f0f2f5;
  --surface:#ffffff;
  --border:#e2e8f0;
  --text:#1a202c;
  --mid:#4a5568;
  --light:#718096;
  --sans:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:var(--sans);background:var(--bg);color:var(--text);font-size:15px;line-height:1.6;}

/* HEADER */
.hdr{background:linear-gradient(135deg,#00512d 0%,var(--green-dark) 60%,#00a651 100%);padding:28px 28px 24px;border-bottom:4px solid rgba(255,255,255,.15);}
.hdr-inner{max-width:980px;margin:0 auto;}
.hdr-flex{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;}
.hdr-left{flex:1;min-width:0;}
.hdr-logo{width:90px;height:90px;object-fit:contain;flex-shrink:0;border-radius:6px;}
.hdr-title{font-size:clamp(13px,1.8vw,17px);font-weight:600;color:rgba(255,255,255,.65);letter-spacing:.01em;line-height:1.2;margin-bottom:10px;}
.hdr-meta-row{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-bottom:8px;}
.hdr-meta-item{font-size:17px;font-weight:700;color:#fff;}
.hdr-meta-sep{color:rgba(255,255,255,.3);font-size:17px;}
.hdr-team-row{font-size:17px;font-weight:800;color:#fff;margin-bottom:14px;}
.hdr-team-label{color:rgba(255,255,255,.6);font-weight:500;font-size:14px;margin-right:6px;}
.hdr-format{display:inline-block;font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;background:rgba(255,255,255,.15);color:#fff;letter-spacing:.08em;text-transform:uppercase;vertical-align:middle;}
.hdr-result-row{padding-top:12px;border-top:1px solid rgba(255,255,255,.15);}
.result-tag{display:inline-block;font-size:15px;font-weight:700;padding:8px 18px;border-radius:8px;white-space:nowrap;letter-spacing:.01em;}
.r-win{background:rgba(39,103,73,.5);color:#a7f3c0;border:1px solid rgba(167,243,192,.3);}
.r-loss{background:rgba(197,48,48,.4);color:#fca5a5;border:1px solid rgba(252,165,165,.3);}

/* MAIN LAYOUT */
.main{max-width:980px;margin:0 auto;padding:24px 16px 60px;}
.innings-block{}

/* SECTION BANNERS - batting=black/gold, bowling=red */
.section-banner{
  display:flex;align-items:center;gap:14px;
  padding:14px 20px;border-radius:10px 10px 0 0;
  margin-top:28px;margin-bottom:0;
}
.section-banner.bat-banner{background:var(--bat-accent);border-bottom:3px solid var(--gold);}
.section-banner.bowl-banner{background:var(--bowl-accent);}
.section-banner-icon{font-size:22px;line-height:1;}
.section-banner-title{font-size:18px;font-weight:700;color:#fff;letter-spacing:.01em;}
.section-banner.bat-banner .section-banner-title{color:var(--gold);}
.section-banner-sub{font-size:13px;color:rgba(255,255,255,.65);margin-top:2px;}
.section-body{
  background:var(--surface);
  border:1px solid var(--border);
  border-top:none;
  border-radius:0 0 10px 10px;
  padding:24px;
  margin-bottom:12px;
}
.section-body.bat-body{border-top:3px solid var(--gold);}
.section-body.bowl-body{border-top:3px solid var(--bowl-border);}

/* SECTION HEADERS inside body */
.bat-sec-hdr{
  font-size:13px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;
  color:var(--heading);padding:10px 14px;
  background:#fdf6e0;
  border-left:4px solid var(--gold);
  border-radius:4px;
  margin:24px 0 14px;
}
.bat-sec-hdr:first-child{margin-top:0;}
.bowl-sec-hdr{
  font-size:13px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;
  color:var(--bowl-accent);padding:10px 14px;
  background:#fff0f2;
  border-left:4px solid var(--bowl-accent);
  border-radius:4px;
  margin:24px 0 14px;
}
.bowl-sec-hdr:first-child{margin-top:0;}

/* SUMMARY CARDS */
.sum-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:20px;}
.sum-card{background:var(--surface);border:2px solid var(--border);border-radius:10px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.07);}
.sum-card.bat-card{border-top:4px solid var(--bat-border);border-color:var(--bat-border);}
.sum-card.bowl-card{border-top:4px solid var(--bowl-border);border-color:var(--bowl-border);}
.sum-card.accent-card{border-top:4px solid var(--green-dark);border-color:var(--green-dark);}
.sum-label{font-size:13px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:var(--heading);margin-bottom:6px;}
.sum-val{font-size:30px;font-weight:700;color:var(--heading);letter-spacing:-.02em;line-height:1;}
.sum-sub{font-size:12px;color:var(--light);margin-top:4px;}

/* TABLES */
.tbl-wrap{background:var(--surface);border:1px solid var(--border);border-radius:8px;overflow:hidden;margin-bottom:16px;overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:14px;min-width:480px;}
thead tr.bat-head{background:var(--bat-accent);}
thead tr.bowl-head{background:var(--bowl-accent);}
th{padding:11px 14px;text-align:left;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:rgba(255,255,255,.92);white-space:nowrap;}
thead tr.bat-head th{color:var(--gold);}
th.r{text-align:right;}
tbody tr{border-bottom:1px solid var(--border);}
tbody tr:last-child{border-bottom:none;}
tbody tr:nth-child(even){background:#fafbfc;}
tbody tr:hover{background:#fdf9ee;}
td{padding:11px 14px;color:var(--mid);vertical-align:middle;white-space:nowrap;}
td.r{text-align:right;}
td.name{font-weight:600;color:var(--text);font-size:14px;}
td.score{font-weight:700;color:var(--heading);font-size:16px;}

/* TOP PERFORMER ROW */
tr.top-row{background:#fefae8!important;}
tr.top-row td.name{color:#7a5c00;font-weight:700;}
tr.top-row td.score{color:#7a5c00;}

/* PHASE BARS */
.phase-name{font-weight:700;font-size:14px;}
.phase-overs{font-size:12px;color:var(--light);}
.pp-bar{border-left:4px solid var(--green-dark);padding-left:10px;}
.mid-bar{border-left:4px solid var(--gold);padding-left:10px;}
.death-bar{border-left:4px solid var(--red);padding-left:10px;}
.pp-bar .phase-name{color:var(--green-dark);}
.mid-bar .phase-name{color:#96720a;}
.death-bar .phase-name{color:var(--red);}

/* STAT COLOURS */
.good{color:#00703C;font-weight:700;}
.warn{color:#b7791f;font-weight:700;}
.bad{color:var(--red);font-weight:700;}

/* DISMISSAL PILLS */
.pill{display:inline-block;font-size:11px;font-weight:600;padding:3px 9px;border-radius:3px;}
.pill-g{background:#f0fff4;color:#00703C;}
.pill-r{background:#fff0f2;color:var(--red);}
.dis-bowler{font-size:13px;color:var(--light);}
.stump-hi{font-weight:700;color:var(--red);}
.consec-wrap{display:flex;flex-direction:column;gap:8px;padding:4px 0 12px;}
.consec-item{display:flex;align-items:center;gap:12px;background:#1a1a1a;border:1px solid #c8102e44;border-left:4px solid var(--red);border-radius:6px;padding:10px 14px;}
.consec-count{font-size:22px;font-weight:800;color:var(--red);min-width:36px;text-align:center;}
.consec-detail{font-size:13px;color:var(--light);}
.dis-tally{display:flex;flex-wrap:wrap;gap:10px;padding:12px 16px;background:#1a1a1a;border:1px solid #333;border-top:none;border-radius:0 0 6px 6px;margin-top:-4px;}
.dis-tally-item{display:flex;align-items:center;gap:6px;background:#252525;border:1px solid #383838;border-radius:20px;padding:4px 12px;}
.dis-tally-count{font-size:15px;font-weight:700;color:var(--gold);}
.dis-tally-label{font-size:12px;color:var(--light);}

/* KEY OBSERVATIONS */
.bullets{
  background:#1a1a1a;
  border:2px solid var(--gold);
  border-left:6px solid var(--gold);
  border-radius:0 8px 8px 0;
  padding:20px 22px;
  margin-bottom:4px;
}
.bullets.bowl-bullets{
  background:#1a1a1a;
  border-color:var(--bowl-accent);
  border-left-color:var(--bowl-accent);
}
.bullet{display:flex;gap:12px;padding:8px 0;font-size:15px;color:#e5e5e5;line-height:1.65;border-bottom:1px solid rgba(255,255,255,.08);}
.bullet:last-child{border-bottom:none;padding-bottom:0;}
.bullet:first-child{padding-top:0;}
.bullet b{color:var(--gold);font-size:15px;}
.bullets.bowl-bullets .bullet b{color:#fca5a5;}
.bDot{width:7px;height:7px;border-radius:50%;background:var(--gold);flex-shrink:0;margin-top:8px;}
.bullets.bowl-bullets .bDot{background:var(--bowl-accent);}

/* TOP PERFORMER HIGHLIGHT IN BULLETS */
.top-perf{
  background:#2a2200;border:1px solid var(--gold);border-radius:6px;
  padding:10px 14px;margin:6px 0;display:flex;gap:12px;font-size:15px;color:#e5e5e5;line-height:1.65;
}
.bullets.bowl-bullets .top-perf{background:#2a0008;border-color:var(--bowl-accent);}
.top-perf .bDot{background:var(--gold);}
.top-perf b{color:var(--gold);}
.bullets.bowl-bullets .top-perf b{color:#fca5a5;}

/* INNINGS SEPARATOR */
.inn-sep{
  margin:40px 0 32px;
  display:flex;align-items:center;gap:16px;
}
.inn-sep-line{flex:1;height:2px;background:linear-gradient(90deg,transparent,var(--green-dark),transparent);}
.inn-sep-label{font-size:11px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:var(--green-dark);white-space:nowrap;}
.section-sep{
  margin:48px 0 36px;
  display:flex;align-items:center;gap:16px;
}
.section-sep-line{flex:1;height:3px;background:linear-gradient(90deg,transparent,var(--heading),transparent);}
.section-sep-label{font-size:13px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--heading);white-space:nowrap;padding:6px 16px;background:#eef2f7;border-radius:20px;}

/* SCORECARD TABLE — reset all global table/td overrides */
.sc-table{border-collapse:collapse;margin-bottom:14px;margin-top:4px;width:auto;min-width:0!important;font-size:inherit;}
.sc-table tr,.sc-table tbody tr{background:transparent!important;border-bottom:1px solid rgba(255,255,255,.1);}
.sc-table tr:last-child,.sc-table tbody tr:last-child{border-bottom:none;}
.sc-table tr:nth-child(even),.sc-table tbody tr:nth-child(even){background:transparent!important;}
.sc-table tr:hover,.sc-table tbody tr:hover{background:transparent!important;}
.sc-table td.sc-team{font-size:13px;font-weight:600;color:rgba(255,255,255,.75)!important;padding:5px 20px 5px 0;white-space:nowrap;}
.sc-table td.sc-score{font-size:20px;font-weight:800;color:#fff!important;letter-spacing:-.01em;padding:5px 16px 5px 0;}
.sc-table td.sc-ov{font-size:13px;color:rgba(255,255,255,.5)!important;padding:5px 0;}

/* FOOTER */
.page-footer{
  margin-top:48px;
  padding:14px 0;
  border-top:2px solid var(--border);
  display:flex;justify-content:space-between;align-items:center;
  flex-wrap:wrap;gap:6px;
}
.footer-left{font-size:12px;font-weight:600;color:var(--light);letter-spacing:.02em;}
.footer-right{font-size:12px;color:var(--light);}

@media(max-width:580px){.sum-row{grid-template-columns:repeat(2,1fr);}.hdr{padding:16px;}.main{padding:14px 10px 36px;}.section-body{padding:14px;}}

@media print{
  @page{
    size:A4 portrait;
    margin:10mm 12mm 18mm 12mm;
    @bottom-right{content:"Page " counter(page) " of " counter(pages);font-size:8pt;color:#718096;font-family:-apple-system,sans-serif;}
  }

  /* Scale entire page to fit A4 width — fixes table overflow in one go */
  html{zoom:72%;}

  /* Preserve all colours and backgrounds */
  *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}

  body{background:#fff!important;}
  .main{max-width:100%!important;padding:8px 0 0;}
  .hdr-inner{max-width:100%!important;}

  /* Fix tables — remove min-width and override inline column widths */
  .tbl-wrap{overflow:visible!important;}
  table{min-width:0!important;width:100%!important;table-layout:auto;}
  th[style]{width:auto!important;}
  td{white-space:normal!important;}

  /* Repeat table headers if a table spans multiple pages */
  thead{display:table-header-group;}

  /* PAGE BREAK RULES
     — Only avoid breaks on individual rows and small cards (prevents gaps)
     — Let sections and tables flow naturally across pages
     — Keep each banner glued to the top of its section body */
  tr{page-break-inside:avoid;break-inside:avoid;}
  .sum-row{page-break-inside:avoid;break-inside:avoid;}
  .sum-card{page-break-inside:avoid;break-inside:avoid;}
  .section-banner{page-break-after:avoid;break-after:avoid;}
  .section-body{page-break-before:avoid;break-before:avoid;page-break-inside:auto;break-inside:auto;}
  .innings-block{page-break-inside:auto;break-inside:auto;}
  .tbl-wrap{page-break-inside:auto;break-inside:auto;}

  .inn-sep{margin:16px 0 12px;}
  .section-sep{margin:20px 0 16px;}
  .page-footer{display:none;}
}"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Match Analysis</title>
<style>{css}</style>
</head>
<body>
<div class="hdr"><div class="hdr-inner"><div class="hdr-flex">
  <div class="hdr-left">
    <div class="hdr-title">Post-Match Analysis</div>
    <div class="hdr-meta-row">
      <span class="hdr-meta-item">{meta}</span>
      <span class="hdr-meta-sep">&middot;</span>
      <span class="hdr-format">{format_label}</span>
    </div>
    <div class="hdr-team-row"><span class="hdr-team-label">Team Analysed:</span>{team}</div>
    {scorecard_html}
    <div class="hdr-result-row">{result_html}</div>
  </div>
  {logo_html}
</div></div></div>
<div class="main">
{content_html}
<div class="page-footer">
  <span class="footer-right">Generated: {generated}</span>
</div>
</div>
</body>
</html>"""

# ── FOLDER STRUCTURE ──────────────────────────────────────────────────────────

TEAM_FOLDERS = ['Example', 'Womens', 'Academy', 'Second Team']

def get_script_dir():
    return Path(os.path.dirname(os.path.abspath(__file__)))

def select_team_folder():
    print("  Select team:")
    for i, name in enumerate(TEAM_FOLDERS):
        print(f"    [{i+1}] {name}")
    print("")
    choice = input("  Enter number: ").strip()
    try:
        return TEAM_FOLDERS[int(choice)-1]
    except:
        print("\n  ERROR: Invalid selection.")
        input("\n  Press Enter to exit.")
        sys.exit(1)

def find_csvs(folder_name):
    csv_dir = get_script_dir() / 'CSV Files' / folder_name
    if not csv_dir.exists():
        print(f"\n  ERROR: Folder not found: {csv_dir}")
        print(f"  Make sure 'CSV Files/{folder_name}' exists in your reports folder.")
        input("\n  Press Enter to exit.")
        sys.exit(1)
    return sorted(csv_dir.glob('*.csv'))

def get_reports_dir(folder_name):
    reports_dir = get_script_dir() / 'Reports' / folder_name
    reports_dir.mkdir(parents=True, exist_ok=True)
    return reports_dir

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    print("")
    print("  Match Analysis Generator")
    print("  " + "-"*40)
    print("")

    folder_name = select_team_folder()
    print("")

    csv_files = find_csvs(folder_name)
    if not csv_files:
        print(f"\n  No CSV files found in CSV Files/{folder_name}/")
        print(f"  Add your match CSV to that folder and try again.")
        input("\n  Press Enter to exit.")
        sys.exit(1)

    print(f"  CSV files in {folder_name}:")
    for i, f in enumerate(csv_files):
        print(f"    [{i+1}] {f.name}")
    print("")
    choice = input("  Enter number: ").strip()
    try:
        csv_path = csv_files[int(choice)-1]
    except:
        print("\n  ERROR: Invalid selection.")
        input("\n  Press Enter to exit.")
        sys.exit(1)

    print(f"\n  Reading: {csv_path.name}")

    rows = []
    try:
        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            for row in reader:
                clean = {k.strip(): v.strip() for k, v in row.items() if k}
                if clean.get('Innings'):
                    rows.append(clean)
    except Exception as e:
        print(f"\n  ERROR reading file: {e}")
        input("\n  Press Enter to exit.")
        sys.exit(1)

    print(f"  Loaded {len(rows)} deliveries")

    # Detect format
    fmt, _ = detect_format(rows, rows[0].get('Batting Team','') if rows else '')
    fmt_display = {'t20':'T20','fifty_over':'50-over','red_ball':'Red Ball'}.get(fmt,'')
    print(f"  Format Detected: {fmt_display}")

    teams = sorted(set(r.get('Batting Team','') for r in rows if r.get('Batting Team','').strip()))
    if not teams:
        print("\n  ERROR: No 'Batting Team' column found in this CSV.")
        input("\n  Press Enter to exit.")
        sys.exit(1)

    print(f"\n  Teams in this match:")
    for i, t in enumerate(teams):
        print(f"    [{i+1}] {t}")
    print("")
    choice = input("  Which team to focus on? Enter number: ").strip()
    try:
        team = teams[int(choice)-1]
    except:
        print("\n  ERROR: Invalid selection.")
        input("\n  Press Enter to exit.")
        sys.exit(1)

    print(f"\n  Generating report for: {team}")

    html = build_html(rows, team)

    # Work out opponent and format for filename
    all_teams = sorted(set(r.get('Batting Team','').strip() for r in rows if r.get('Batting Team','').strip()))
    opponent = next((t for t in all_teams if t != team), 'Unknown')

    fmt_detected, _ = detect_format(rows, team)
    fmt_label = {'t20': 'T20', 'fifty_over': '50-Over', 'red_ball': 'Red-Ball'}.get(fmt_detected, 'Cricket')

    def safe_slug(s):
        return s.replace('/', '-').replace('\\', '-').replace(' ', '_').replace('(', '').replace(')', '')

    team_slug     = safe_slug(team)
    opponent_slug = safe_slug(opponent)
    out_name = f"{team_slug}_vs_{opponent_slug}_({fmt_label}).html"
    reports_dir = get_reports_dir(folder_name)
    out_path = reports_dir / out_name

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n  Report Saved To:")
    print(f"  Reports/{folder_name}/{out_name}")
    print(f"\n  Open in Chrome or Safari and send to coaches.")
    print("")
    input("  Press Enter to exit.")

if __name__ == '__main__':
    main()
    
