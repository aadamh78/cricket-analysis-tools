#!/usr/bin/env python3
"""
Pre-Game Cricket Analysis Pack Generator — Ball-by-Ball CSV Edition
=====================================================================
Parses ball-by-ball CSV data (Hawkeye/PlayCricket format),
computes player stats automatically, lets you pick who to include,
and generates a fully styled editable HTML analysis pack.

Requirements:
  pip install pandas

Usage:
  python generate_pack.py data.csv
  python generate_pack.py match1.csv match2.csv -o vs_yorkshire.html
"""

import sys, os, html, argparse
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required.  Run:  pip install pandas")
    sys.exit(1)

# ─── CONSTANTS ────────────────────────────────────────────────────────────────

RIGHT_SEAM = {'RF', 'RFM', 'RM'}
LEFT_SEAM  = {'LF', 'LFM', 'LM'}
SEAM_TYPES = RIGHT_SEAM | LEFT_SEAM
SPIN_TYPES = {'ROB', 'RLB', 'LOB', 'LWS'}

BOWL_TYPE_FULL = {
    'RF':  'Right arm fast',         'RFM': 'Right arm fast-medium',
    'RM':  'Right arm medium',       'LF':  'Left arm fast',
    'LFM': 'Left arm fast-medium',  'LM':  'Left arm medium',
    'ROB': 'Right arm off-break',   'RLB': 'Right arm leg-break',
    'LOB': 'Left arm orthodox',      'LWS': 'Left arm wrist spin',
}

HAND_MAP = {'RHB': 'Right hand bat', 'LHB': 'Left hand bat'}

SPIN_LABELS = {
    'ROB': 'ROB — Off-break',
    'RLB': 'RLB — Leg-break',
    'LOB': 'LOB — Left arm orthodox',
    'LWS': 'LWS — Left arm wrist spin',
}

BOWL_TYPE_OPTIONS = [
    '', 'Right arm fast', 'Right arm fast-medium', 'Right arm medium',
    'Left arm fast', 'Left arm fast-medium', 'Left arm medium',
    'Right arm off-break', 'Right arm leg-break',
    'Left arm orthodox', 'Left arm wrist spin',
]

# ─── CSV LOADING ──────────────────────────────────────────────────────────────

def load_csv(path):
    df = pd.read_csv(path, dtype=str, encoding='utf-8-sig')
    df.columns = [c.strip() for c in df.columns]
    # Ensure numeric columns are numeric
    for col in ['Runs', 'Extra Runs', 'Bowler Extra Runs', 'Innings']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    # Normalise blank strings
    str_cols = [c for c in df.columns if df[c].dtype == object]
    df[str_cols] = df[str_cols].fillna('')
    return df

# ─── STATS HELPERS ────────────────────────────────────────────────────────────

def _is_wide(df):
    return df['Extra'].str.contains('Wide', na=False)

def _batting_line(df):
    """Runs, balls, SR, boundary%, dot%, wickets for a subset of balls (wides already excluded)."""
    r     = int(df['Runs'].sum())
    balls = len(df)
    if balls == 0:
        return {'runs': 0, 'balls': 0, 'sr': 0.0, 'boundary_pct': 0.0, 'dot_pct': 0.0, 'wickets': 0}
    bdry    = int((df['Runs'] >= 4).sum())
    sr      = round(r / balls * 100, 1)
    bpct    = round(bdry / balls * 100, 1)
    dots    = int((df['Runs'] == 0).sum())
    dot_pct = round(dots / balls * 100, 1)
    if 'Wicket' in df.columns:
        wkt_col = df['Wicket'].fillna('').astype(str).str.strip()
        wickets = int((wkt_col.ne('') & ~wkt_col.str.contains('Run Out', case=False, na=False)).sum())
    else:
        wickets = 0
    return {'runs': r, 'balls': balls, 'sr': sr, 'boundary_pct': bpct, 'dot_pct': dot_pct, 'wickets': wickets}

def compute_batting_stats(df, player):
    bat = df[df['Batter'] == player].copy()
    if bat.empty:
        return None

    hand_raw  = bat['Batting Hand'].replace('', pd.NA).dropna().iloc[0] \
                if 'Batting Hand' in bat.columns else ''
    hand_full = HAND_MAP.get(hand_raw.strip(), hand_raw.strip())

    # Balls faced = not a wide
    faced  = bat[~_is_wide(bat)]

    def _line_for(sub):
        return _batting_line(sub[~_is_wide(sub)])

    seam_df  = bat[bat['Bowler Type'].isin(SEAM_TYPES)]
    spin_df  = bat[bat['Bowler Type'].isin(SPIN_TYPES)]
    rseam_df = bat[bat['Bowler Type'].isin(RIGHT_SEAM)]
    lseam_df = bat[bat['Bowler Type'].isin(LEFT_SEAM)]

    spin_sub = {code: _line_for(bat[bat['Bowler Type'] == code])
                for code in ['ROB', 'RLB', 'LOB', 'LWS']}

    return {
        'name':     player,
        'hand':     hand_full,
        'overall':  _batting_line(faced),
        'seam':   {'overall': _line_for(seam_df),
                   'right':   _line_for(rseam_df),
                   'left':    _line_for(lseam_df),
                   'df':      seam_df},
        'spin':   {'overall': _line_for(spin_df),
                   'sub':     spin_sub,
                   'df':      spin_df},
    }

def compute_bowling_stats(df, player):
    bowl = df[df['Bowler'] == player].copy()
    if bowl.empty:
        return None

    # Bowl type (most frequent)
    bt_raw = ''
    if 'Bowler Type' in bowl.columns:
        counts = bowl['Bowler Type'].replace('', pd.NA).dropna().value_counts()
        bt_raw = counts.index[0] if not counts.empty else ''
    bt_full = BOWL_TYPE_FULL.get(bt_raw, bt_raw)

    # Legal balls = 'Legal Ball' == 'Yes'
    legal_mask  = bowl['Legal Ball'] == 'Yes'
    legal_count = int(legal_mask.sum())
    ovs_int, ovs_rem = divmod(legal_count, 6)
    overs_str = f"{ovs_int}.{ovs_rem}"

    # Runs conceded
    runs_col  = bowl['Runs'].fillna(0).astype(int)
    extra_col = pd.to_numeric(bowl.get('Bowler Extra Runs', 0), errors='coerce').fillna(0).astype(int)
    runs_con  = int(runs_col.sum() + extra_col.sum())
    economy   = round(runs_con / (legal_count / 6), 2) if legal_count else 0.0

    # Wickets (exclude run-outs)
    # fillna('') must come before astype(str) — otherwise NaN becomes the string 'nan'
    wicket_col = bowl['Wicket'].fillna('').astype(str).str.strip()
    wkt_mask = wicket_col.ne('') & ~wicket_col.str.contains('Run Out', case=False, na=False)
    wickets  = int(wkt_mask.sum())
    average  = round(runs_con / wickets, 1) if wickets else 0.0
    sr_bowl  = round(legal_count / wickets, 1) if wickets else 0.0

    # Dot balls
    dot_mask = legal_mask & (runs_col == 0) & (extra_col == 0)
    dots     = int(dot_mask.sum())
    dot_pct  = round(dots / legal_count * 100, 1) if legal_count else 0.0

    # Boundaries conceded
    bdry     = int((runs_col >= 4).sum())
    bdry_pct = round(bdry / legal_count * 100, 1) if legal_count else 0.0

    # Matches
    matches  = int(bowl['Match'].nunique())

    # Best figures
    best_w, best_r = 0, 9999
    for (m, inn), grp in bowl.groupby(['Match', 'Innings']):
        wkt_col = grp['Wicket'].fillna('').astype(str).str.strip()
        ww = int((wkt_col.ne('') & ~wkt_col.str.contains('Run Out', case=False, na=False)).sum())
        rr = int(grp['Runs'].fillna(0).astype(int).sum()) + \
             int(pd.to_numeric(grp.get('Bowler Extra Runs', 0), errors='coerce').fillna(0).astype(int).sum())
        if ww > best_w or (ww == best_w and rr < best_r):
            best_w, best_r = ww, rr
    best_str = f"{best_w}/{best_r}" if best_w > 0 else '-'

    return {
        'name':         player,
        'bowl_type':    bt_full,
        'bowl_type_raw': bt_raw,
        'overs':        overs_str,
        'wickets':      wickets,
        'runs':         runs_con,
        'average':      average,
        'economy':      economy,
        'sr':           sr_bowl,
        'dot_pct':      dot_pct,
        'boundary_pct': bdry_pct,
        'best':         best_str,
        'matches':      matches,
        'df':           bowl,
    }

# ─── DISPLAY HELPERS ──────────────────────────────────────────────────────────

def esc(s):
    return html.escape(str(s or '').strip())

def fmt_rb(s):
    if not s or s['balls'] == 0:
        return '—'
    return f"{s['runs']}({s['balls']})"

def fmt_sr(s):
    if not s or s['balls'] == 0:
        return '—'
    return f"{s['sr']:.0f}"

def fmt_bp(s):
    if not s or s['balls'] == 0:
        return '—'
    return f"{s['boundary_pct']:.1f}%"

def fmt_dp(s):
    if not s or s['balls'] == 0:
        return '—'
    return f"{s.get('dot_pct', 0):.0f}%"

def fmt_wk(s):
    if not s or s['balls'] == 0:
        return '—'
    return str(s.get('wickets', 0))

def ta(placeholder=''):
    return f'<textarea placeholder="{esc(placeholder)}"></textarea>'

def drop_zone(id_, dtype, label):
    return (f'<div class="image-drop-zone" id="drop-{id_}-{dtype}" '
            f'title="Click or drag" onclick="triggerFileInput(event,\'{id_}\',\'{dtype}\')">'
            f'<div class="drop-icon">📷</div>'
            f'<div class="drop-label">{esc(label)}<br>'
            f'<span style="font-size:10px;opacity:0.6;">drag &amp; drop or click</span></div>'
            f'<input type="file" accept="image/*" id="file-{id_}-{dtype}" '
            f'onchange="handleFileInput(event,\'{id_}\',\'{dtype}\')">'
            f'<button class="remove-img" '
            f'onclick="event.stopPropagation();clearImage(\'{id_}\',\'{dtype}\')" '
            f'title="Remove">✕</button>'
            f'</div>')

def stat_cell(val):
    return f'<td class="stat-val">{esc(str(val))}</td>'

# ─── SVG GENERATORS ───────────────────────────────────────────────────────────

def _prep_coord_df(df, x_col, y_col):
    """Filter df to rows with valid numeric values for x_col and y_col."""
    if df is None or df.empty:
        return None
    if x_col not in df.columns or y_col not in df.columns:
        return None
    d = df.copy()
    d[x_col] = pd.to_numeric(d[x_col], errors='coerce')
    d[y_col] = pd.to_numeric(d[y_col], errors='coerce')
    d = d.dropna(subset=[x_col, y_col])
    return d if not d.empty else None

def _make_pitchmap_svg(df, mode):
    """
    Return an SVG string for a cricket pitch map.
    mode='dots_wkts': green=dot, red=wicket
    mode='boundaries': gold=4, red=6
    mode='wickets': red=wicket only
    df must already have numeric PitchX, PitchY (NaN stripped).
    PitchX: 0=bowler end (short/top), ~12=batter end (full/bottom).
    PitchY: negative=off side, positive=leg side.
    """
    W, H   = 160, 280
    PL, PR, PT, PB = 28, 28, 18, 18
    pw = W - PL - PR
    ph = H - PT - PB

    X_MIN, X_MAX = 0.0, 12.0
    Y_MIN, Y_MAX = -1.5, 1.5

    def tx(py):  # lateral → SVG x
        return PL + (py - Y_MIN) / (Y_MAX - Y_MIN) * pw

    def ty(px):  # low PitchX=top=yorker/batter end, high PitchX=bottom=short/bowler end
        return PT + (px - X_MIN) / (X_MAX - X_MIN) * ph

    y_bat  = ty(0.0)    # batter/yorker crease (top)
    y_bowl = ty(10.0)   # bowler crease (near bottom)
    cx_    = PL + pw / 2

    # Base pitch surface
    pitch_base = (
        f'<rect x="{PL}" y="{PT}" width="{pw}" height="{ph}" fill="#d4c884" rx="2"/>'
        f'<rect x="{PL + pw*0.25:.1f}" y="{PT}" width="{pw*0.5:.1f}" height="{ph}" fill="#cfc07a" rx="1"/>'
    )

    # Zone shading — low PitchX = near batter (top=Yorker), high PitchX = near bowler (bottom=Short)
    ZONES = [
        (0.0,  2.0, 'rgba(50,120,210,0.20)',  'Yorker'),
        (2.0,  4.5, 'rgba(60,170,60,0.17)',   'Full'),
        (4.5,  7.0, 'rgba(210,195,0,0.17)',   'Length'),
        (7.0,  9.0, 'rgba(225,140,30,0.17)',  'Back of Len'),
        (9.0, 12.0, 'rgba(210,70,70,0.17)',   'Short'),
    ]
    zone_svgs = ''
    for x_lo, x_hi, color, label in ZONES:
        y_top = max(ty(x_lo), PT)
        y_bot = min(ty(x_hi), PT + ph)
        h_z   = y_bot - y_top
        if h_z > 0:
            mid_y = (y_top + y_bot) / 2 + 2.5
            zone_svgs += (
                f'<rect x="{PL}" y="{y_top:.1f}" width="{pw}" height="{h_z:.1f}" fill="{color}"/>'
                f'<text x="{PL + pw - 3:.1f}" y="{mid_y:.1f}" font-family="sans-serif" '
                f'font-size="5.8" fill="rgba(0,0,0,0.38)" text-anchor="end">{label}</text>'
            )

    # Crease lines and stumps drawn over zone shading
    pitch_overlay = (
        f'<line x1="{PL}" y1="{y_bat:.1f}" x2="{PL+pw}" y2="{y_bat:.1f}" stroke="white" stroke-width="1.5"/>'
        f'<line x1="{PL}" y1="{y_bowl:.1f}" x2="{PL+pw}" y2="{y_bowl:.1f}" stroke="white" stroke-width="1.5"/>'
        f'<line x1="{cx_:.1f}" y1="{PT}" x2="{cx_:.1f}" y2="{PT+ph}" stroke="white" stroke-width="0.5" stroke-dasharray="4,3" opacity="0.5"/>'
        f'<rect x="{cx_-5:.1f}" y="{y_bat:.1f}" width="10" height="4" fill="#c9a227" rx="1"/>'
        f'<rect x="{cx_-5:.1f}" y="{y_bowl:.1f}" width="10" height="4" fill="#c9a227" rx="1"/>'
    )

    circles = []
    for _, row in df.iterrows():
        x    = tx(float(row['PitchY']))
        y    = ty(float(row['PitchX']))
        runs = int(row.get('Runs', 0) or 0)
        # Safe wicket check — guard against NaN slipping through
        wkt_raw = row.get('Wicket', '')
        wkt = '' if pd.isnull(wkt_raw) else str(wkt_raw).strip()
        is_wkt  = bool(wkt) and 'Run Out' not in wkt
        is_wide = 'Wide' in str(row.get('Extra', '') or '')

        if mode == 'dots_wkts':
            if is_wkt:
                circles.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="#e74c3c" fill-opacity="0.85"/>')
            elif runs == 0 and not is_wide:
                circles.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="#27ae60" fill-opacity="0.75"/>')
        elif mode == 'wickets':
            if is_wkt:
                circles.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="#e74c3c" fill-opacity="0.85"/>')
        else:
            if runs == 6:
                circles.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="#e74c3c" fill-opacity="0.85"/>')
            elif runs == 4:
                circles.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="#f39c12" fill-opacity="0.85"/>')

    if mode == 'dots_wkts':
        legend = (f'<circle cx="{PL+5}" cy="{H-8}" r="3" fill="#27ae60" fill-opacity="0.9"/>'
                  f'<text x="{PL+11}" y="{H-4}" font-family="sans-serif" font-size="7.5" fill="#555">Dot</text>'
                  f'<circle cx="{PL+38}" cy="{H-8}" r="3" fill="#e74c3c" fill-opacity="0.9"/>'
                  f'<text x="{PL+44}" y="{H-4}" font-family="sans-serif" font-size="7.5" fill="#555">Wicket</text>')
    elif mode == 'wickets':
        legend = (f'<circle cx="{PL+5}" cy="{H-8}" r="3" fill="#e74c3c" fill-opacity="0.9"/>'
                  f'<text x="{PL+11}" y="{H-4}" font-family="sans-serif" font-size="7.5" fill="#555">Wicket</text>')
    else:
        legend = (f'<circle cx="{PL+5}" cy="{H-8}" r="3" fill="#f39c12" fill-opacity="0.9"/>'
                  f'<text x="{PL+11}" y="{H-4}" font-family="sans-serif" font-size="7.5" fill="#555">Four</text>'
                  f'<circle cx="{PL+36}" cy="{H-8}" r="3" fill="#e74c3c" fill-opacity="0.9"/>'
                  f'<text x="{PL+42}" y="{H-4}" font-family="sans-serif" font-size="7.5" fill="#555">Six</text>')

    body = pitch_base + zone_svgs + pitch_overlay + ''.join(circles) + legend
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" '
            f'style="width:100%;display:block;border-radius:6px;border:1px solid #dde3ed;background:#f0f4f0;">'
            f'{body}</svg>')


def _make_wagon_wheel_svg(df, boundaries_only=False):
    """
    Return an SVG wagon wheel. df must have numeric FieldX, FieldY (NaN stripped).
    Field grid ≈ 350×350; centre ≈ (175, 175).
    If boundaries_only=True, only 4s and 6s are drawn.
    """
    if boundaries_only:
        df = df[df['Runs'] >= 4]

    W, H   = 220, 220
    cx, cy = W / 2, H / 2
    R      = min(W, H) / 2 - 8
    FS     = 350.0

    def fx(v): return cx + (v - FS/2) / (FS/2) * R
    def fy(v): return cy + (v - FS/2) / (FS/2) * R

    field = (
        f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{R:.0f}" fill="#eaf4ea" stroke="#bbb" stroke-width="1"/>'
        f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{R*0.65:.0f}" fill="none" stroke="#ccc" stroke-width="0.5" stroke-dasharray="3,3"/>'
        f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{R*0.3:.0f}" fill="none" stroke="#ccc" stroke-width="0.5" stroke-dasharray="3,3"/>'
        f'<line x1="{cx:.0f}" y1="{cy-R:.0f}" x2="{cx:.0f}" y2="{cy+R:.0f}" stroke="#ddd" stroke-width="0.5"/>'
        f'<line x1="{cx-R:.0f}" y1="{cy:.0f}" x2="{cx+R:.0f}" y2="{cy:.0f}" stroke="#ddd" stroke-width="0.5"/>'
    )

    lines_, dots_ = [], []
    for _, row in df.iterrows():
        sx = fx(float(row['FieldX']))
        sy = fy(float(row['FieldY']))
        runs = int(row.get('Runs', 0) or 0)
        if boundaries_only:
            col = '#f39c12' if runs == 4 else '#e74c3c'
        else:
            col = '#27ae60' if runs == 0 else ('#f39c12' if runs == 4 else ('#e74c3c' if runs >= 6 else '#3498db'))
        lines_.append(f'<line x1="{cx:.0f}" y1="{cy:.0f}" x2="{sx:.1f}" y2="{sy:.1f}" stroke="{col}" stroke-width="1.2" stroke-opacity="0.45"/>')
        dots_.append(f'<circle cx="{sx:.1f}" cy="{sy:.1f}" r="2.5" fill="{col}" fill-opacity="0.85"/>')

    lx = W - 72
    if boundaries_only:
        legend = (
            f'<circle cx="{lx+14}" cy="{H-8}" r="3" fill="#f39c12"/>'
            f'<text x="{lx+20}" y="{H-4}" font-family="sans-serif" font-size="7" fill="#555">Four</text>'
            f'<circle cx="{lx+50}" cy="{H-8}" r="3" fill="#e74c3c"/>'
            f'<text x="{lx+56}" y="{H-4}" font-family="sans-serif" font-size="7" fill="#555">Six</text>'
        )
    else:
        legend = (
            f'<circle cx="{lx}" cy="{H-16}" r="3" fill="#27ae60"/>'
            f'<text x="{lx+6}" y="{H-12}" font-family="sans-serif" font-size="7" fill="#555">Dot</text>'
            f'<circle cx="{lx+28}" cy="{H-16}" r="3" fill="#3498db"/>'
            f'<text x="{lx+34}" y="{H-12}" font-family="sans-serif" font-size="7" fill="#555">1-3</text>'
            f'<circle cx="{lx}" cy="{H-5}" r="3" fill="#f39c12"/>'
            f'<text x="{lx+6}" y="{H-1}" font-family="sans-serif" font-size="7" fill="#555">4</text>'
            f'<circle cx="{lx+28}" cy="{H-5}" r="3" fill="#e74c3c"/>'
            f'<text x="{lx+34}" y="{H-1}" font-family="sans-serif" font-size="7" fill="#555">6</text>'
        )

    body = field + ''.join(lines_) + ''.join(dots_) + f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="3" fill="#555"/>' + legend
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" '
            f'style="width:100%;display:block;border-radius:6px;border:1px solid #dde3ed;">'
            f'{body}</svg>')


def _svg_or_drop(df, mode_or_none, card_id, dtype, label):
    """Return a generated SVG (pitchmap or wagon wheel) or fall back to a drop zone."""
    if mode_or_none in ('dots_wkts', 'boundaries', 'wickets'):
        d = _prep_coord_df(df, 'PitchX', 'PitchY')
        if d is not None:
            return _make_pitchmap_svg(d, mode_or_none)
    elif mode_or_none == 'wagon_boundaries':
        d = _prep_coord_df(df, 'FieldX', 'FieldY')
        if d is not None:
            return _make_wagon_wheel_svg(d, boundaries_only=True)
    else:  # wagon_wheel
        d = _prep_coord_df(df, 'FieldX', 'FieldY')
        if d is not None:
            return _make_wagon_wheel_svg(d)
    return drop_zone(card_id, dtype, label)


def _titled_map(title, svg_html):
    return (f'<div style="display:flex;flex-direction:column;gap:6px;">'
            f'<div class="pitch-map-title">{title}</div>'
            f'{svg_html}</div>')

# ─── BATTING SECTION (seam or spin) ──────────────────────────────────────────

def _batting_section(card_id, type_, overall, breakdown_rows, images):
    is_seam   = (type_ == 'seam')
    icon      = '🔴' if is_seam else '🟣'
    hdr_label = 'vs Seam' if is_seam else 'vs Spin'
    hdr_cls   = 'seam-header' if is_seam else 'spin-header'

    return f'''
    <div class="bowling-type-section">
      <div class="bowling-type-header {hdr_cls}"><span>{icon}</span> {hdr_label}</div>
      <div class="bowling-type-body">

        <div class="stats-overall-bar">
          <div class="overall-chip">
            <span class="chip-label">Runs(Balls)</span>
            <span class="chip-val">{fmt_rb(overall)}</span>
          </div>
          <div class="overall-chip">
            <span class="chip-label">Strike Rate</span>
            <span class="chip-val">{fmt_sr(overall)}</span>
          </div>
          <div class="overall-chip">
            <span class="chip-label">Boundary %</span>
            <span class="chip-val">{fmt_bp(overall)}</span>
          </div>
        </div>

        <table class="breakdown-table">
          <thead><tr><th>Type</th><th>Runs(Balls)</th><th>SR</th><th>Bdry%</th><th>Dot%</th><th>Wkts</th></tr></thead>
          <tbody>{breakdown_rows}</tbody>
        </table>

        <div class="images-three-col">
          {images}
        </div>

        <div class="card-notes" style="margin-top:12px;">
          <div class="notes-panel strength">
            <h4>+ Strengths {hdr_label}</h4>
            {ta(f'Scoring areas, favoured shots, lines to avoid {hdr_label}...')}
          </div>
          <div class="notes-panel weakness">
            <h4>- Weaknesses / Dismissal Patterns {hdr_label}</h4>
            {ta(f'Dismissal patterns, technical weaknesses, danger deliveries {hdr_label}...')}
          </div>
        </div>

      </div>
    </div>'''

# ─── BATSMAN CARD ─────────────────────────────────────────────────────────────

def build_batsman_card(n, stats):
    card_id  = f'bat{n}'
    name     = esc(stats['name'])
    hand     = stats.get('hand', '')

    hand_opts = ''
    for h in ['', 'Right hand bat', 'Left hand bat']:
        sel = ' selected' if h == hand else ''
        hand_opts += f'<option{sel}>{esc(h)}</option>'

    # ── SEAM section ──────────────────────────────────────────────────────────
    seam_all = stats['seam']['overall']
    seam_r   = stats['seam']['right']
    seam_l   = stats['seam']['left']
    seam_df  = stats['seam'].get('df')

    seam_rows = ''
    for label, s in [('Right arm seam', seam_r), ('Left arm seam', seam_l)]:
        seam_rows += f'''
          <tr>
            <td class="breakdown-type">{label}</td>
            {stat_cell(fmt_rb(s))}
            {stat_cell(fmt_sr(s))}
            {stat_cell(fmt_bp(s))}
            {stat_cell(fmt_dp(s))}
            {stat_cell(fmt_wk(s))}
          </tr>'''

    seam_imgs = (
        _titled_map('Wickets', _svg_or_drop(seam_df, 'wickets', card_id, 'seam_dw', 'Wickets Pitchmap')) +
        _titled_map('Boundaries', _svg_or_drop(seam_df, 'boundaries', card_id, 'seam_bd', 'Boundaries Pitchmap')) +
        _titled_map('Boundaries Wagon Wheel', _svg_or_drop(seam_df, 'wagon_boundaries', card_id, 'seam_ww', 'Boundaries Wagon Wheel'))
    )

    seam_section = _batting_section(card_id, 'seam', seam_all, seam_rows, seam_imgs)

    # ── SPIN section ──────────────────────────────────────────────────────────
    spin_all = stats['spin']['overall']
    spin_sub = stats['spin']['sub']
    spin_df  = stats['spin'].get('df')

    spin_rows = ''
    for code, label in SPIN_LABELS.items():
        s = spin_sub.get(code)
        if s and s['balls'] > 0:
            spin_rows += f'''
          <tr>
            <td class="breakdown-type">{label}</td>
            {stat_cell(fmt_rb(s))}
            {stat_cell(fmt_sr(s))}
            {stat_cell(fmt_bp(s))}
            {stat_cell(fmt_dp(s))}
            {stat_cell(fmt_wk(s))}
          </tr>'''

    spin_imgs = (
        _titled_map('Wickets', _svg_or_drop(spin_df, 'wickets', card_id, 'spin_dw', 'Wickets Pitchmap')) +
        _titled_map('Boundaries', _svg_or_drop(spin_df, 'boundaries', card_id, 'spin_bd', 'Boundaries Pitchmap')) +
        _titled_map('Boundaries Wagon Wheel', _svg_or_drop(spin_df, 'wagon_boundaries', card_id, 'spin_ww', 'Boundaries Wagon Wheel'))
    )

    spin_section = _batting_section(card_id, 'spin', spin_all, spin_rows, spin_imgs)

    return f'''
  <div class="player-card" id="batsman-{n}">
    <button class="delete-card-btn no-print"
      onclick="this.closest('.player-card').remove()" title="Remove">✕</button>
    <div class="player-card-inner">
      <div class="card-header">
        <div class="player-number">{n}</div>
        <div class="player-info">
          <input class="player-name-input" value="{name}" placeholder="Batsman Name">
          <div class="player-meta-row">
            <select title="Batting hand">{hand_opts}</select>
            <input type="text" placeholder="Role e.g. Opener / Middle order"
              style="width:220px;border:1.5px solid var(--grey-border);border-radius:5px;
                     padding:5px 10px;font-family:'DM Sans',sans-serif;font-size:12px;">
            <button class="threat-badge threat-UNKNOWN" onclick="cycleThreat(this)">THREAT: UNKNOWN</button>
          </div>
        </div>
      </div>
      {seam_section}
      {spin_section}
    </div>
  </div>'''

# ─── BOWLER CARD ──────────────────────────────────────────────────────────────

SECTION_H4 = ('style="font-size:11px;text-transform:uppercase;letter-spacing:0.8px;'
               'color:var(--navy);font-weight:700;margin-bottom:10px;padding:5px 8px 5px 10px;'
               'background:rgba(0,112,60,0.07);border-radius:4px;border-left:3px solid var(--navy);"')

def build_bowler_card(n, stats):
    card_id  = f'bowl{n}'
    name     = esc(stats['name'])
    bt_full  = stats.get('bowl_type', '')
    bt_raw   = stats.get('bowl_type_raw', '')

    bowl_opts = ''
    for opt in BOWL_TYPE_OPTIONS:
        sel = ' selected' if opt == bt_full else ''
        bowl_opts += f'<option{sel}>{esc(opt)}</option>'

    def ms(val):
        return esc(str(val)) if val not in (0, 0.0, '') else '—'

    return f'''
  <div class="bowler-card" id="bowler-{n}">
    <button class="delete-card-btn no-print"
      onclick="this.closest('.bowler-card').remove()" title="Remove">✕</button>
    <div class="bowler-card-inner-wrap">
      <div class="card-header">
        <div class="bowler-number">{n}</div>
        <div class="bowler-info">
          <input class="bowler-name-input" value="{name}" placeholder="Bowler Name">
          <div class="bowler-meta-row">
            <select title="Bowling type">{bowl_opts}</select>
            <input type="text" placeholder="Role e.g. Opening bowler / 1st change"
              style="width:230px;border:1.5px solid var(--grey-border);border-radius:5px;
                     padding:5px 10px;font-family:'DM Sans',sans-serif;font-size:12px;">
            <button class="threat-badge threat-UNKNOWN" onclick="cycleThreat(this)">THREAT: UNKNOWN</button>
          </div>
        </div>
      </div>

      <div class="bowler-overall-stats" style="margin-bottom:20px;">
        <h4 {SECTION_H4}>Overall Bowling Stats</h4>
        <div class="mini-stats cols-5">
          <div class="mini-stat"><label>Overs</label>
            <input value="{esc(stats['overs'])}" placeholder="e.g. 24.3"></div>
          <div class="mini-stat"><label>Wickets</label>
            <input value="{ms(stats['wickets'])}" placeholder="e.g. 8"></div>
          <div class="mini-stat"><label>Average</label>
            <input value="{ms(stats['average'])}" placeholder="e.g. 22.4"></div>
          <div class="mini-stat"><label>Economy</label>
            <input value="{ms(stats['economy'])}" placeholder="e.g. 7.2"></div>
          <div class="mini-stat"><label>Strike Rate</label>
            <input value="{ms(stats['sr'])}" placeholder="e.g. 18.6"></div>
        </div>
        <div class="mini-stats cols-4" style="margin-top:8px;">
          <div class="mini-stat"><label>Dot %</label>
            <input value="{ms(stats['dot_pct'])}{'%' if stats['dot_pct'] else ''}" placeholder="e.g. 42%"></div>
          <div class="mini-stat"><label>Boundary %</label>
            <input value="{ms(stats['boundary_pct'])}{'%' if stats['boundary_pct'] else ''}" placeholder="e.g. 14%"></div>
          <div class="mini-stat"><label>Best Figures</label>
            <input value="{esc(stats['best'])}" placeholder="e.g. 3/28"></div>
          <div class="mini-stat"><label>Matches</label>
            <input value="{ms(stats['matches'])}" placeholder="e.g. 8"></div>
        </div>
      </div>

      <div style="margin-bottom:20px;">
        <h4 {SECTION_H4}>Pitch Maps</h4>
        <div class="images-two-col">
          {_titled_map('Wickets', _svg_or_drop(stats.get('df'), 'wickets', card_id, 'dw_pitchmap', 'Wickets Pitchmap'))}
          {_titled_map('Boundaries Conceded', _svg_or_drop(stats.get('df'), 'boundaries', card_id, 'bd_pitchmap', 'Pitch Map'))}
        </div>
      </div>

      <div class="card-notes" style="margin-bottom:16px;">
        <div class="notes-panel strength">
          <h4>+ Strengths</h4>
          {ta('What makes this bowler dangerous? Best deliveries, best conditions...')}
        </div>
        <div class="notes-panel weakness">
          <h4>- Weaknesses</h4>
          {ta('When does this bowler struggle? Lengths and lines to target...')}
        </div>
      </div>

      <div class="bowling-plan-area">
        <h4>Batting Plan Against This Bowler</h4>
        {ta('How should we approach this bowler? Target overs, matchups, scoring areas, when to attack vs defend...')}
      </div>
    </div>
  </div>'''

# ─── CSS ──────────────────────────────────────────────────────────────────────

CSS = r"""
<style>
  :root {
    --navy: #00703C;
    --navy-mid: #005a30;
    --accent: #c9a227;
    --accent-light: #f5e89a;
    --red: #c0392b;
    --red-light: #f9e8e6;
    --green: #1a7a4a;
    --green-light: #e6f5ed;
    --amber: #d4820a;
    --grey-bg: #f4f6f9;
    --grey-border: #dde3ed;
    --grey-mid: #8a95a8;
    --white: #ffffff;
    --text: #1c2533;
    --text-light: #5a6678;
    --bowling-red: #c8102e;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'DM Sans', sans-serif; background: var(--grey-bg); color: var(--text); font-size: 14px; line-height: 1.5; }

  /* TOOLBAR */
  .toolbar { background: var(--navy); color: var(--white); padding: 14px 24px; display: flex; align-items: center; gap: 16px; flex-wrap: wrap; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 12px rgba(0,0,0,.25); }
  .toolbar-brand { font-family: 'DM Serif Display', serif; font-size: 18px; color: var(--accent); margin-right: auto; letter-spacing: .3px; }
  .btn { padding: 8px 18px; border-radius: 6px; border: none; cursor: pointer; font-family: 'DM Sans', sans-serif; font-weight: 600; font-size: 13px; transition: all .15s; }
  .btn-primary { background: var(--accent); color: var(--navy); }
  .btn-primary:hover { background: var(--accent-light); }
  .btn-outline { background: transparent; color: var(--white); border: 1.5px solid rgba(255,255,255,.3); }
  .btn-outline:hover { border-color: var(--accent); color: var(--accent); }
  .btn-bat  { background: #1a1a1a; color: #c9a227; }
  .btn-bat:hover  { background: #333; }
  .btn-bowl { background: var(--bowling-red); color: white; }
  .btn-bowl:hover { background: #a00d25; }

  /* DOCUMENT */
  .document { max-width: 960px; margin: 28px auto; padding: 0 16px 60px; }

  /* MATCH HEADER */
  .match-header { background: var(--navy); border-radius: 12px 12px 0 0; padding: 32px 36px 28px; color: white; position: relative; overflow: hidden; }
  .match-header::before { content:''; position:absolute; top:-40px; right:-40px; width:200px; height:200px; border-radius:50%; background:rgba(200,168,75,.08); pointer-events:none; }
  .header-top { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px; margin-bottom: 20px; }
  .header-left .subtitle { color: var(--accent); font-size: 13px; font-weight: 500; letter-spacing: .5px; text-transform: uppercase; }
  .header-meta { display: flex; gap: 24px; flex-wrap: wrap; }
  .meta-item { text-align: right; }
  .meta-item label { font-size: 11px; text-transform: uppercase; letter-spacing: .6px; color: rgba(255,255,255,.5); display: block; margin-bottom: 2px; }
  .meta-item .meta-val { font-size: 15px; font-weight: 600; color: white; background: transparent; border: none; border-bottom: 1.5px solid rgba(255,255,255,.2); font-family: 'DM Sans', sans-serif; text-align: right; width: 160px; padding: 2px 4px; }
  .meta-item .meta-val:focus { outline: none; border-bottom-color: var(--accent); }
  .team-input { background: transparent; border: none; border-bottom: 2px solid rgba(255,255,255,.3); color: white; font-family: 'DM Serif Display', serif; font-size: 28px; width: 100%; padding: 2px 0; }
  .team-input:focus { outline: none; border-bottom-color: var(--accent); }
  .team-input::placeholder { color: rgba(255,255,255,.35); }
  .header-summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px,1fr)); gap: 12px; background: rgba(255,255,255,.06); border-radius: 8px; padding: 16px 20px; margin-top: 8px; }
  .summary-stat label { font-size: 10px; text-transform: uppercase; letter-spacing: .7px; color: rgba(255,255,255,.5); display: block; margin-bottom: 4px; }
  .summary-stat input { background: transparent; border: none; border-bottom: 1px solid rgba(255,255,255,.2); color: white; font-family: 'DM Sans', sans-serif; font-size: 20px; font-weight: 700; width: 100%; padding: 2px 0; }
  .summary-stat input:focus { outline: none; border-bottom-color: var(--accent); }

  /* SECTION HEADERS */
  .section-header { display: flex; align-items: center; gap: 10px; padding: 14px 36px; font-size: 11px; font-weight: 700; letter-spacing: 1.2px; text-transform: uppercase; color: white; }
  .section-header.batting { background: #1a1a1a; color: #c9a227; }
  .section-header.bowling-threats { background: var(--bowling-red); }
  .section-icon { font-size: 16px; }

  /* CARDS */
  .players-container, .bowlers-container { background: white; }
  .player-card, .bowler-card { border-bottom: 1px solid var(--grey-border); position: relative; }
  .player-card:last-child, .bowler-card:last-child { border-bottom: none; }
  .player-card-inner, .bowler-card-inner-wrap { padding: 24px 28px; }

  .card-header { display: flex; align-items: flex-start; gap: 16px; margin-bottom: 18px; flex-wrap: wrap; }
  .player-number, .bowler-number { width: 36px; height: 36px; border-radius: 50%; font-family: 'DM Serif Display', serif; font-size: 16px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 4px; }
  .player-number { background: #1a1a1a; color: #c9a227; }
  .bowler-number { background: var(--bowling-red); color: white; }
  .player-info, .bowler-info { flex: 1; min-width: 200px; }
  .player-name-input, .bowler-name-input { font-family: 'DM Serif Display', serif; font-size: 20px; border: none; border-bottom: 2px solid var(--grey-border); width: 100%; padding: 2px 4px; background: transparent; }
  .player-name-input { color: var(--navy); }
  .bowler-name-input { color: var(--bowling-red); }
  .player-name-input:focus { outline: none; border-bottom-color: var(--navy); }
  .bowler-name-input:focus { outline: none; border-bottom-color: var(--bowling-red); }
  .player-name-input::placeholder, .bowler-name-input::placeholder { color: #bbb; font-size: 18px; }

  .player-meta-row, .bowler-meta-row { display: flex; gap: 10px; margin-top: 8px; flex-wrap: wrap; align-items: center; }
  .player-meta-row select, .bowler-meta-row select { border: 1.5px solid var(--grey-border); border-radius: 5px; padding: 5px 10px; font-family: 'DM Sans', sans-serif; font-size: 12px; color: var(--text); background: var(--grey-bg); }

  .threat-badge { padding: 5px 14px; border-radius: 20px; font-size: 11px; font-weight: 700; letter-spacing: .8px; text-transform: uppercase; cursor: pointer; border: none; font-family: 'DM Sans', sans-serif; }
  .threat-HIGH   { background: var(--red);   color: white; }
  .threat-MEDIUM { background: var(--amber); color: white; }
  .threat-LOW    { background: var(--green); color: white; }
  .threat-UNKNOWN { background: #ccc; color: #555; }

  .delete-card-btn { position: absolute; top: 16px; right: 20px; background: none; border: none; color: var(--grey-mid); cursor: pointer; font-size: 18px; line-height: 1; padding: 4px 6px; border-radius: 4px; transition: all .15s; }
  .delete-card-btn:hover { background: var(--red-light); color: var(--red); }

  /* BOWLING TYPE SECTION */
  .bowling-type-section { margin-bottom: 20px; border: 1.5px solid var(--grey-border); border-radius: 8px; overflow: hidden; }
  .bowling-type-section:last-child { margin-bottom: 0; }
  .bowling-type-header { display: flex; align-items: center; gap: 10px; padding: 10px 16px; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }
  .seam-header { background: #1a1a1a; color: #c9a227; border-left: 4px solid #c9a227; }
  .spin-header { background: #1a1a1a; color: #c9a227; border-left: 4px solid #c9a227; }
  .bowling-type-body { padding: 16px; background: white; }

  /* OVERALL STATS BAR */
  .stats-overall-bar { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 12px; }
  .overall-chip { background: var(--grey-bg); border-radius: 8px; padding: 10px 16px; display: flex; flex-direction: column; align-items: center; min-width: 100px; flex: 1; }
  .chip-label { font-size: 9px; text-transform: uppercase; letter-spacing: .7px; color: var(--grey-mid); margin-bottom: 4px; }
  .chip-val { font-size: 20px; font-weight: 700; color: var(--navy); }

  /* BREAKDOWN TABLE */
  .breakdown-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 14px; }
  .breakdown-table th { background: #1a1a1a; color: #c9a227; font-size: 9px; text-transform: uppercase; letter-spacing: .6px; padding: 6px 8px; text-align: left; font-weight: 600; }
  .breakdown-table td { padding: 6px 8px; border-bottom: 1px solid var(--grey-border); }
  .breakdown-table tr:last-child td { border-bottom: none; }
  .breakdown-table tr:nth-child(even) td { background: var(--grey-bg); }
  .breakdown-table .stat-val { font-weight: 600; color: var(--navy); }
  .breakdown-table .breakdown-type { color: var(--text-light); font-weight: 500; }

  /* IMAGE ZONES — 3-col for batting, 2-col for bowling */
  .images-three-col { display: grid; grid-template-columns: repeat(3,1fr); gap: 10px; margin-bottom: 4px; }
  .images-two-col   { display: grid; grid-template-columns: repeat(2,1fr); gap: 10px; margin-bottom: 4px; }
  @media(max-width:600px) { .images-three-col,.images-two-col { grid-template-columns: 1fr; } }

  .image-drop-zone { border: 2px dashed var(--grey-border); border-radius: 8px; aspect-ratio: 1/1; width: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; cursor: pointer; transition: all .2s; position: relative; overflow: hidden; background: var(--grey-bg); }
  .image-drop-zone:hover { border-color: var(--navy); background: #eef1f8; }
  .image-drop-zone.dragover { border-color: var(--accent); background: #fdf8ec; }
  .image-drop-zone .drop-label { font-size: 11px; color: var(--grey-mid); text-align: center; padding: 8px; pointer-events: none; }
  .image-drop-zone .drop-icon { font-size: 22px; margin-bottom: 4px; pointer-events: none; }
  .image-drop-zone img { width: 100%; height: 100%; object-fit: contain; position: absolute; top: 0; left: 0; border-radius: 6px; }
  .image-drop-zone .remove-img { position: absolute; top: 4px; right: 4px; background: rgba(0,0,0,.5); color: white; border: none; border-radius: 50%; width: 20px; height: 20px; cursor: pointer; font-size: 12px; display: none; align-items: center; justify-content: center; }
  .image-drop-zone:has(img):hover .remove-img { display: flex; }
  .image-drop-zone input[type="file"] { position: absolute; inset: 0; opacity: 0; cursor: pointer; pointer-events: none; }

  .pitch-map-title { font-size: 11px; text-transform: uppercase; letter-spacing: .8px; color: var(--text); font-weight: 700; margin-bottom: 8px; padding: 5px 8px 5px 10px; background: rgba(0,0,0,.05); border-radius: 4px; border-left: 3px solid var(--text); }

  /* NOTES */
  .card-notes { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  @media(max-width:600px) { .card-notes { grid-template-columns: 1fr; } }
  .notes-panel textarea { width: 100%; border: 1.5px solid var(--grey-border); border-radius: 6px; padding: 10px 12px; font-family: 'DM Sans', sans-serif; font-size: 13px; color: var(--text); resize: vertical; overflow: hidden; min-height: 80px; background: var(--grey-bg); transition: border-color .15s; }
  .notes-panel textarea:focus { outline: none; }
  .notes-panel.strength { background: #fdf9ec; border-radius: 8px; padding: 12px; }
  .notes-panel.weakness { background: #fdf2f2; border-radius: 8px; padding: 12px; }
  .notes-panel.strength textarea:focus { border-color: var(--green); }
  .notes-panel.weakness textarea:focus { border-color: var(--red); }
  .notes-panel.strength textarea { background: #fffdf5; }
  .notes-panel.weakness textarea { background: #fff8f8; }
  .notes-panel h4 { font-size: 11px; text-transform: uppercase; letter-spacing: .8px; font-weight: 700; margin-bottom: 10px; padding: 5px 8px 5px 10px; border-radius: 4px; border-left: 3px solid; }
  .notes-panel.strength h4 { color: var(--green); border-left-color: var(--green); background: rgba(26,122,74,.07); }
  .notes-panel.weakness h4 { color: var(--red);   border-left-color: var(--red);   background: rgba(192,57,43,.07); }

  /* BOWLING PLAN */
  .bowling-plan-area textarea { width: 100%; border: 1.5px solid #e8d0d0; border-radius: 6px; padding: 10px 12px; font-family: 'DM Sans', sans-serif; font-size: 13px; color: var(--text); resize: vertical; overflow: hidden; min-height: 80px; background: white; }
  .bowling-plan-area textarea:focus { outline: none; border-color: var(--bowling-red); }
  .bowling-plan-area h4 { font-size: 11px; text-transform: uppercase; letter-spacing: .8px; color: var(--bowling-red); font-weight: 700; margin-bottom: 10px; padding: 5px 8px 5px 10px; background: rgba(192,57,43,.07); border-radius: 4px; border-left: 3px solid var(--bowling-red); }

  /* MINI STATS (bowler card) */
  .mini-stats { display: grid; gap: 8px; margin-bottom: 12px; }
  .mini-stats.cols-5 { grid-template-columns: repeat(5,1fr); }
  .mini-stats.cols-4 { grid-template-columns: repeat(4,1fr); }
  @media(max-width:600px) { .mini-stats.cols-5,.mini-stats.cols-4 { grid-template-columns: repeat(3,1fr); } }
  .mini-stat { background: var(--grey-bg); border-radius: 6px; padding: 8px 10px; text-align: center; }
  .mini-stat label { font-size: 9px; text-transform: uppercase; letter-spacing: .6px; color: var(--grey-mid); display: block; margin-bottom: 3px; }
  .mini-stat input { background: transparent; border: none; text-align: center; font-family: 'DM Sans', sans-serif; font-size: 17px; font-weight: 700; color: var(--navy); width: 100%; }
  .mini-stat input:focus { outline: none; }
  .mini-stat input::placeholder { font-size: 13px; color: #bbb; font-weight: 400; }

  /* ADD ROW */
  .add-player-row, .add-bowler-row { padding: 20px 28px; background: white; border-top: 2px dashed var(--grey-border); text-align: center; }

  /* FOOTER */
  .doc-footer { background: var(--navy); border-radius: 0 0 12px 12px; padding: 16px 28px; display: flex; justify-content: space-between; align-items: center; color: rgba(255,255,255,.4); font-size: 11px; }
  .doc-footer input { background: transparent; border: none; color: rgba(255,255,255,.5); font-family: 'DM Sans', sans-serif; font-size: 11px; }

  /* PRINT */
  @page { size: A4 portrait; margin: 12mm 10mm; }
  @media print {
    .toolbar,.delete-card-btn,.add-player-row,.add-bowler-row,
    .image-drop-zone input[type="file"],.remove-img { display: none !important; }
    html,body { width:210mm; background:white !important; font-size:11px; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .document { width:190mm; max-width:190mm; margin:0 auto; padding:0; }
    .player-card,.bowler-card { break-inside:avoid; page-break-inside:avoid; }
    .match-header,.doc-footer { border-radius:0; }
    input,textarea,select { border:none !important; background:transparent !important; resize:none !important; }
    * { max-width:100% !important; overflow:visible !important; word-break:break-word; }
    img { max-width:100% !important; height:auto !important; }
    textarea { min-height:0 !important; }
  }
</style>
"""

# ─── JS ───────────────────────────────────────────────────────────────────────

JS = r"""
<script>
function autoResize(el){el.style.height='auto';el.style.height=el.scrollHeight+'px';}
document.addEventListener('input',e=>{if(e.target.tagName==='TEXTAREA')autoResize(e.target);});
document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('textarea').forEach(autoResize));

const threatCycle=['UNKNOWN','HIGH','MEDIUM','LOW'];
function cycleThreat(btn){
  const cur=btn.textContent.replace('THREAT: ','').trim();
  const next=threatCycle[(threatCycle.indexOf(cur)+1)%threatCycle.length];
  btn.textContent=`THREAT: ${next}`;
  btn.className=`threat-badge threat-${next}`;
}

function triggerFileInput(event,id,type){
  if(event.target.classList.contains('remove-img'))return;
  const fi=document.getElementById(`file-${id}-${type}`);
  if(fi)fi.click();
}
function setupDropZone(id,type){
  const zone=document.getElementById(`drop-${id}-${type}`);
  if(!zone)return;
  zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('dragover');});
  zone.addEventListener('dragleave',()=>zone.classList.remove('dragover'));
  zone.addEventListener('drop',e=>{
    e.preventDefault();zone.classList.remove('dragover');
    const file=e.dataTransfer.files[0];
    if(file&&file.type.startsWith('image/'))loadImg(id,type,file);
  });
}
function handleFileInput(event,id,type){const f=event.target.files[0];if(f)loadImg(id,type,f);}
function loadImg(id,type,file){
  const reader=new FileReader();
  reader.onload=e=>{
    const zone=document.getElementById(`drop-${id}-${type}`);
    const old=zone.querySelector('img');if(old)old.remove();
    const img=document.createElement('img');img.src=e.target.result;zone.appendChild(img);
    zone.querySelector('.drop-label').style.display='none';
    zone.querySelector('.drop-icon').style.display='none';
  };
  reader.readAsDataURL(file);
}
function clearImage(id,type){
  const zone=document.getElementById(`drop-${id}-${type}`);
  const img=zone.querySelector('img');if(img)img.remove();
  zone.querySelector('.drop-label').style.display='';
  zone.querySelector('.drop-icon').style.display='';
}
document.addEventListener('DOMContentLoaded',()=>{
  document.querySelectorAll('.image-drop-zone').forEach(zone=>{
    const m=zone.id.match(/^drop-(.+?)-(.+)$/);
    if(m)setupDropZone(m[1],m[2]);
  });
});

// Blank card adders
let xb=2000,xw=2000;
function addBlankBatsman(){
  xb++;
  const card=document.createElement('div');
  card.className='player-card';card.id=`batsman-${xb}`;
  card.innerHTML=`
    <button class="delete-card-btn no-print" onclick="this.closest('.player-card').remove()">✕</button>
    <div class="player-card-inner">
      <div class="card-header">
        <div class="player-number">${xb}</div>
        <div class="player-info">
          <input class="player-name-input" placeholder="Batsman Name">
          <div class="player-meta-row">
            <select><option>Right hand bat</option><option>Left hand bat</option></select>
            <input type="text" placeholder="Role e.g. Opener" style="width:220px;border:1.5px solid var(--grey-border);border-radius:5px;padding:5px 10px;font-family:'DM Sans',sans-serif;font-size:12px;">
            <button class="threat-badge threat-UNKNOWN" onclick="cycleThreat(this)">THREAT: UNKNOWN</button>
          </div>
        </div>
      </div>
    </div>`;
  document.getElementById('playersContainer').appendChild(card);
}
function addBlankBowler(){
  xw++;
  const card=document.createElement('div');
  card.className='bowler-card';card.id=`bowler-${xw}`;
  card.innerHTML=`
    <button class="delete-card-btn no-print" onclick="this.closest('.bowler-card').remove()">✕</button>
    <div class="bowler-card-inner-wrap">
      <div class="card-header">
        <div class="bowler-number">${xw}</div>
        <div class="bowler-info">
          <input class="bowler-name-input" placeholder="Bowler Name">
          <div class="bowler-meta-row">
            <select><option value="">Bowl type</option><option>Right arm fast</option><option>Right arm fast-medium</option><option>Right arm medium</option><option>Left arm fast</option><option>Left arm fast-medium</option><option>Left arm medium</option><option>Right arm off-break</option><option>Right arm leg-break</option><option>Left arm orthodox</option><option>Left arm wrist spin</option></select>
            <input type="text" placeholder="Role e.g. Opening bowler" style="width:230px;border:1.5px solid var(--grey-border);border-radius:5px;padding:5px 10px;font-family:'DM Sans',sans-serif;font-size:12px;">
            <button class="threat-badge threat-UNKNOWN" onclick="cycleThreat(this)">THREAT: UNKNOWN</button>
          </div>
        </div>
      </div>
    </div>`;
  document.getElementById('bowlersContainer').appendChild(card);
}
</script>
"""

# ─── HTML ASSEMBLY ────────────────────────────────────────────────────────────

def generate_html(bat_stats_list, bowl_stats_list, meta):
    opp   = esc(meta.get('opposition', ''))
    date  = esc(meta.get('date', ''))
    venue = esc(meta.get('venue', ''))

    bat_cards  = '\n'.join(build_batsman_card(i+1, s) for i,s in enumerate(bat_stats_list))
    bowl_cards = '\n'.join(build_bowler_card(i+1, s)  for i,s in enumerate(bowl_stats_list))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pre-Game Analysis Pack{(' — ' + opp) if opp else ''}</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
{CSS}
</head>
<body>

<div class="toolbar">
  <div class="toolbar-brand">Pre-Game Pack Builder</div>
  <button class="btn btn-bat"  onclick="addBlankBatsman()">+ Add Batsman</button>
  <button class="btn btn-bowl" onclick="addBlankBowler()">+ Add Bowler</button>
  <button class="btn btn-outline" onclick="window.print()">Print / Export PDF</button>
</div>

<div class="document">

  <div class="match-header">
    <div class="header-top">
      <div class="header-left" style="flex:1;min-width:220px;">
        <div class="subtitle">Pre-Game Opposition Analysis</div>
        <input class="team-input" placeholder="Opposition Team Name" value="{opp}">
      </div>
      <div class="header-meta">
        <div class="meta-item"><label>Date</label>
          <input class="meta-val" type="text" placeholder="DD/MM/YYYY" value="{date}"></div>
        <div class="meta-item"><label>Venue</label>
          <input class="meta-val" type="text" placeholder="Ground name" value="{venue}"></div>
        <div class="meta-item"><label>Format</label>
          <input class="meta-val" type="text" placeholder="T20 / 50-over"></div>
      </div>
    </div>
    <div class="header-summary">
      <div class="summary-stat"><label>Competition</label>
        <input type="text" placeholder="e.g. U18 County T20"></div>
    </div>
  </div>

  <div class="section-header batting"><span class="section-icon">🏏</span> Player Profiles — Batting Threats</div>
  <div class="players-container" id="playersContainer">
    {bat_cards if bat_cards.strip() else '<div style="padding:24px 28px;color:var(--grey-mid);font-size:13px;">No batters selected.</div>'}
  </div>
  <div class="add-player-row no-print">
    <button class="btn btn-bat" onclick="addBlankBatsman()">+ Add Batsman Card</button>
  </div>

  <div class="section-header bowling-threats"><span class="section-icon">🎳</span> Player Profiles — Bowling Threats</div>
  <div class="bowlers-container" id="bowlersContainer">
    {bowl_cards if bowl_cards.strip() else '<div style="padding:24px 28px;color:var(--grey-mid);font-size:13px;">No bowlers selected.</div>'}
  </div>
  <div class="add-bowler-row no-print">
    <button class="btn btn-bowl" onclick="addBlankBowler()">+ Add Bowler Card</button>
  </div>

  <div class="doc-footer">
    <span>Prepared by: <input type="text" placeholder="Analyst name" style="width:140px;"></span>
    <span>Pre-Game Analysis Pack — <input type="text" placeholder="Season / Competition" style="width:180px;"></span>
    <span>CONFIDENTIAL</span>
  </div>

</div>

{JS}
</body>
</html>"""

# ─── FILENAME PARSING ─────────────────────────────────────────────────────────

import re

_ROLE_KEYWORDS = {
    'batting': 'Batsman', 'batsman': 'Batsman', 'batter': 'Batsman',
    'bowling': 'Bowler',  'bowler':  'Bowler',
    'allrounder': 'All-rounder', 'all-rounder': 'All-rounder', 'all rounder': 'All-rounder',
}

def parse_filename(path):
    """
    Parse player name and role from a filename.
    e.g. "James Thornton batting Notts.csv" → ('James Thornton', 'Batsman')
         "Marcus Webb bowling Notts.csv"     → ('Marcus Webb', 'Bowler')
    Returns (player_name, role) or (None, None) if not parseable.
    """
    stem = Path(path).stem  # original case for name extraction
    stem_lower = stem.lower()

    # Try each keyword; split on first match
    for kw, role in sorted(_ROLE_KEYWORDS.items(), key=lambda x: -len(x[0])):
        idx = stem_lower.find(kw)
        if idx != -1:
            player_name = stem[:idx].strip()
            if player_name:
                return player_name, role
    return None, None

# ─── PLAYER SELECTION ─────────────────────────────────────────────────────────

def select_players(df, source_label):
    """
    Fallback interactive selection for files with no parseable player name.
    Returns list of {'name', 'role'} dicts.
    """
    all_batters = sorted(df['Batter'].replace('', pd.NA).dropna().unique())
    all_bowlers = sorted(df['Bowler'].replace('', pd.NA).dropna().unique())
    all_players = sorted(set(all_batters) | set(all_bowlers))

    print(f"\n{'─'*65}")
    print(f"Players found in: {source_label}")
    print(f"{'─'*65}")
    for i, name in enumerate(all_players, 1):
        roles = []
        if name in all_batters: roles.append('bat')
        if name in all_bowlers: roles.append('bowl')
        print(f"  {i:>3}.  {name:<35}  [{'/'.join(roles)}]")
    print(f"{'─'*65}")
    print("Enter numbers to include (e.g. 1,3,5 or 1-4 or all):")

    while True:
        raw = input("  > ").strip().lower()
        if not raw:
            print("  (skipping)")
            return []
        if raw == 'all':
            indices = list(range(1, len(all_players)+1))
        else:
            try:
                indices = set()
                for part in raw.split(','):
                    part = part.strip()
                    if '-' in part:
                        a, b = part.split('-', 1)
                        indices.update(range(int(a), int(b)+1))
                    else:
                        indices.add(int(part))
                indices = sorted(indices)
            except ValueError:
                print("  Invalid — try again.")
                continue

        selected_names = [all_players[i-1] for i in indices if 1 <= i <= len(all_players)]
        if not selected_names:
            print("  Nothing selected.")
            return []

        result = []
        print()
        for name in selected_names:
            in_bat  = name in all_batters
            in_bowl = name in all_bowlers
            if in_bat and in_bowl:
                r = input(f"  {name} appears as both batter and bowler. Role? [b]atsman / [w]owler / [a]ll-rounder: ").strip().lower()
                role = 'All-rounder' if r.startswith('a') else ('Bowler' if r.startswith('w') else 'Batsman')
            elif in_bat:
                role = 'Batsman'
            else:
                role = 'Bowler'
            result.append({'name': name, 'role': role})
            print(f"    ✓ {name} → {role}")

        return result

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Generate a Pre-Game Cricket Analysis Pack from ball-by-ball CSV data.')
    parser.add_argument('csv_files', nargs='*', metavar='CSV',
                        help='One or more ball-by-ball CSV files (or use -d for a folder)')
    parser.add_argument('-d', '--folder', metavar='FOLDER',
                        help='Folder containing CSV files — loads all .csv files in it')
    parser.add_argument('-o', '--output', default='pre_game_pack.html',
                        help='Output HTML file (default: pre_game_pack.html)')
    args = parser.parse_args()

    # Resolve file list
    if args.folder:
        folder = Path(args.folder)
        if not folder.is_dir():
            print(f"ERROR: {args.folder} is not a directory", file=sys.stderr)
            sys.exit(1)
        csv_paths = sorted(folder.glob('*.csv'))
        if not csv_paths:
            print(f"ERROR: no .csv files found in {args.folder}", file=sys.stderr)
            sys.exit(1)
    elif args.csv_files:
        csv_paths = [Path(p) for p in args.csv_files]
    else:
        parser.print_help()
        sys.exit(1)

    # Load all CSVs, parsing player name + role from each filename
    frames = []
    auto_selected = []   # list of {'name', 'role'} parsed from filenames
    bowling_dfs = {}     # player_name (lower) → DataFrame from their bowling CSV
    untagged_frames = [] # files where we couldn't parse a name — fall back to interactive

    for p in csv_paths:
        if not p.exists():
            print(f"ERROR: {p} not found", file=sys.stderr)
            sys.exit(1)
        player_name, role = parse_filename(p)
        csv_df = load_csv(p)
        if player_name and role:
            print(f"Loading {p.name}…  → {player_name} ({role})")
            # Keep the bowling CSV DataFrame for this player
            if role == 'Bowler':
                bowling_dfs[player_name.lower()] = csv_df
            # Check for duplicate: same name, different role → upgrade to All-rounder
            existing = next((x for x in auto_selected if x['name'].lower() == player_name.lower()), None)
            if existing:
                if existing['role'] != role:
                    existing['role'] = 'All-rounder'
            else:
                auto_selected.append({'name': player_name, 'role': role})
        else:
            print(f"Loading {p.name}…  (no player name detected — will prompt)")
            untagged_frames.append(p)
        frames.append(csv_df)

    df = pd.concat(frames, ignore_index=True)
    print(f"Loaded {len(df):,} delivery rows from {len(frames)} file(s).")

    # Pull match metadata from the data
    matches = df['Match'].dropna().replace('', pd.NA).dropna()
    last_match = matches.iloc[-1] if not matches.empty else ''
    venue  = df['Venue'].dropna().replace('', pd.NA).dropna().iloc[-1] if 'Venue' in df.columns else ''
    date   = df['Date'].dropna().replace('', pd.NA).dropna().iloc[-1]  if 'Date'  in df.columns else ''

    # Try to extract opposition team name from match title
    opp = ''
    if last_match:
        parts = str(last_match).split(' v ')
        if len(parts) == 2:
            opp = parts[1].strip()

    meta = {'opposition': opp, 'date': date, 'venue': venue}

    # Match parsed names to actual CSV names (case-insensitive)
    all_csv_names = set(
        n.strip() for col in ('Batter', 'Bowler') if col in df.columns
        for n in df[col].replace('', pd.NA).dropna().unique()
    )
    name_map = {n.lower(): n for n in all_csv_names}

    selected = []
    print()
    for item in auto_selected:
        matched = name_map.get(item['name'].lower())
        if matched:
            print(f"  ✓ {matched} → {item['role']}")
            selected.append({'name': matched, 'role': item['role']})
        else:
            print(f"  WARNING: '{item['name']}' from filename not found in CSV data — skipping")

    # If any files had no parseable name, fall back to interactive for those
    if untagged_frames:
        untagged_df = pd.concat([load_csv(p) for p in untagged_frames], ignore_index=True)
        extra = select_players(untagged_df, f"{len(untagged_frames)} untagged file(s)")
        selected.extend(extra)

    if not selected:
        print("No players selected. Nothing to generate.")
        sys.exit(0)

    # Compute stats and build cards
    bat_stats_list  = []
    bowl_stats_list = []

    for item in selected:
        name = item['name']
        role = item['role']

        if role in ('Batsman', 'All-rounder'):
            s = compute_batting_stats(df, name)
            if s:
                bat_stats_list.append(s)
            else:
                print(f"  WARNING: no batting data found for {name}")

        if role in ('Bowler', 'All-rounder'):
            bowl_df = bowling_dfs.get(name.lower(), df)
            s = compute_bowling_stats(bowl_df, name)
            if s:
                bowl_stats_list.append(s)
            else:
                print(f"  WARNING: no bowling data found for {name}")

    print(f"\nGenerating pack: {len(bat_stats_list)} batter card(s), "
          f"{len(bowl_stats_list)} bowler card(s)…")

    html_content = generate_html(bat_stats_list, bowl_stats_list, meta)

    # Default output goes into the input folder when using -d
    out = args.output
    if args.folder and out == 'pre_game_pack.html':
        out = str(Path(args.folder) / 'pre_game_pack.html')

    with open(out, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"✓ Saved to: {os.path.abspath(out)}")
    print("  Open in a browser to edit and add images, then Print / Export PDF.")


if __name__ == '__main__':
    main()
