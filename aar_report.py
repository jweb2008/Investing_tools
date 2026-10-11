"""
After Action Review dashboard: aar_report.html, rebuilt by aar.py on every run.

One self-contained page (no internet needed, opens in any browser). Charts are
drawn as inline SVG from the same tables that go into aar.xlsx, so the page
and the workbook always agree. Every chart has a hover readout and a table
view; light and dark mode follow the computer's setting.
"""
import datetime as dt
import html
import math

import numpy as np
import pandas as pd

# Palette: validated categorical slots 1-3, fixed status colors, blue/red
# diverging pair (see the dataviz reference palette).
CSS = """
:root{
  --page:#f9f9f7;--surface:#fcfcfb;--ink:#0b0b0b;--ink2:#52514e;--muted:#898781;
  --grid:#e1e0d9;--axis:#c3c2b7;--ring:rgba(11,11,11,.10);
  --s1:#2a78d6;--s2:#eb6834;--s3:#1baf7a;
  --good:#0ca30c;--neutral:#b9b7ae;--bad:#d03b3b;--warn:#fab219;
  --pos:#2a78d6;--neg:#e34948;--goodtext:#006300;--badtext:#b42f2f;
  color-scheme:light;
}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme="light"])){
  --page:#0d0d0d;--surface:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--muted:#898781;
  --grid:#2c2c2a;--axis:#383835;--ring:rgba(255,255,255,.10);
  --s1:#3987e5;--s2:#d95926;--s3:#199e70;
  --neutral:#5c5b56;--pos:#3987e5;--neg:#e66767;--goodtext:#0ca30c;--badtext:#e66767;
  color-scheme:dark;}}
:root[data-theme="dark"]{
  --page:#0d0d0d;--surface:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--muted:#898781;
  --grid:#2c2c2a;--axis:#383835;--ring:rgba(255,255,255,.10);
  --s1:#3987e5;--s2:#d95926;--s3:#199e70;
  --neutral:#5c5b56;--pos:#3987e5;--neg:#e66767;--goodtext:#0ca30c;--badtext:#e66767;
  color-scheme:dark;}
*{box-sizing:border-box}
body{margin:0;background:var(--page);color:var(--ink);
  font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:24px 16px 48px}
header{display:flex;flex-wrap:wrap;align-items:baseline;justify-content:space-between;gap:8px;margin-bottom:20px}
h1{font-size:24px;margin:0;font-weight:650;letter-spacing:-.01em}
.meta{color:var(--ink2);font-size:13px}
h2{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--ink2);
  margin:32px 0 10px;font-weight:600}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
.tile,.card{background:var(--surface);border-radius:12px;box-shadow:0 0 0 1px var(--ring)}
.tile{padding:14px 16px}
.tile .lab{color:var(--ink2);font-size:13px}
.tile .val{font-size:28px;font-weight:620;margin-top:2px}
.tile .sub{color:var(--muted);font-size:12px;margin-top:2px}
.tile.hero .val{font-size:clamp(36px,9vw,48px);line-height:1.05;white-space:nowrap}
@media (max-width:620px){.tile.hero{grid-column:1/-1}}
.up{color:var(--goodtext)}.down{color:var(--badtext)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,520px),1fr));gap:16px}
.card{padding:16px 16px 10px;min-width:0}
.card h3{margin:0;font-size:15px;font-weight:620}
.card .why{color:var(--ink2);font-size:13px;margin:2px 0 8px}
.legend{display:flex;flex-wrap:wrap;gap:4px 14px;color:var(--ink2);font-size:12px;margin:2px 0 6px}
.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:-1px}
.legend i.ln{height:2px;width:14px;border-radius:1px;vertical-align:3px}
svg{display:block;width:100%;height:auto;overflow:visible}
svg text{fill:var(--muted);font-size:11px;font-family:inherit;pointer-events:none}
svg text.lab{fill:var(--ink2)}
svg text.val{fill:var(--ink);font-weight:600}
.mark{transition:opacity .12s}
.mark:hover,.mark:focus{opacity:.78;outline:none}
.empty{color:var(--muted);border:1px dashed var(--axis);border-radius:10px;padding:26px 16px;
  text-align:center;margin:8px 0 6px;font-size:13px}
details{margin-top:6px}
summary{cursor:pointer;color:var(--muted);font-size:12px;list-style:none}
summary::before{content:"\\25B8  "}details[open] summary::before{content:"\\25BE  "}
table{border-collapse:collapse;width:100%;font-size:12px;margin-top:6px;font-variant-numeric:tabular-nums}
th,td{text-align:right;padding:3px 6px;border-bottom:1px solid var(--grid)}
th:first-child,td:first-child{text-align:left}
th{color:var(--ink2);font-weight:600}
.tablewrap{overflow-x:auto}
#tip{position:fixed;pointer-events:none;background:var(--surface);color:var(--ink);
  box-shadow:0 0 0 1px var(--ring),0 6px 18px rgba(0,0,0,.14);border-radius:8px;
  padding:7px 10px;font-size:12px;display:none;z-index:9;max-width:260px}
#tip b{font-size:14px;display:block}
#tip span{color:var(--ink2)}
.note{color:var(--muted);font-size:12px;margin-top:28px}
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chip{background:var(--surface);box-shadow:0 0 0 1px var(--ring);border-radius:999px;padding:4px 12px;font-size:13px}
.chip b{margin-left:4px}
"""

JS = """
const tip=document.getElementById('tip');
function show(e,el){const v=el.getAttribute('data-v'),l=el.getAttribute('data-l');
  tip.replaceChildren();const b=document.createElement('b');b.textContent=v;tip.appendChild(b);
  const s=document.createElement('span');s.textContent=l;tip.appendChild(s);tip.style.display='block';
  let r=el.getBoundingClientRect(),x=(e&&e.clientX)||r.left+r.width/2,y=(e&&e.clientY)||r.top;
  const w=tip.offsetWidth;tip.style.left=Math.min(x+14,innerWidth-w-8)+'px';tip.style.top=(y+14)+'px';}
document.querySelectorAll('[data-v]').forEach(el=>{
  el.addEventListener('pointermove',e=>show(e,el));el.addEventListener('focus',()=>show(null,el));
  el.addEventListener('pointerleave',()=>tip.style.display='none');el.addEventListener('blur',()=>tip.style.display='none');});
"""

OUTCOME_COLORS = [("Win", "var(--good)"), ("Scratch", "var(--neutral)"), ("Loss", "var(--bad)")]
RESULT_COLORS = [("Target first", "var(--good)"), ("Neither", "var(--neutral)"),
                 ("Stop first", "var(--bad)"), ("Not triggered", "url(#hatch)")]
HATCH_DEF = ('<defs><pattern id="hatch" width="6" height="6" patternUnits="userSpaceOnUse" '
             'patternTransform="rotate(45)"><rect width="6" height="6" style="fill:var(--surface)"/>'
             '<line x1="0" y1="0" x2="0" y2="6" style="stroke:var(--neutral);stroke-width:2.5"/>'
             '</pattern></defs>')
STYLE_COLORS = {"Swing": "var(--s1)", "Hedge": "var(--s2)", "Other": "var(--s3)"}
DTE_COLORS = {"0-14 days": "var(--s2)", "15-45 days": "var(--s1)", "46+ days": "var(--s3)"}

W = 640  # SVG coordinate width; scales to the card


# --------------------------------------------------------------- formatting
def esc(s):
    return html.escape(str(s), quote=True)


def money(v, sign=True):
    if v is None or pd.isna(v):
        return "n/a"
    s = "-" if v < 0 else ("+" if sign and v > 0 else "")
    a = abs(v)
    if a >= 1e6:
        return f"{s}${a / 1e6:.1f}M"
    if a >= 1e4:
        return f"{s}${a / 1e3:.1f}K"
    return f"{s}${a:,.0f}"


def pct(v, sign=True, d=1):
    if v is None or pd.isna(v):
        return "n/a"
    return f"{v:+.{d}f}%" if sign else f"{v:.{d}f}%"


def nice_ticks(lo, hi, n=5):
    if lo == hi:
        lo, hi = lo - 1, hi + 1
    span = hi - lo
    step = 10 ** math.floor(math.log10(span / n))
    for m in (1, 2, 2.5, 5, 10):
        if span / (step * m) <= n:
            step *= m
            break
    start = math.floor(lo / step) * step
    ticks, t = [], start
    while t <= hi + step * 0.001:
        ticks.append(round(t, 10))
        t += step
    if ticks[-1] < hi:
        ticks.append(ticks[-1] + step)
    return ticks


def compact(v, unit=""):
    if unit == "$":
        return money(v, sign=False) if v >= 0 else money(v)
    if unit == "%":
        return f"{v:g}%"
    a = abs(v)
    return f"{v / 1e3:g}K" if a >= 1000 else f"{v:g}"


def mark(shape, value, label, **attrs):
    """An SVG element with a hover/focus readout (value leads, label follows)."""
    a = " ".join(f'{k.replace("_", "-")}="{esc(v)}"' for k, v in attrs.items())
    return (f'<{shape} class="mark" tabindex="0" data-v="{esc(value)}" data-l="{esc(label)}" '
            f'{a}/>')


def table(df):
    if df is None or len(df) == 0:
        return ""
    head = "".join(f"<th>{esc(c)}</th>" for c in df.columns)
    body = "".join("<tr>" + "".join(f"<td>{esc('' if pd.isna(v) else v)}</td>" for v in r) + "</tr>"
                   for r in df.itertuples(index=False))
    return (f'<details><summary>Table view</summary><div class="tablewrap"><table><tr>{head}</tr>'
            f"{body}</table></div></details>")


def legend(items, line=False):
    def sw(c):
        if c.startswith("url("):   # hatched swatch, same 45 degree lines as the mark
            return ("background:repeating-linear-gradient(45deg,var(--neutral) 0 2px,"
                    "var(--surface) 2px 4px);box-shadow:0 0 0 1px var(--neutral) inset")
        return f"background:{c}"
    return '<div class="legend">' + "".join(
        f'<span><i class="{"ln" if line else ""}" style="{sw(c)}"></i>{esc(n)}</span>'
        for n, c in items) + "</div>"


def card(title, why, body, tbl=""):
    return (f'<div class="card"><h3>{esc(title)}</h3><div class="why">{esc(why)}</div>'
            f"{body}{tbl}</div>")


def empty(msg):
    return f'<div class="empty">{esc(msg)}</div>'


# ------------------------------------------------------------------- charts
def stacked_100(rows, segs, label_w=150):
    """rows: [(label, {seg: count}, n)]. Horizontal 100% bars, 2px surface gaps."""
    bar_h, gap = 22, 14
    h = len(rows) * (bar_h + gap) + 6
    x0, x1 = label_w, W - 44
    out = [f'<svg viewBox="0 0 {W} {h}" role="img">' + (HATCH_DEF if any(
        c.startswith("url(") for _, c in segs) else "")]
    for i, (lab, counts, n) in enumerate(rows):
        y = i * (bar_h + gap) + 4
        out.append(f'<text class="lab" x="0" y="{y + bar_h / 2 + 4}">{esc(lab)}</text>')
        out.append(f'<text x="{W}" y="{y + bar_h / 2 + 4}" text-anchor="end">n={n}</text>')
        x = x0
        total = sum(counts.values()) or 1
        visible = [(s, c) for s, c in segs if counts.get(s, 0) > 0]
        for j, (seg, color) in enumerate(visible):
            w = (x1 - x0) * counts[seg] / total
            wd = max(w - (2 if j < len(visible) - 1 else 0), 1)
            rx = 4 if j == len(visible) - 1 else 0
            share = counts[seg] / total * 100
            out.append(_rect(x, y, wd, bar_h, color, rx_right=rx, rx_left=4 if j == 0 else 0,
                             value=f"{share:.0f}%  ({counts[seg]})", label=f"{lab}: {seg}"))
            if wd > 34:   # label inside only when it fits; ink on light fills, white on strong ones
                ink = "var(--ink)" if color in ("var(--neutral)", "var(--axis)", "url(#hatch)") else "#fff"
                out.append(f'<text x="{x + wd / 2}" y="{y + bar_h / 2 + 4}" text-anchor="middle" '
                           f'style="fill:{ink};font-weight:600">{share:.0f}%</text>')
            x += w
    out.append("</svg>")
    return "".join(out)


def _rect(x, y, w, h, color, rx_left=0, rx_right=0, value="", label=""):
    """Rect with optional rounding on the left or right end only."""
    r = min(rx_left or rx_right, h / 2, w / 2)
    if not (rx_left or rx_right) or r <= 0:
        return mark("rect", value, label, x=f"{x:.1f}", y=f"{y:.1f}", width=f"{w:.1f}",
                    height=f"{h:.1f}", style=f"fill:{color}")
    rl, rr = (r if rx_left else 0), (r if rx_right else 0)
    d = (f"M{x + rl:.1f},{y:.1f} H{x + w - rr:.1f} "
         + (f"A{rr},{rr} 0 0 1 {x + w:.1f},{y + rr:.1f} V{y + h - rr:.1f} A{rr},{rr} 0 0 1 {x + w - rr:.1f},{y + h:.1f} "
            if rr else f"H{x + w:.1f} V{y + h:.1f} ")
         + f"H{x + rl:.1f} "
         + (f"A{rl},{rl} 0 0 1 {x:.1f},{y + h - rl:.1f} V{y + rl:.1f} A{rl},{rl} 0 0 1 {x + rl:.1f},{y:.1f} Z"
            if rl else f"H{x:.1f} V{y:.1f} Z"))
    return mark("path", value, label, d=d, style=f"fill:{color}")


def columns_signed(items, unit="$", h=230):
    """items: [(label, value, tip_label)]. Columns from a zero baseline, blue up / red down,
    4px rounded data end, value on the cap."""
    vals = [v for _, v, _ in items]
    lo, hi = min(0, min(vals)), max(0, max(vals))
    ticks = nice_ticks(lo, hi, 4)
    lo, hi = ticks[0], ticks[-1]
    top, bot, left, right = 16, h - 26, 56, W - 8
    ys = lambda v: bot - (v - lo) / (hi - lo) * (bot - top)
    out = [f'<svg viewBox="0 0 {W} {h}" role="img">']
    for t in ticks:
        out.append(f'<line x1="{left}" x2="{right}" y1="{ys(t):.1f}" y2="{ys(t):.1f}" '
                   f'style="stroke:var({"--axis" if t == 0 else "--grid"})"/>')
        out.append(f'<text x="{left - 6}" y="{ys(t) + 4:.1f}" text-anchor="end">{esc(compact(t, unit))}</text>')
    band = (right - left) / len(items)
    bw = min(48, band * 0.55)
    for i, (lab, v, tl) in enumerate(items):
        cx = left + band * (i + 0.5)
        y0, y1 = ys(0), ys(v)
        color = "var(--pos)" if v >= 0 else "var(--neg)"
        top_y, height = min(y0, y1), max(abs(y1 - y0), 1)
        txt = money(v) if unit == "$" else pct(v)
        out.append(_vrect(cx - bw / 2, top_y, bw, height, color, up=v >= 0, value=txt, label=tl))
        ly = y1 - 6 if v >= 0 else y1 + 14
        out.append(f'<text class="val" x="{cx:.1f}" y="{ly:.1f}" text-anchor="middle">{esc(txt)}</text>')
        out.append(f'<text class="lab" x="{cx:.1f}" y="{h - 8}" text-anchor="middle">{esc(lab)}</text>')
    out.append("</svg>")
    return "".join(out)


def _vrect(x, y, w, h, color, up=True, value="", label=""):
    r = min(4, w / 2, h)
    if up:
        d = (f"M{x:.1f},{y + h:.1f} V{y + r:.1f} A{r},{r} 0 0 1 {x + r:.1f},{y:.1f} "
             f"H{x + w - r:.1f} A{r},{r} 0 0 1 {x + w:.1f},{y + r:.1f} V{y + h:.1f} Z")
    else:
        d = (f"M{x:.1f},{y:.1f} V{y + h - r:.1f} A{r},{r} 0 0 0 {x + r:.1f},{y + h:.1f} "
             f"H{x + w - r:.1f} A{r},{r} 0 0 0 {x + w:.1f},{y + h - r:.1f} V{y:.1f} Z")
    return mark("path", value, label, d=d, style=f"fill:{color}")


def histogram(values, edges, h=200, unit="%"):
    counts, _ = np.histogram(np.clip(values, edges[0], edges[-1]), bins=edges)
    top, bot, left, right = 12, h - 26, 34, W - 8
    hi = max(nice_ticks(0, max(counts.max(), 1), 4)[-1], 1)
    ys = lambda v: bot - v / hi * (bot - top)
    out = [f'<svg viewBox="0 0 {W} {h}" role="img">']
    for t in nice_ticks(0, hi, 4):
        out.append(f'<line x1="{left}" x2="{right}" y1="{ys(t):.1f}" y2="{ys(t):.1f}" '
                   f'style="stroke:var({"--axis" if t == 0 else "--grid"})"/>')
        out.append(f'<text x="{left - 6}" y="{ys(t) + 4:.1f}" text-anchor="end">{t:g}</text>')
    n = len(counts)
    band = (right - left) / n
    for i, c in enumerate(counts):
        a, b = edges[i], edges[i + 1]
        x = left + band * i + 1
        color = "var(--neg)" if b <= 0 else "var(--pos)"
        lab = (f"{'≤' if i == 0 else ''}{a:g}% to {b:g}%{'+' if i == n - 1 else ''}")
        if c:
            out.append(_vrect(x, ys(c), band - 2, bot - ys(c), color, value=f"{c} trades", label=lab))
    for i in range(0, n + 1, max(1, n // 6)):
        e = edges[i]
        x = left + band * i
        out.append(f'<text x="{x:.1f}" y="{h - 8}" text-anchor="middle">{e:g}%</text>')
    out.append("</svg>")
    return "".join(out)


def line_chart(series, h=250):
    """series: [(name, color, [(date, value)])] cumulative $; 2px lines, end dots, end labels."""
    alld = [d for _, _, pts in series for d, _ in pts]
    allv = [v for _, _, pts in series for _, v in pts] + [0]
    d0, d1 = min(alld), max(alld)
    ticks = nice_ticks(min(allv), max(allv), 4)
    lo, hi = ticks[0], ticks[-1]
    top, bot, left, right = 14, h - 26, 58, W - 92
    span = max((d1 - d0).days, 1)
    xs = lambda d: left + (d - d0).days / span * (right - left)
    ys = lambda v: bot - (v - lo) / (hi - lo) * (bot - top)
    out = [f'<svg viewBox="0 0 {W} {h}" role="img">']
    for t in ticks:
        out.append(f'<line x1="{left}" x2="{right}" y1="{ys(t):.1f}" y2="{ys(t):.1f}" '
                   f'style="stroke:var({"--axis" if t == 0 else "--grid"})"/>')
        out.append(f'<text x="{left - 6}" y="{ys(t) + 4:.1f}" text-anchor="end">{esc(compact(t, "$"))}</text>')
    months = pd.date_range(d0.replace(day=1), d1, freq="MS")
    step = max(1, len(months) // 7)
    for m in months[::step]:
        if m >= d0:
            out.append(f'<text x="{xs(m):.1f}" y="{h - 8}" text-anchor="middle">{m:%b %y}</text>')
    ends = []
    for name, color, pts in series:
        if not pts:
            continue
        path = "M" + " L".join(f"{xs(d):.1f},{ys(v):.1f}" for d, v in pts)
        out.append(f'<path d="{path}" style="fill:none;stroke:{color};stroke-width:2;'
                   f'stroke-linejoin:round;stroke-linecap:round"/>')
        for d, v in pts:   # invisible hover targets on every step
            out.append(mark("circle", money(v), f"{name} · {d:%b %d, %Y}", cx=f"{xs(d):.1f}",
                            cy=f"{ys(v):.1f}", r="9", style="fill:transparent"))
        d, v = pts[-1]
        out.append(f'<circle cx="{xs(d):.1f}" cy="{ys(v):.1f}" r="4.5" '
                   f'style="fill:{color};stroke:var(--surface);stroke-width:2"/>')
        ends.append([ys(v), name, v])
    ends.sort()
    for k in range(1, len(ends)):       # keep end labels from overlapping
        if ends[k][0] - ends[k - 1][0] < 14:
            ends[k][0] = ends[k - 1][0] + 14
    for y, name, v in ends:
        out.append(f'<text class="val" x="{right + 10}" y="{y + 4:.1f}">{esc(money(v))}</text>')
    out.append("</svg>")
    return "".join(out)


Y_CAP = 300   # option results above this are drawn at the top edge (hover shows the real value)


def scatter(points, h=300):
    """points: [(x, y, color, tip)]. Quadrant chart: underlying move vs option result."""
    xs_v = [p[0] for p in points] + [0]
    ys_v = [min(p[1], Y_CAP) for p in points] + [0]
    xt = nice_ticks(min(xs_v), max(xs_v), 5)
    yt = nice_ticks(max(min(ys_v), -100), max(ys_v), 5)
    top, bot, left, right = 14, h - 30, 46, W - 10
    sx = lambda v: left + (v - xt[0]) / (xt[-1] - xt[0]) * (right - left)
    sy = lambda v: bot - (v - yt[0]) / (yt[-1] - yt[0]) * (bot - top)
    out = [f'<svg viewBox="0 0 {W} {h}" role="img">']
    for t in yt:
        out.append(f'<line x1="{left}" x2="{right}" y1="{sy(t):.1f}" y2="{sy(t):.1f}" '
                   f'style="stroke:var({"--axis" if t == 0 else "--grid"})"/>')
        out.append(f'<text x="{left - 6}" y="{sy(t) + 4:.1f}" text-anchor="end">{t:g}%</text>')
    for t in xt:
        out.append(f'<line x1="{sx(t):.1f}" x2="{sx(t):.1f}" y1="{top}" y2="{bot}" '
                   f'style="stroke:var({"--axis" if t == 0 else "--grid"})"/>')
        out.append(f'<text x="{sx(t):.1f}" y="{bot + 16}" text-anchor="middle">{t:g}%</text>')
    out.append(f'<text x="{right}" y="{h - 2}" text-anchor="end">underlying move while held →</text>')
    q = [(right - 4, top + 12, "end", "Stock up, option won"), (left + 4, top + 12, "start", "Stock down, option won"),
         (right - 4, bot - 6, "end", "Stock up, option LOST"), (left + 4, bot - 6, "start", "Stock down, option lost")]
    for x, y, anc, t in q:
        out.append(f'<text x="{x}" y="{y}" text-anchor="{anc}" style="font-size:10.5px">{esc(t)}</text>')
    for x, y, color, tl in points:
        cx, cy = sx(x), sy(min(max(y, yt[0]), yt[-1]))
        out.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="5" '
                   f'style="fill:{color};stroke:var(--surface);stroke-width:2"/>')
        out.append(mark("circle", f"option {pct(y, d=0)} · stock {pct(x)}", tl, cx=f"{cx:.1f}",
                        cy=f"{cy:.1f}", r="12", style="fill:transparent"))
    out.append("</svg>")
    return "".join(out)


# ------------------------------------------------------------------ sections
def _outcome_rows(df, by):
    rows = []
    for key, g in df.groupby(by, sort=True):
        key = key if isinstance(key, tuple) else (key,)
        rows.append((" · ".join(str(k) for k in key), g["Outcome"].value_counts().to_dict(), len(g)))
    return rows


def build(decisions, trades, flags, review, info, out_path):
    d = decisions.copy()
    d = d[d["NetPct"].notna()]
    swing = d[d["Style"] == "Swing"]
    closed = trades[(trades["Status"] == "Closed") & trades["NetProfit"].notna()].copy()
    run_info = dict(zip(info["Item"], info["Value"]))
    parts = []

    # ---- headline tiles
    net = closed["NetProfit"].sum()
    ytd = closed[pd.to_datetime(closed["ExitDate"]) >= pd.Timestamp(dt.date.today().year, 1, 1)]["NetProfit"].sum()
    sw_win = (swing["Outcome"] == "Win").mean() * 100 if len(swing) else None
    sw_med = swing["NetPct"].median() if len(swing) else None
    n_scan = int((trades["Source"] == "Scanner").sum()) if "Source" in trades else 0
    n_flags = len(flags) if flags is not None else 0
    pending = int((flags["Taken"] == "Window open").sum()) if n_flags else 0
    tiles = [
        ("hero", "Net profit, closed trades", money(ytd), f"this year · {money(net)} all rows",
         "up" if ytd >= 0 else "down"),
        ("", "Swing win rate", pct(sw_win, sign=False, d=0) if sw_win is not None else "n/a",
         f"{len(swing)} swing decisions", ""),
        ("", "Swing median return", pct(sw_med), "half of swing trades did better",
         "up" if (sw_med or 0) >= 0 else "down"),
        ("", "Scanner trades", f"{n_scan}", "entered 1-5 days after a flag", ""),
        ("", "Setups logged", f"{n_flags}", f"{pending} still in their 5-day window", ""),
    ]
    parts.append('<div class="tiles">' + "".join(
        f'<div class="tile {k}"><div class="lab">{esc(l)}</div><div class="val {c}">{esc(v)}</div>'
        f'<div class="sub">{esc(s)}</div></div>' for k, l, v, s, c in tiles) + "</div>")

    # ---- results
    parts.append("<h2>How trades turned out</h2><div class=\"grid\">")
    rows = _outcome_rows(d, ["Style", "Asset"])
    tbl = d.groupby(["Style", "Asset"]).agg(N=("NetPct", "size"),
                                            Win=("Outcome", lambda s: f"{(s == 'Win').mean() * 100:.0f}%"),
                                            Scratch=("Outcome", lambda s: f"{(s == 'Scratch').mean() * 100:.0f}%"),
                                            Loss=("Outcome", lambda s: f"{(s == 'Loss').mean() * 100:.0f}%"),
                                            Median=("NetPct", lambda s: pct(s.median()))).reset_index()
    parts.append(card("Win, scratch, loss by style",
                      "Closed decisions. Scratch = within ±2% for stocks, ±10% for options.",
                      legend(OUTCOME_COLORS) + stacked_100(rows, OUTCOME_COLORS), table(tbl)))

    # equity curve by style
    closed["Bucket"] = closed["Style"].where(closed["Style"].isin(["Swing", "Hedge"]), "Other")
    closed["ExitDate"] = pd.to_datetime(closed["ExitDate"])
    series = []
    for name in ("Swing", "Hedge", "Other"):
        g = closed[closed["Bucket"] == name].groupby("ExitDate")["NetProfit"].sum().sort_index().cumsum()
        if len(g):
            series.append((name, STYLE_COLORS[name], list(zip(g.index, g.values))))
    if series:
        tb = (closed.groupby([closed["ExitDate"].dt.to_period("M").astype(str), "Bucket"])["NetProfit"]
              .sum().unstack(fill_value=0).round(0).reset_index().rename(columns={"ExitDate": "Month"}))
        parts.append(card("Cumulative net profit by style",
                          "Running total of recorded Net Profit by exit date. Other = Day and Long term.",
                          legend([(n, c) for n, c, _ in series], line=True) + line_chart(series), table(tb)))
    parts.append("</div>")

    # ---- swing returns
    parts.append("<h2>Swing trades up close</h2><div class=\"grid\">")
    for asset, edges in (("Stock", [-20, -10, -5, -2, 0, 2, 5, 10, 20, 40]),
                         ("Option", [-100, -75, -50, -25, -10, 0, 10, 25, 50, 100, 200])):
        v = swing[swing["Asset"] == asset]["NetPct"].astype(float).values
        body = histogram(v, edges) if len(v) else empty("No closed swing trades yet.")
        tb = pd.DataFrame({"Range": [f"{edges[i]:g}% to {edges[i + 1]:g}%" for i in range(len(edges) - 1)],
                           "Trades": np.histogram(np.clip(v, edges[0], edges[-1]), bins=edges)[0]}) if len(v) else None
        parts.append(card(f"Swing {asset.lower()} returns",
                          f"How {len(v)} closed swing {asset.lower()} decisions were spread. "
                          "Red = losses, blue = gains; the ends collect everything beyond them.",
                          legend([("Loss", "var(--neg)"), ("Gain", "var(--pos)")]) + body, table(tb)))

    so = swing[swing["Asset"] == "Option"]
    if "DTEBucket" in so and so["DTEBucket"].notna().any():
        g = so.groupby("DTEBucket").agg(N=("NetPct", "size"), Net=("Net", "sum"),
                                         Win=("Outcome", lambda s: (s == "Win").mean() * 100),
                                         Median=("NetPct", "median"))
        order = [b for b in ("0-14 days", "15-45 days", "46+ days") if b in g.index]
        items = [(f"{b} (n={int(g.loc[b, 'N'])})", float(g.loc[b, "Net"]),
                  f"{b} to expiry · win {g.loc[b, 'Win']:.0f}% · median {pct(g.loc[b, 'Median'])}")
                 for b in order]
        tb = g.loc[order].reset_index().assign(Net=lambda x: x["Net"].map(money),
                                               Win=lambda x: x["Win"].map(lambda v: f"{v:.0f}%"),
                                               Median=lambda x: x["Median"].map(pct))
        parts.append(card("Swing options: net profit by days to expiry at entry",
                          "Short-dated contracts need the move to come fast.",
                          columns_signed(items), table(tb)))
    pts = so[so.get("UnderMovePct", pd.Series(dtype=float)).notna()] if "UnderMovePct" in so else so.iloc[0:0]
    if len(pts):
        points = [(float(r.UnderMovePct), float(r.NetPct), DTE_COLORS.get(r.DTEBucket, "var(--s1)"),
                   f"{r.Symbol} · {pd.Timestamp(r.EntryDate):%b %d} · {r.DTEBucket or ''}")
                  for r in pts.itertuples()]
        tb = pts[["Symbol", "EntryDate", "DTEBucket", "UnderMovePct", "NetPct"]].copy()
        tb["EntryDate"] = pd.to_datetime(tb["EntryDate"]).dt.date
        parts.append(card("Swing options: stock move vs option result",
                          "Each dot is a trade. Bottom right = right on direction, still lost "
                          f"(time or strike). Results above +{Y_CAP}% sit on the top edge.",
                          legend([(k, c) for k, c in DTE_COLORS.items()]) + scatter(points), table(tb)))
    else:
        parts.append(card("Swing options: stock move vs option result",
                          "Each dot will be a trade: did the stock move your way, and did the option pay?",
                          empty("Needs price history. Run py aar.py without --no-prices.")))
    parts.append("</div>")

    # ---- scanner
    parts.append("<h2>Scanner</h2><div class=\"grid\">")
    live = swing[swing.get("Source", pd.Series(index=swing.index, dtype=object)).isin(["Scanner", "Discretionary"])] \
        if "Source" in swing else swing.iloc[0:0]
    if len(live):
        rows = _outcome_rows(live, ["Source", "Asset"])
        parts.append(card("Scanner vs discretionary swing trades",
                          "Only trades entered since the scan log began.",
                          legend(OUTCOME_COLORS) + stacked_100(rows, OUTCOME_COLORS)))
    else:
        parts.append(card("Scanner vs discretionary swing trades",
                          "Trades entered 1-5 trading days after a flag vs everything else.",
                          empty("No swing trades since the scan log began yet. This fills in as you trade.")))
    if n_flags:
        fl = flags.copy()
        fl["FirstFlagDate"] = pd.to_datetime(fl["FirstFlagDate"])
        cnt = fl.groupby([fl["FirstFlagDate"].dt.date, "BestStatus"]).size().unstack(fill_value=0)
        stat_cols = [("Setup", "var(--s1)"), ("Near miss", "var(--s3)"), ("Breakout", "var(--s2)")]
        rows = [(f"{d:%a %b %d}", {k: int(cnt.loc[d].get(k, 0)) for k, _ in stat_cols},
                 int(cnt.loc[d].sum())) for d in list(cnt.index)[-12:]]
        parts.append(card("Setups found per scan day",
                          "New setup episodes by first-flag day (last 12 scan days), by status.",
                          legend(stat_cols) + stacked_100(rows, stat_cols, label_w=110),
                          table(cnt.reset_index().rename(columns={"FirstFlagDate": "Day"}))))
        res = fl[fl.get("OutcomeStatus", pd.Series(index=fl.index, dtype=object))
                 .isin(["complete", "resolved, 20d return pending"]) | (fl.get("Result") == "Not triggered")] \
            if "Result" in fl else fl.iloc[0:0]
        if len(res):
            res = res.assign(R=res["Result"].fillna("").map(
                lambda r: "Neither" if str(r).startswith("Neither") else r))
            for by, title in (("BestStatus", "status"), ("Combo", "check combination")):
                rows = [(k, g["R"].value_counts().to_dict(), len(g)) for k, g in res.groupby(by)]
                parts.append(card(f"How setups played out, by {title}",
                                  "Every setup on its own, taken or not: did it trigger, then hit the "
                                  "target or the stop first (20 sessions)?",
                                  legend(RESULT_COLORS) + stacked_100(rows, RESULT_COLORS, label_w=170)))
        else:
            parts.append(card("How setups played out",
                              "Every setup on its own: triggered, then target or stop first.",
                              empty(f"{n_flags} setups logged; outcomes need about a week of price history "
                                    "after each flag. Check back after a few trading days.")))
    parts.append("</div>")

    # ---- housekeeping
    if review is not None and len(review):
        vc = review["Issue"].value_counts()
        parts.append("<h2>Review sheet</h2><div class=\"chips\">" + "".join(
            f'<span class="chip">{esc(k)}<b>{v}</b></span>' for k, v in vc.items()) + "</div>")

    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>After Action Review</title><style>{CSS}</style></head><body><div class="wrap">
<header><h1>After Action Review</h1><div class="meta">{esc(run_info.get('Tracker', ''))} ·
run {esc(run_info.get('Run', ''))} · prices through {esc(run_info.get('Bars complete through', 'n/a'))}</div></header>
{''.join(parts)}
<div class="note">Built by aar.py from the tracker (net % as recorded, never recalculated). Decisions combine
same-day entries across accounts. Groups under 20 are too few to read; tune nothing until about 50.
Details and every number: aar.xlsx.</div></div><div id="tip"></div><script>{JS}</script></body></html>"""
    out_path.write_text(page, encoding="utf-8")
    return out_path
