#!/usr/bin/env python3
"""
IM8 Winners — XLS processor + Meta thumbnail fetcher
Reads the XLS from the repo, fetches thumbnails via Meta API, generates index.html
"""

import os, json, re, glob
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.parse import urlencode
from urllib.error import URLError

META_TOKEN  = os.environ.get("META_ACCESS_TOKEN", "")
AD_ACCOUNT  = "act_1000723654649396"
API_VER     = "v20.0"
BASE        = f"https://graph.facebook.com/{API_VER}"

WINNER_POOL_KW = ['Winner','Winners','WINNER','TOP30','l7d winner','TOP 50']
ICP_KW = ['ICP','GLP','Menopause','Collagen','ANGLE','ACTIVE SENIOR','Senior',
          'Cognitive','Immune','Fitness','Sleep','Weight','Gut','Joint','Pill',
          'Energy','Green','Young Prof','Persona','NERMW','HCSS','RECOVERY',
          'Aging Athlete','Performance','FREQUENTFLYER','Traveler','Travel']
L3_EXCL = ['Retargeting','ENGAGER','ATC','GEISTM']
LP_MAP = {
    'NOBSLDP':'https://get.im8health.com/pages/no-bs',
    'GLP1LDP':'https://get.im8health.com/pages/glp1',
    'FEELAGAINLDP':'https://get.im8health.com/pages/feel-again',
    'BKMFORMULALDP':'https://get.im8health.com/pages/beckham-formula',
    'WHYIM8LDP':'https://get.im8health.com/pages/why-im8',
    'SENIORSLDP':'https://get.im8health.com/pages/seniors',
    'MENOPAUSELDP':'https://get.im8health.com/pages/menopause',
    '16IN1DRJAMES':'https://get.im8health.com/pages/dr-james',
    'GETGUTLDP':'https://get.im8health.com/for/gut',
    'GETRECOVERYACTLDP':'https://get.im8health.com/recovery/active',
    'GETJOINTSLDP':'https://get.im8health.com/for/joints',
    'GETTRAVELLDP':'https://get.im8health.com/for/travel',
    'SCIENCELDP':'https://get.im8health.com/pages/science',
    'PROOFLDP':'https://get.im8health.com/pages/proof',
    'ACTNOWLDP':'https://get.im8health.com/pages/act-now',
    'GETPDP':'https://get.im8health.com/essentials',
    'HOMEPAGE':'https://im8health.com/',
    'PDP':'https://im8health.com/products/essentials',
    'PROMPTLDP':'https://get.im8health.com/prompt',
    'PROUPGRADELDP':'https://get.im8health.com/pages/pro-upgrade',
}

# ── HELPERS ───────────────────────────────────────────
def get_tier(c):
    s = str(c)
    for t in ['XX','L1','L2','L3']:
        if s.startswith(t): return t
    return 'OTHER'

def has_kw(s, kws):
    s = str(s).lower()
    return any(k.lower() in s for k in kws)

def is_tagged(n): return bool(re.search(r'WIN2\d', str(n), re.I))

def ad_type(n):
    u = str(n).upper()
    if 'KOLUGC' in u or 'KOL_UGC' in u: return 'KOL UGC'
    if 'CREATORUGC' in u: return 'Creator UGC'
    if 'JAMESPOST' in u or 'IGPOST' in u: return 'IG Post'
    if any(x in u for x in ['_VID_','_VSL_','_WOTXT_','_TALKH_','_VLOG_']) or u.startswith('VID_'): return 'Video'
    if '_IMG_' in u or u.startswith('IMG_'): return 'Static'
    return 'Other'

def note_type(ad_name, adset):
    if has_kw(adset, ICP_KW): return 'icp'
    u = str(ad_name).upper()
    if 'KOLUGC' in u or 'KOL_UGC' in u or 'CREATORUGC' in u: return 'kol'
    return 'generic'

def get_lp(n):
    toks = [t.strip().strip('*') for t in str(n).split('_')]
    for t in reversed(toks):
        if t in LP_MAP: return LP_MAP[t]
    return ''

# ── PARSE XLS ─────────────────────────────────────────
def parse_xls(path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return []
        headers = [str(h or '').strip() for h in rows[0]]
        data = [dict(zip(headers, row)) for row in rows[1:]]
        wb.close()
        return data
    except Exception as e:
        print(f"openpyxl failed: {e}, trying xlrd")
        import xlrd
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
        return [dict(zip(headers, [ws.cell_value(r, c) for c in range(ws.ncols)])) for r in range(1, ws.nrows)]

def col(headers, pat):
    for h in headers:
        if re.search(pat, h, re.I): return h
    return ''

def process_xls(path):
    rows = parse_xls(path)
    if not rows: return []

    headers = list(rows[0].keys())
    roas_col  = col(headers, r'roas')
    purch_col = col(headers, r'^purchases$|^purchases$')
    spend_col = col(headers, r'amount spent|spend')
    rev_col   = col(headers, r'conversion value')
    cpa_col   = col(headers, r'cost per purchase')
    ctr_col   = col(headers, r'ctr|click.through')
    hook_col  = col(headers, r'hook')
    hold_col  = col(headers, r'hold')
    url_col   = col(headers, r'website url')
    ad_id_col = col(headers, r'ad id|^id$')
    ad_name_col = col(headers, r'ad name')
    camp_col  = col(headers, r'campaign name')
    adset_col = col(headers, r'ad set name')

    ads = []
    for row in rows:
        ad_name = str(row.get(ad_name_col, '') or '')
        if not ad_name or ad_name == 'Ad name': continue

        camp   = str(row.get(camp_col, '') or '')
        adset  = str(row.get(adset_col, '') or '')
        tier   = get_tier(camp)
        tagged = is_tagged(ad_name)

        try: roas  = float(row.get(roas_col, 0) or 0)
        except: roas = 0
        try: purch = float(row.get(purch_col, 0) or 0)
        except: purch = 0
        try: spend = float(row.get(spend_col, 0) or 0)
        except: spend = 0
        try: rev   = float(row.get(rev_col, 0) or 0)
        except: rev = 0
        try: cpa   = float(row.get(cpa_col, 0) or 0)
        except: cpa = 0
        try:
            ctr_raw = float(row.get(ctr_col, 0) or 0)
            ctr = round(ctr_raw * 100, 2) if 0 < ctr_raw < 1 else round(ctr_raw, 2)
        except: ctr = 0
        try:
            h = row.get(hook_col)
            hook = round(float(h) * 100, 1) if h and float(h) > 0 else None
        except: hook = None
        try:
            h = row.get(hold_col)
            hold = round(float(h) * 100, 1) if h and float(h) > 0 else None
        except: hold = None

        url   = str(row.get(url_col, '') or '')
        ad_id = str(row.get(ad_id_col, '') or '')

        if tier in ('XX', 'OTHER'): continue
        if tier == 'L3' and has_kw(camp, L3_EXCL): continue
        if not tagged and tier == 'L1' and has_kw(adset, WINNER_POOL_KW): continue
        if not tagged and roas <= 1.0: continue
        if not tagged and purch <= 10: continue
        if spend <= 0: continue

        nt = note_type(ad_name, adset)
        lp = get_lp(ad_name) or url

        ads.append({
            'adId': ad_id, 'adName': ad_name, 'camp': camp, 'adset': adset,
            'tier': tier, 'tagged': tagged, 'roas': round(roas, 2),
            'purch': int(purch), 'spend': round(spend, 2), 'rev': round(rev, 2),
            'cpa': round(cpa, 2), 'ctr': ctr, 'hook': hook, 'hold': hold,
            'lp': lp, 'type': ad_type(ad_name), 'nt': nt,
            'thumbnail': '', 'fbLink': ''
        })

    ads.sort(key=lambda x: (x['tagged'], {'icp':0,'kol':1,'generic':2}.get(x['nt'],2), -x['roas']))
    print(f"  Parsed {len(ads)} qualifying ads")
    return ads

# ── META API ──────────────────────────────────────────
def api_get(path, params):
    params['access_token'] = META_TOKEN
    url = f"{BASE}/{path}?{urlencode(params)}"
    try:
        req = Request(url, headers={'User-Agent': 'IM8WinnersBot/1.0'})
        with urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"  API error: {e}")
        return None

def fetch_creatives(ads):
    if not META_TOKEN:
        print("  No token — skipping creatives")
        return
    ids = list(set(a['adId'] for a in ads if a['adId']))
    if not ids:
        print("  No ad IDs found")
        return
    print(f"  Fetching creatives for {len(ids)} ads...")
    for i in range(0, len(ids), 50):
        batch = ids[i:i+50]
        data = api_get(f"{AD_ACCOUNT}/ads", {
            'fields': 'id,creative{thumbnail_url,effective_object_story_id}',
            'filtering': json.dumps([{'field':'id','operator':'IN','value':batch}]),
            'limit': 50,
        })
        if not data or 'data' not in data: continue
        for ad in data['data']:
            cr = ad.get('creative', {})
            thumb = cr.get('thumbnail_url', '')
            story_id = cr.get('effective_object_story_id', '')
            fb_link = ''
            if story_id:
                parts = story_id.split('_', 1)
                if len(parts) == 2:
                    fb_link = f"https://www.facebook.com/permalink.php?story_fbid={parts[1]}&id={parts[0]}"
            for a in ads:
                if a['adId'] == ad['id']:
                    a['thumbnail'] = thumb
                    a['fbLink'] = fb_link
    print(f"  Creatives done")

# ── GENERATE HTML ─────────────────────────────────────
def generate_html(ads, xls_filename):
    now_str = datetime.now().strftime('%-d %b %Y, %H:%M UTC')
    untagged = [a for a in ads if not a['tagged']]
    tagged   = [a for a in ads if a['tagged']]
    ads_json = json.dumps(ads)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>IM8 Winners</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
:root{{--bg:#0a0a0f;--surface:#111118;--surface2:#18181f;--border:rgba(255,255,255,.07);--gold:#e8b450;--gold-dim:rgba(232,180,80,.12);--teal:#3ecfb2;--teal-dim:rgba(62,207,178,.1);--purple:#9b7cff;--purple-dim:rgba(155,124,255,.1);--green:#4ade80;--amber:#fb923c;--blue:#60a5fa;--white:#f0f0f8;--muted:#5a5a7a;--l1:#fb923c;--l3:#60a5fa;}}
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:var(--bg);color:var(--white);font-family:'DM Sans',sans-serif;font-weight:300;min-height:100vh;}}
body::before{{content:'';position:fixed;inset:0;background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='.04'/%3E%3C/svg%3E");pointer-events:none;z-index:0;opacity:.4;}}
.wrap{{position:relative;z-index:1;max-width:1400px;margin:0 auto;padding:0 32px 80px;}}
.dash-header{{position:sticky;top:0;z-index:100;background:var(--surface);border-bottom:1px solid var(--border);}}
.dash-top{{display:flex;align-items:center;gap:16px;padding:0 32px;height:52px;}}
.dash-brand{{font-family:'Syne',sans-serif;font-size:16px;font-weight:800;color:var(--gold);flex-shrink:0;}}
.dash-file{{font-family:'DM Mono',monospace;font-size:11px;color:var(--muted);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.dash-meta{{font-family:'DM Mono',monospace;font-size:10px;color:var(--muted);display:flex;align-items:center;gap:6px;flex-shrink:0;}}
.live-dot{{width:6px;height:6px;background:var(--green);border-radius:50%;animation:pulse 2s infinite;display:inline-block;}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.3}}}}
.dash-filters{{display:flex;align-items:center;gap:6px;padding:10px 32px;flex-wrap:wrap;border-top:1px solid var(--border);}}
.fl{{font-family:'DM Mono',monospace;font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;white-space:nowrap;}}
.fsep{{width:1px;height:14px;background:var(--border);margin:0 4px;}}
.pill{{font-family:'DM Mono',monospace;font-size:10px;font-weight:500;padding:4px 12px;border-radius:20px;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;transition:all .15s;white-space:nowrap;}}
.pill:hover{{border-color:var(--gold);color:var(--gold);}}
.pill.on{{background:var(--gold-dim);border-color:var(--gold);color:var(--gold);}}
.pill.ap.on{{background:rgba(251,146,60,.15);border-color:var(--amber);color:var(--amber);}}
.pill.tp.on{{background:rgba(74,222,128,.1);border-color:var(--green);color:var(--green);}}
.statsbar{{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;padding:20px 32px;border-bottom:1px solid var(--border);}}
.sc{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;position:relative;overflow:hidden;}}
.sc::after{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--gold),transparent);opacity:.4;}}
.sl{{font-size:9px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:5px;font-family:'DM Mono',monospace;}}
.sv{{font-family:'Syne',sans-serif;font-size:22px;font-weight:700;line-height:1;}}
.cards-wrap{{padding:28px 32px;}}
.section-head{{display:flex;align-items:center;gap:14px;margin-bottom:12px;}}
.section-head h2{{font-family:'Syne',sans-serif;font-size:17px;font-weight:700;letter-spacing:-.02em;}}
.sh-count{{font-family:'DM Mono',monospace;font-size:11px;background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:3px 12px;color:#8888aa;}}
.sh-line{{flex:1;height:1px;background:var(--border);}}
.cards{{display:flex;flex-direction:column;gap:6px;margin-bottom:40px;}}
.card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;transition:border-color .2s,transform .15s;animation:fadeUp .3s ease both;cursor:pointer;}}
.card:hover{{border-color:rgba(232,180,80,.25);transform:translateY(-1px);}}
.card.open{{border-color:rgba(232,180,80,.35);}}
.card.untagged{{border-left:3px solid var(--amber);}}
.card.tagged{{border-left:3px solid var(--green);}}
.card-main{{display:grid;grid-template-columns:110px 1fr auto;align-items:stretch;min-height:64px;}}
.card-left{{display:flex;flex-direction:column;justify-content:center;gap:4px;padding:10px 14px;border-right:1px solid var(--border);flex-shrink:0;}}
.tb{{display:inline-flex;align-items:center;justify-content:center;font-family:'DM Mono',monospace;font-size:10px;font-weight:500;padding:2px 8px;border-radius:3px;width:fit-content;}}
.tb.L1{{background:rgba(251,146,60,.15);color:var(--l1);border:1px solid rgba(251,146,60,.3);}}
.tb.L2{{background:rgba(74,222,128,.1);color:var(--green);border:1px solid rgba(74,222,128,.2);}}
.tb.L3{{background:rgba(96,165,250,.1);color:var(--l3);border:1px solid rgba(96,165,250,.25);}}
.tc{{font-size:10px;color:var(--muted);}}
.ab{{font-size:10px;font-family:'DM Mono',monospace;font-weight:500;padding:2px 7px;border-radius:3px;text-align:center;line-height:1.4;}}
.ab.kol{{background:var(--teal-dim);color:var(--teal);border:1px solid rgba(62,207,178,.2);}}
.ab.icp{{background:var(--purple-dim);color:var(--purple);border:1px solid rgba(155,124,255,.2);}}
.ab.generic{{background:var(--gold-dim);color:var(--gold);}}
.card-body{{display:flex;flex-direction:column;justify-content:center;padding:10px 18px;gap:3px;min-width:0;}}
.an{{font-family:'DM Mono',monospace;font-size:11px;color:var(--white);opacity:.85;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.mr{{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:2px;}}
.im{{font-family:'DM Mono',monospace;font-size:11px;font-weight:500;}}
.im.sp{{color:var(--amber);}}
.im.ro.great{{color:var(--green);font-weight:700;}} .im.ro.good{{color:#a3e635;}} .im.ro.ok{{color:var(--amber);}}
.im.pu{{color:var(--white);opacity:.6;}}
.ch-wrap{{display:flex;align-items:center;padding:0 14px;border-left:1px solid var(--border);}}
.ch{{color:var(--muted);transition:transform .2s;display:flex;}}
.card.open .ch{{transform:rotate(180deg);}}
.card-expand{{display:none;border-top:1px solid var(--border);background:var(--surface2);padding:16px 18px;flex-direction:column;gap:12px;}}
.card.open .card-expand{{display:flex;}}
.expand-row{{display:flex;gap:16px;align-items:flex-start;}}
.thumb-wrap{{flex-shrink:0;width:72px;height:54px;border-radius:6px;overflow:hidden;background:var(--bg);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;}}
.thumb-wrap img{{width:100%;height:100%;object-fit:cover;display:block;}}
.no-img{{font-size:16px;color:var(--muted);}}
.expand-details{{display:flex;gap:16px;flex-wrap:wrap;flex:1;}}
.eg{{display:flex;flex-direction:column;gap:2px;min-width:110px;}}
.el{{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;font-family:'DM Mono',monospace;}}
.ev{{font-size:11px;color:var(--white);}}
.ea{{display:flex;gap:8px;flex-wrap:wrap;}}
.ea-btn{{display:inline-flex;align-items:center;gap:5px;font-family:'DM Mono',monospace;font-size:10px;font-weight:500;padding:4px 12px;border-radius:6px;text-decoration:none;transition:all .15s;border:1px solid var(--border);color:var(--muted);}}
.ea-btn.fb{{color:var(--blue);border-color:rgba(96,165,250,.3);background:rgba(96,165,250,.07);}}
.ea-btn.fb:hover{{background:rgba(96,165,250,.18);}}
.ea-btn.lp{{color:var(--gold);border-color:rgba(232,180,80,.3);background:var(--gold-dim);}}
.ea-btn.lp:hover{{background:rgba(232,180,80,.2);}}
.empty-state{{text-align:center;padding:60px;border:1px dashed var(--border);border-radius:12px;}}
.empty-state h3{{font-family:'Syne',sans-serif;font-size:20px;font-weight:700;margin-bottom:8px;}}
.empty-state p{{font-size:11px;color:var(--muted);font-family:'DM Mono',monospace;}}
@keyframes fadeUp{{from{{opacity:0;transform:translateY(8px)}}to{{opacity:1;transform:translateY(0)}}}}
@media(max-width:768px){{.statsbar{{grid-template-columns:repeat(3,1fr);}} .card-main{{grid-template-columns:90px 1fr auto;}}}}
</style>
</head>
<body>
<div class="dash-header">
  <div class="dash-top">
    <div class="dash-brand">IM8 Winners</div>
    <div class="dash-file">{xls_filename}</div>
    <div class="dash-meta"><span class="live-dot"></span>Updated {now_str}</div>
  </div>
  <div class="dash-filters">
    <span class="fl">Status</span>
    <button class="pill on" data-g="status" data-v="all" onclick="setPill(this,'status')">All</button>
    <button class="pill ap" data-g="status" data-v="untagged" onclick="setPill(this,'status')">🔴 Needs Action</button>
    <button class="pill tp" data-g="status" data-v="tagged" onclick="setPill(this,'status')">✅ Tagged</button>
    <div class="fsep"></div>
    <span class="fl">Format</span>
    <button class="pill on" data-g="type" data-v="all" onclick="setPill(this,'type')">All</button>
    <button class="pill" data-g="type" data-v="KOL UGC" onclick="setPill(this,'type')">KOL UGC</button>
    <button class="pill" data-g="type" data-v="Static" onclick="setPill(this,'type')">Static</button>
    <button class="pill" data-g="type" data-v="Video" onclick="setPill(this,'type')">Video</button>
    <button class="pill" data-g="type" data-v="IG Post" onclick="setPill(this,'type')">IG Post</button>
    <button class="pill" data-g="type" data-v="Creator UGC" onclick="setPill(this,'type')">Creator UGC</button>
    <div class="fsep"></div>
    <span class="fl">Tier</span>
    <button class="pill on" data-g="tier" data-v="all" onclick="setPill(this,'tier')">All</button>
    <button class="pill" data-g="tier" data-v="L1" onclick="setPill(this,'tier')">L1</button>
    <button class="pill" data-g="tier" data-v="L3" onclick="setPill(this,'tier')">L3</button>
  </div>
</div>
<div class="statsbar" id="statsbar"></div>
<div class="cards-wrap" id="cards-wrap"></div>
<script>
const ALL_ADS={ads_json};
let filters={{status:'all',type:'all',tier:'all'}};
const fmtUSD=v=>v>0?'$'+Math.round(v).toLocaleString('en-US'):'—';
const fmtROAS=v=>v>0?v.toFixed(2)+'x':'—';
const fmtNum=v=>v>0?Math.round(v).toLocaleString('en-US'):'—';
const rc=r=>r>=2?'great':r>=1.5?'good':'ok';
function setPill(btn,group){{document.querySelectorAll(`[data-g="${{group}}"]`).forEach(b=>b.classList.remove('on'));btn.classList.add('on');filters[group]=btn.dataset.v;render();}}
function toggleCard(card){{const w=card.classList.contains('open');document.querySelectorAll('.card.open').forEach(c=>c.classList.remove('open'));if(!w)card.classList.add('open');}}
function render(){{
  const f=filters;
  const ads=ALL_ADS.filter(a=>{{
    if(f.status==='untagged'&&a.tagged)return false;
    if(f.status==='tagged'&&!a.tagged)return false;
    if(f.type!=='all'&&a.type!==f.type)return false;
    if(f.tier!=='all'&&a.tier!==f.tier)return false;
    return true;
  }});
  const un=ads.filter(a=>!a.tagged).length,tg=ads.filter(a=>a.tagged).length;
  const avgR=ads.length?(ads.reduce((s,a)=>s+a.roas,0)/ads.length).toFixed(2)+'x':'—';
  const totR=ads.reduce((s,a)=>s+a.rev,0),totP=ads.reduce((s,a)=>s+a.purch,0);
  document.getElementById('statsbar').innerHTML=[
    {{l:'Needs Action',v:un,c:'var(--amber)'}},{{l:'Tagged',v:tg,c:'var(--green)'}},
    {{l:'Avg ROAS',v:avgR,c:'var(--gold)'}},{{l:'Total Purchases',v:fmtNum(totP),c:'var(--white)'}},
    {{l:'Revenue',v:fmtUSD(totR),c:'var(--green)'}},
  ].map(s=>`<div class="sc"><div class="sl">${{s.l}}</div><div class="sv" style="color:${{s.c}}">${{s.v}}</div></div>`).join('');
  const ntL={{icp:'ICP — Tag only',kol:'Dupe → Pool',generic:'Dupe → Pool'}};
  function card(a){{
    const fb=a.fbLink?`<a class="ea-btn fb" href="${{a.fbLink}}" target="_blank" onclick="event.stopPropagation()"><svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg> View Post</a>`:'';
    const lp=a.lp?`<a class="ea-btn lp" href="${{a.lp}}" target="_blank" onclick="event.stopPropagation()">↗ Landing Page</a>`:'';
    const th=a.thumbnail?`<img src="${{a.thumbnail}}" alt="" loading="lazy" onerror="this.parentElement.innerHTML='<div class=\\"no-img\\">📷</div>'">`:`<div class="no-img">📷</div>`;
    return `<div class="card ${{a.tagged?'tagged':'untagged'}}" onclick="toggleCard(this)">
      <div class="card-main">
        <div class="card-left"><span class="tb ${{a.tier}}">${{a.tier}}</span><span class="tc">${{a.type}}</span><span class="ab ${{a.nt}}">${{ntL[a.nt]}}</span></div>
        <div class="card-body">
          <div class="an" title="${{a.adName}}">${{a.adName}}</div>
          <div class="mr"><span class="im sp">${{fmtUSD(a.spend)}}</span><span class="im ro ${{rc(a.roas)}}">ROAS ${{fmtROAS(a.roas)}}</span><span class="im pu">${{fmtNum(a.purch)}} purchases</span></div>
        </div>
        <div class="ch-wrap"><div class="ch"><svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 5l5 5 5-5"/></svg></div></div>
      </div>
      <div class="card-expand">
        <div class="expand-row">
          <div class="thumb-wrap">${{th}}</div>
          <div class="expand-details">
            <div class="eg"><div class="el">Campaign</div><div class="ev">${{a.camp}}</div></div>
            <div class="eg"><div class="el">Ad Set</div><div class="ev">${{a.adset}}</div></div>
            <div class="eg"><div class="el">Revenue</div><div class="ev" style="color:var(--green)">${{fmtUSD(a.rev)}}</div></div>
            <div class="eg"><div class="el">CPA</div><div class="ev">${{fmtUSD(a.cpa)}}</div></div>
            ${{a.ctr?`<div class="eg"><div class="el">CTR</div><div class="ev">${{a.ctr}}%</div></div>`:''}}
            ${{a.hook!=null?`<div class="eg"><div class="el">Hook Rate</div><div class="ev">${{a.hook}}%</div></div>`:''}}
            ${{a.hold!=null?`<div class="eg"><div class="el">Hold Rate</div><div class="ev">${{a.hold}}%</div></div>`:''}}
            <div class="eg"><div class="el">Status</div><div class="ev" style="color:${{a.tagged?'var(--green)':'var(--amber)'}}">${{a.tagged?'✅ WIN Tagged':'⚡ Needs Action'}}</div></div>
          </div>
        </div>
        ${{(fb||lp)?`<div class="ea">${{fb}}${{lp}}</div>`:''}}
      </div>
    </div>`;
  }}
  const unads=ads.filter(a=>!a.tagged),tads=ads.filter(a=>a.tagged);
  let html='';
  if(!ads.length){{html=`<div class="empty-state"><h3>No ads match</h3><p>Try adjusting filters</p></div>`;}}
  else{{
    if(unads.length)html+=`<div class="section-head"><h2>🔴 Needs Action</h2><div class="sh-count">${{unads.length}} ads</div><div class="sh-line"></div></div><div class="cards">${{unads.map(card).join('')}}</div>`;
    if(tads.length)html+=`<div class="section-head"><h2>✅ Already Tagged</h2><div class="sh-count">${{tads.length}} ads</div><div class="sh-line"></div></div><div class="cards">${{tads.map(card).join('')}}</div>`;
  }}
  document.getElementById('cards-wrap').innerHTML=html;
  document.querySelectorAll('.card').forEach((c,i)=>c.style.animationDelay=Math.min(i,30)*0.018+'s');
}}
render();
</script>
</body>
</html>'''

# ── MAIN ──────────────────────────────────────────────
if __name__ == '__main__':
    import sys

    # Find XLS file — look for any xlsx/xls in repo root
    xls_files = glob.glob('*.xlsx') + glob.glob('*.xls') + glob.glob('exports/*.xlsx') + glob.glob('exports/*.xls')
    if not xls_files:
        print("No XLS file found in repo root or exports/ folder")
        sys.exit(1)

    # Use the most recently modified one
    xls_path = max(xls_files, key=os.path.getmtime)
    xls_filename = os.path.basename(xls_path)
    print(f"Processing: {xls_path}")

    ads = process_xls(xls_path)
    if not ads:
        print("No qualifying ads found")
        sys.exit(0)

    fetch_creatives(ads)

    html = generate_html(ads, xls_filename)
    with open('index.html', 'w') as f:
        f.write(html)

    untagged = sum(1 for a in ads if not a['tagged'])
    tagged   = sum(1 for a in ads if a['tagged'])
    print(f"✅ Done — {untagged} needs action | {tagged} tagged | {len(ads)} total")
    print(f"   index.html written ({len(html)} chars)")
