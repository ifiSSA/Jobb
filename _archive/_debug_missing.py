import pandas as pd

LARK_PATH = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\Målebrev\LARK_planter.xlsx'
OMEGA_PATH = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\APM FOB Gjennomgang\Omega_K7_Mars.xlsx'

lark_raw = pd.read_excel(LARK_PATH, sheet_name='Tabell1', header=0)
lark_raw.columns = [c.strip() for c in lark_raw.columns]

omega_raw = pd.read_excel(OMEGA_PATH, sheet_name='BOQNT', header=2)
omega = omega_raw[['Navn (Prosesskode)', 'Enhet', 'Enhetspris', 'Kontraktsmengde']].copy()
omega.columns = ['post', 'enhet_omega', 'enhetspris', 'kontraktsmengde']
omega['post'] = omega['post'].astype(str).str.strip()
omega = omega[omega['post'].notna() & (omega['post'] != 'nan')]
omega_sum = omega.groupby('post', sort=False).agg(
    kontraktsmengde=('kontraktsmengde', 'sum'),
    enhet_omega=('enhet_omega', 'first'),
    enhetspris=('enhetspris', 'first'),
).reset_index()
omega_stripped_lookup = {p.replace('.', ''): p for p in omega_sum['post']}
omega_posts_set = set(omega_sum['post'])

TARGET = '1303.99.79.1.4'

# 1. Er posten i rå LARK data?
raw_hits = []
for _, r in lark_raw.iterrows():
    posts_raw   = str(r['FOB_Info.FOB_Mengdelistepost']) if pd.notna(r['FOB_Info.FOB_Mengdelistepost']) else ''
    mengder_raw = str(r['FOB_Info.FOB_Mengde'])          if pd.notna(r['FOB_Info.FOB_Mengde'])          else ''
    for i, p in enumerate(posts_raw.split(';')):
        if p.strip() == TARGET:
            mengder = mengder_raw.split(';')
            mengde = mengder[i].strip() if i < len(mengder) else '?'
            raw_hits.append({'posts_raw': posts_raw, 'mengde_pos': mengde})

print(f"Rader i LARK med {TARGET}: {len(raw_hits)}")
for h in raw_hits[:10]:
    print(f"  posts: {h['posts_raw']!r}  mengde[{TARGET}]: {h['mengde_pos']!r}")

# 2. Er posten i Omega?
print(f"\nI Omega direkte: {TARGET in omega_posts_set}")
canonical = omega_stripped_lookup.get(TARGET.replace('.', ''))
print(f"Omega-kandidat via stripped lookup: {canonical!r}")

# 3. Hva skjer i expand-løkken?
rows = []
for _, r in lark_raw.iterrows():
    posts_raw   = str(r['FOB_Info.FOB_Mengdelistepost']) if pd.notna(r['FOB_Info.FOB_Mengdelistepost']) else ''
    mengder_raw = str(r['FOB_Info.FOB_Mengde'])          if pd.notna(r['FOB_Info.FOB_Mengde'])          else ''
    enheter_raw = str(r['FOB_Info.FOB_Mengdeenhet'])     if pd.notna(r['FOB_Info.FOB_Mengdeenhet'])     else ''
    mengdetype  = str(r['PGF_Info.PGF_Mengdetype'])      if pd.notna(r['PGF_Info.PGF_Mengdetype'])      else ''
    fob_status  = str(r['FOB_Info.FOB_Status'])           if pd.notna(r['FOB_Info.FOB_Status'])           else ''

    posts   = [p.strip() for p in posts_raw.split(';')]
    mengder = [m.strip() for m in mengder_raw.split(';')]
    enheter = [e.strip() for e in enheter_raw.split(';')]

    for i, post in enumerate(posts):
        if not post or post == 'nan':
            continue
        original = post
        if post not in omega_posts_set:
            c = omega_stripped_lookup.get(post.replace('.', ''))
            if c:
                post = c
        if original == TARGET or post == TARGET:
            try:
                mengde = float(mengder[i]) if i < len(mengder) else 0.0
            except ValueError:
                mengde = 0.0
            rows.append({'original': original, 'resolved': post, 'mengde': mengde})

print(f"\nRader i expand med {TARGET} (original eller resolved): {len(rows)}")
for rw in rows[:10]:
    print(f"  {rw}")

# 4. Tell alle unike poster
all_rows = []
for _, r in lark_raw.iterrows():
    posts_raw   = str(r['FOB_Info.FOB_Mengdelistepost']) if pd.notna(r['FOB_Info.FOB_Mengdelistepost']) else ''
    mengder_raw = str(r['FOB_Info.FOB_Mengde'])          if pd.notna(r['FOB_Info.FOB_Mengde'])          else ''
    enheter_raw = str(r['FOB_Info.FOB_Mengdeenhet'])     if pd.notna(r['FOB_Info.FOB_Mengdeenhet'])     else ''
    mengdetype  = str(r['PGF_Info.PGF_Mengdetype'])      if pd.notna(r['PGF_Info.PGF_Mengdetype'])      else ''
    fob_status  = str(r['FOB_Info.FOB_Status'])           if pd.notna(r['FOB_Info.FOB_Status'])           else ''
    posts   = [p.strip() for p in posts_raw.split(';')]
    mengder = [m.strip() for m in mengder_raw.split(';')]
    enheter = [e.strip() for e in enheter_raw.split(';')]
    for i, post in enumerate(posts):
        if not post or post == 'nan':
            continue
        if post not in omega_posts_set:
            c = omega_stripped_lookup.get(post.replace('.', ''))
            if c:
                post = c
        try:
            mengde = float(mengder[i]) if i < len(mengder) else 0.0
        except ValueError:
            mengde = 0.0
        enhet = enheter[i] if i < len(enheter) else ''
        all_rows.append({'post': post, 'mengde_lark': mengde, 'enhet_lark': enhet, 'mengdetype': mengdetype, 'fob_status': fob_status})

lark_exp = pd.DataFrame(all_rows)
lark_sum = lark_exp.groupby('post', sort=False).agg(
    mengde_lark=('mengde_lark', 'sum'),
    enhet_lark=('enhet_lark', 'first'),
    mengdetype=('mengdetype', 'first'),
    fob_status=('fob_status', 'first'),
).reset_index()

print(f"\nTotalt unike poster i lark_sum: {len(lark_sum)}")
print(f"Er {TARGET} i lark_sum: {TARGET in set(lark_sum['post'])}")

# Finn alle 1303.99.79 poster i lark_sum
hits = lark_sum[lark_sum['post'].str.startswith('1303.99.79')]
print(f"\nAlle 1303.99.79* poster i lark_sum:")
print(hits[['post', 'mengde_lark']].to_string())
