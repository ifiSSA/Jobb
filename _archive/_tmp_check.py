import pandas as pd, re
ISY = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\Målebrev\Mengdelistpost NS3420 isylinker.xlsx'
df = pd.read_excel(ISY, engine='openpyxl')
df['Post nummer'] = df['Post nummer'].astype(str).str.strip()

# Post 1303.30.312.5.3
mask = df['Post nummer'].str.startswith('1303.30.312', na=False)
sub = df[mask][['Post nummer','Kode','All tekst']]
print('=== 1303.30.312.x i ISY ===')
for _, r in sub.head(5).iterrows():
    tekst = str(r['All tekst'])[:400].replace('_x000D_', '').replace('\n', ' | ')
    print(f"  [{r['Post nummer']}] {tekst[:300]}")
print()

# Lokalisering-mønster i 1303-poster
lok_mask = df['All tekst'].astype(str).str.contains('Lokalisering', na=False)
sample = df[lok_mask & df['Post nummer'].str.startswith('1303', na=False)].head(8)
print('=== Lokalisering-felt i 1303-poster ===')
for _, r in sample.iterrows():
    tekst = str(r['All tekst'])
    lok_pos = tekst.find('Lokalisering')
    snip = tekst[lok_pos:lok_pos+100].replace('_x000D_','').replace('\n',' ')
    print(f"  [{r['Post nummer']}] {snip}")
print()

# Søk etter mønster "[A-Z]{3} LP" i ISY All tekst (escape for raw string)
import warnings; warnings.filterwarnings('ignore')
m = df['All tekst'].astype(str).str.extract(r'([A-Z]{3})\s+LP', expand=False)
non_null = m.dropna()
print('=== XXX LP mønster i All tekst ===')
print(non_null.value_counts().head(20))
