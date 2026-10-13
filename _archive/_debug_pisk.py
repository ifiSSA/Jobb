import pandas as pd

LARK_PATH = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\Målebrev\LARK_planter.xlsx'
lark_raw = pd.read_excel(LARK_PATH, sheet_name='Tabell1', header=0)
lark_raw.columns = [c.strip() for c in lark_raw.columns]

def kategoriser(mt):
    mt = str(mt)
    if '772_Trær' in mt and '(pisk)' in mt: return 'Trær (pisk)'
    if '772_Trær' in mt: return 'Trær'
    if '772_Buskfelt' in mt or mt == '772_Pisk': return 'Busker'
    if '772_Staudebed' in mt or mt == '772_Staudefelt': return 'Stauder'
    return None

pisk = lark_raw[lark_raw['PGF_Info.PGF_Mengdetype'].apply(lambda x: kategoriser(x) == 'Trær (pisk)')]
print(f'Antall rader i LARK med pisk-mengdetype: {len(pisk)}')
print()
print('Mengdetyper:')
print(pisk['PGF_Info.PGF_Mengdetype'].value_counts().to_string())
print()
print('FOB_Mengde og FOB_Mengdelistepost eksempler:')
for _, r in pisk.head(10).iterrows():
    post = r['FOB_Info.FOB_Mengdelistepost']
    mengde = r['FOB_Info.FOB_Mengde']
    mtype = r['PGF_Info.PGF_Mengdetype']
    print(f'  post={post!r}  mengde={mengde!r}  type={mtype!r}')

# Hva betyr mengden - er det 1 per tre eller en sum?
print()
print('Alle unike FOB_Mengde-verdier for pisk (kun første post):')
vals = []
for _, r in pisk.iterrows():
    try:
        vals.append(float(str(r['FOB_Info.FOB_Mengde']).split(';')[0]))
    except:
        pass
s = pd.Series(vals)
print(f'  Min: {s.min()}, Max: {s.max()}, Sum: {s.sum()}, Antall rader: {len(s)}')
print(f'  Verdier > 1: {(s > 1).sum()}')
print(f'  Verdier == 1: {(s == 1).sum()}')
