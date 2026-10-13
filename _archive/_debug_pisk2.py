import pandas as pd
LARK_PATH = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\Målebrev\LARK_planter.xlsx'
lark_raw = pd.read_excel(LARK_PATH, sheet_name='Tabell1', header=0)
lark_raw.columns = [c.strip() for c in lark_raw.columns]

pisk_types = ['772_Trær_Hengebjørk(pisk)', '772_Trær_Spisslønn(pisk)', '772_Trær_Småbladlink(pisk)', '772_Pisk']
for t in pisk_types:
    rows = lark_raw[lark_raw['PGF_Info.PGF_Mengdetype'] == t]
    total = 0
    for _, r in rows.iterrows():
        try:
            total += float(str(r['FOB_Info.FOB_Mengde']).split(';')[0])
        except:
            pass
    print(f'{t}: {len(rows)} rader, sum = {total}')

print()
pisk = lark_raw[lark_raw['PGF_Info.PGF_Mengdetype'] == '772_Pisk']
print('772_Pisk eksempler:')
for _, r in pisk.head(5).iterrows():
    col_mengde = r['FOB_Info.FOB_Mengde']
    col_post = str(r['FOB_Info.FOB_Mengdelistepost'])[:80]
    print(f'  mengde={col_mengde!r}  post={col_post!r}')

# Sum av alle mengder i 772_Pisk
print()
for _, r in pisk.head(3).iterrows():
    print('  alle mengder:', r['FOB_Info.FOB_Mengde'])
    print('  alle poster:', r['FOB_Info.FOB_Mengdelistepost'])
