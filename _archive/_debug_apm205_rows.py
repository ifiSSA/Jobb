# -*- coding: utf-8 -*-
"""Debug APM205: vis rader og siste-segment-analyse."""
import sys
import os
import re
import glob

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ifi.pdf_maalebrev_til_excel import les_pdf_hoved, les_omega

omega_path = glob.glob(
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\**\Omega_K7_Mars.xlsx",
    recursive=True,
)[0]
apm205 = glob.glob(
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\**\*APM205*.pdf",
    recursive=True,
)[0]

omega = les_omega(omega_path)
rader = les_pdf_hoved(apm205)

print("APM205 rader (uten omega_data):")
print(f"  {'PNS':<14} {'POST':<32} {'Mengde':<12} {'Enhet':<8} Kommentar")
print("-" * 100)
for r in rader:
    pns    = r.get("PNS", "")
    post   = r.get("Post", "")
    mengde = str(r.get("Mengde fra APM", "") or "")
    enhet  = r.get("Enhet fra APM", "")
    komm   = (r.get("Kommentar", "") or "")[:30]
    print(f"  {pns:<14} {post:<32} {mengde:<12} {enhet:<8} {komm}")

print()
print("Digit-analyse av siste 2 segmenter:")
print(f"  {'POST':<32}  last2_digits  ->  Omega-match (is-subsequence)?")
print("-" * 90)
for r in rader:
    post = r.get("Post", "")
    enhet = r.get("Enhet fra APM", "")
    segs = re.split(r"[.\-]", post)
    last2_digits = re.sub(r"[^\d]", "", "".join(segs[-2:]))
    # Find Omega codes under the first 4 segments as prefix
    prefix4 = ".".join(segs[:4]) + "."
    kandidater = sorted([k for k in omega if k.startswith(prefix4)])
    # Subsequence test on last-2 digits
    def er_sub(sub_str, full_str):
        it = iter(full_str)
        return all(c in it for c in sub_str)

    matches = []
    for k in kandidater:
        k_segs = re.split(r"[.\-]", k)
        k_last = re.sub(r"[^\d]", "", "".join(k_segs[4:]))
        if er_sub(k_last, last2_digits):
            matches.append(k)
    enhet_match = [k for k in matches
                   if omega[k].get("enhet","").lower().replace("\u00b2","2").replace("\u00b3","3") ==
                      enhet.lower().replace("\u00b2","2").replace("\u00b3","3")]
    print(f"  {post:<32}  {last2_digits:<14}  subseq:{len(matches)} enhet-filter:{len(enhet_match)}")
    if len(enhet_match) <= 3:
        for k in enhet_match:
            print(f"    -> {k}  [{omega[k].get('enhet','')}]")
