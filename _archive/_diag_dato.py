from pdf_maalebrev_til_excel import les_omega, _dato_til_maaned_aar

omega_path = r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\APM FOB Gjennomgang\Omega_K7_April.xlsx"
data = les_omega(omega_path)

post = "1302.01.05.241.14.2"
info = data.get(post)
if info:
    raa = info.get("delmaal")
    konvertert = _dato_til_maaned_aar(raa)
    print(f"Post:       {post}")
    print(f"Raa dato:   {repr(raa)}  (type: {type(raa).__name__})")
    print(f"Konvertert: {konvertert}")
    if hasattr(raa, "day"):
        print(f"Dag i maaneden: {raa.day}  (> 20 => denne mnd, <= 20 => forrige mnd)")
else:
    print(f"Post {post} ikkje funne i Omega")
    like = [k for k in data if k.startswith("1302.01.05.241")]
    print("Liknande postar:", like[:10])

# Sjekk nokre postar rundt same prefix for samanlikning
print()
print("=== Alle 1302.01.05.241.14.x ===")
for k, v in sorted(data.items()):
    if k.startswith("1302.01.05.241.14"):
        raa = v.get("delmaal")
        print(f"  {k}  ->  raa={repr(raa)}  konv={_dato_til_maaned_aar(raa)}")
