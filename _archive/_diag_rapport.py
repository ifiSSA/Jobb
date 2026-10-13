from base_lesning import les_omega

omega_path = r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\APM FOB Gjennomgang\Omega_K7_April.xlsx"
df = les_omega(omega_path)

print("Kolonnar i df:", list(df.columns))
print()

post = "1303.01.05.241.11.4"
rad = df[df["Navn (Prosesskode)"] == post]
if rad.empty:
    print(f"Post {post} ikkje funne")
else:
    for col in rad.columns:
        val = rad.iloc[0][col]
        if val is not None and str(val) not in ("nan", ""):
            print(f"  {col}: {val!r}")
