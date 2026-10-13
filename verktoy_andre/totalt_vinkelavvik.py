import math

def beregn_arctan(a: float, b: float, i_grader: bool = True) -> float:
    """
    Beregner atan(sqrt((tan(a)*2 + tan(b)*2))).
    
    :param a: Vinkel a (i grader eller radianer)
    :param b: Vinkel b (i grader eller radianer)
    :param i_grader: True hvis vinklene er i grader, False hvis de er i radianer
    :return: Resultatet av uttrykket i radianer
    """
    if i_grader:
        a = math.radians(a)  # Konverterer til radianer
        b = math.radians(b)
    
    uttrykk = (math.tan(a)) ** 2 + (math.tan(b)) ** 2
    return math.atan(math.sqrt(uttrykk))

if __name__ == "__main__":
    a = float(input("Oppgi vinkel a (i grader): "))
    b = float(input("Oppgi vinkel b (i grader): "))
    resultat = beregn_arctan(a, b)
    
    print(f"Resultat i radianer: {resultat:.4f}")
    print(f"Resultat i grader: {math.degrees(resultat):.2f}")
    
    # Holder vinduet åpent
    input("\nTrykk Enter for å avslutte...")

