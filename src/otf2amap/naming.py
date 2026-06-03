"""Nommage par défaut des fichiers de sortie (préfixe semaine ISO)."""

from datetime import datetime, timedelta

SORTIE = "feuille_paniers_amap"


def prefixe_semaine(date_str):
    """
    Calcule le préfixe YYYY_SNN à partir d'une date de panier DD/MM/YYYY.
    La distribution a lieu le mercredi suivant (ou le jour même si c'est mercredi).
    Retourne None si la date est invalide.
    """
    try:
        d = datetime.strptime(date_str, '%d/%m/%Y')
        jours = (2 - d.weekday()) % 7   # 0 si mercredi, sinon jours jusqu'au prochain
        mercredi = d + timedelta(days=jours)
        annee, semaine, _ = mercredi.isocalendar()
        return f"{annee}_S{semaine:02d}"
    except (ValueError, TypeError):
        return None
