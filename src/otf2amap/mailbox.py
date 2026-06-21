"""Récupération des tableurs de distribution dans la boîte Gmail (IMAP).

Les mails proviennent d'AmapJ (ex. saint-malo@m.amapj.fr). Pour chaque
livraison, un mail par contrat (Légumes, Pain, …) avec le tableur en pièce
jointe. L'objet contient la date de livraison, ex. :

    AMAP du Bocage - Feuille de livraison du mercredi 24 juin 2026

`parse_date_livraison` en extrait la date ; le dossier de destination est
ensuite nommé d'après la semaine ISO (cf. naming.prefixe_semaine).
"""

import email
import imaplib
import re
from datetime import datetime
from email.header import decode_header
from pathlib import Path

MOIS = {
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4,
    "mai": 5, "juin": 6, "juillet": 7, "août": 8, "aout": 8,
    "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12,
}

# « … du [mercredi] 24 juin 2026 » — le jour de semaine est optionnel.
_DATE_RE = re.compile(
    r"(\d{1,2})\s+([a-zàâäéèêëîïôöùûüç]+)\s+(\d{4})", re.IGNORECASE
)

# Corps du mail : « Nom du contrat : Légumes (octobre 2025 - mars 2026) ».
# Indispensable car objet + expéditeur sont IDENTIQUES pour tous les contrats
# d'une même livraison (Légumes, Pain, Pommes, Poires…).
_CONTRAT_RE = re.compile(r"Nom du contrat\s*:\s*(.+)", re.IGNORECASE)


def _decode(value):
    """Décode un en-tête MIME (objet, nom de fichier) en str lisible."""
    if value is None:
        return ""
    parts = []
    for chunk, enc in decode_header(value):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            parts.append(chunk)
    return "".join(parts)


def parse_date_livraison(sujet):
    """Extrait la date de livraison de l'objet du mail → datetime, ou None.

    Reconnaît « … 24 juin 2026 » (mois français, accents ou non).
    """
    m = _DATE_RE.search(sujet or "")
    if not m:
        return None
    jour, mois_nom, annee = m.group(1), m.group(2).lower(), m.group(3)
    mois = MOIS.get(mois_nom)
    if not mois:
        return None
    try:
        return datetime(int(annee), mois, int(jour))
    except ValueError:
        return None


def _text_body(msg):
    """Texte brut du corps du mail (concatène les parties text/plain)."""
    parts = []
    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                parts.append(payload.decode(charset, errors="replace"))
    if not parts:                                    # repli : payload simple
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            parts.append(payload.decode(charset, errors="replace"))
    return "\n".join(parts)


def extract_contrat(msg):
    """Nom du contrat lu dans le corps (« Nom du contrat : … »), ou None.

    C'est le seul élément qui distingue les mails d'une même livraison.
    """
    m = _CONTRAT_RE.search(_text_body(msg))
    return m.group(1).strip() if m else None


def connect(host, user, password):
    """Ouvre une connexion IMAP4 SSL authentifiée et sélectionne INBOX."""
    imap = imaplib.IMAP4_SSL(host)
    imap.login(user, password)
    imap.select("INBOX")
    return imap


def _attachments(msg):
    """Itère les pièces jointes (nom_décodé, octets) d'un message email."""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None and not part.get_filename():
            continue
        filename = _decode(part.get_filename())
        if not filename:
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        yield filename, payload


def find_distributions(imap, expediteur, objet_motif):
    """Cherche les mails de distribution.

    Retourne une liste de dicts {sujet, date (datetime|None), contrat, message}
    pour les mails de `expediteur` dont l'objet contient `objet_motif`, triés du
    plus ancien au plus récent par date de livraison. `contrat` est lu dans le
    corps du mail (cf. extract_contrat) — c'est ce qui distingue Légumes/Pain/….
    """
    typ, data = imap.search(None, "FROM", f'"{expediteur}"')
    if typ != "OK":
        return []
    resultats = []
    for num in data[0].split():
        typ, raw = imap.fetch(num, "(RFC822)")
        if typ != "OK" or not raw or raw[0] is None:
            continue
        msg = email.message_from_bytes(raw[0][1])
        sujet = _decode(msg.get("Subject"))
        if objet_motif.lower() not in sujet.lower():
            continue
        resultats.append({
            "sujet": sujet,
            "date": parse_date_livraison(sujet),
            "contrat": extract_contrat(msg),
            "message": msg,
        })
    resultats.sort(key=lambda r: (r["date"] is None, r["date"] or datetime.min))
    return resultats


def save_attachments(msg, dest_dir):
    """Écrit les pièces jointes du message dans dest_dir. Retourne les chemins.

    Le dossier est créé si besoin. Les fichiers existants sont écrasés (le
    tableur le plus récent fait foi).
    """
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    chemins = []
    for filename, payload in _attachments(msg):
        cible = dest_dir / Path(filename).name      # garde le nom AmapJ, sans chemin
        cible.write_bytes(payload)
        chemins.append(cible)
    return chemins
