"""Traduction vers l'arabe des champs que l'Excel ne fournit pas en arabe, par l'outil de traduction en ligne
(traduction_auto.py : Google Translate + cache traductions_auto.json).

Mêmes fonctions qu'avant (lieu, discipline, specialite, mention, formation, nom_complet), utilisées par app.enrichir_ar.
- Une valeur arabe saisie dans l'Excel n'est jamais écrasée (c'est app.py qui décide d'appeler ces fonctions).
- Corrections manuelles, prioritaires sur l'outil : traductions_ar.json
  (format : {"lieux": {"taroudant": "تارودانت"}, "disciplines": {...}, "noms": {...}, "prenoms": {...}, "mentions": {...}, "formations": {...}}).
  On peut aussi corriger directement traductions_auto.json (le cache de l'outil).
- Outil injoignable -> '' : le champ reste « ……… » et le dossier est signalé dans RAPPORT.txt.
"""
import json, os, re, unicodedata
import traduction_auto as TA


def _sa(s): return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn')
def cle(s):                      # « El Mousadik » -> « elmousadik »
    return re.sub(r'[^a-z0-9]+', '', _sa(s).lower())
def _norm(s): return re.sub(r'\s+', ' ', _sa(s).lower().replace('’', "'")).strip(' .')
def _propre(s):                  # « EL MOUSADIK » -> « El Mousadik » (l'outil translittère mieux les noms en casse normale)
    return ' '.join('-'.join(x[:1].upper() + x[1:].lower() for x in m.split('-')) for m in str(s or '').split())


def _perso(nom):
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'traductions_ar.json')
    try:
        with open(p, encoding='utf-8') as f: d = json.load(f).get(nom, {})
    except Exception: return {}
    return {(cle(k) if nom in ('noms', 'prenoms') else _norm(k)): v for k, v in d.items()}


def _traduire(s, groupe, nature):
    """Correction manuelle (traductions_ar.json) sinon outil en ligne ; '' si rien."""
    s = re.sub(r'\s+', ' ', str(s or '')).strip()
    if not s: return ''
    k = cle(s) if groupe in ('noms', 'prenoms') else _norm(s)
    return _perso(groupe).get(k) or TA.traduire(s, nature)


def lieu(s): return _traduire(_propre(s) if s and (str(s).isupper() or str(s).islower()) else s, 'lieux', 'lieu')
def discipline(s): return _traduire(s, 'disciplines', 'terme')
specialite = discipline
def mention(s): return _traduire(s, 'mentions', 'terme')
def formation(s): return _traduire(s, 'formations', 'terme')


def nom_complet(nom, prenom):
    """« EL MOUSADIK » + « Abdelhamid » -> « الموساديق عبد الحميد » ; '' si l'un des deux n'a pas pu être traduit."""
    n = _traduire(_propre(nom), 'noms', 'nom') if str(nom or '').strip() else ''
    p = _traduire(_propre(prenom), 'prenoms', 'prenom') if str(prenom or '').strip() else ''
    return f"{n} {p}" if n and p else ''