"""Genre (M / F) des membres du jury d'après leur prénom, via l'API genderize.io.

Ordre de décision (voir sexes_jury) :
  1. sexe saisi dans l'Excel (colonne J{i}_SEXE) ;
  2. rôle déjà au féminin (« Présidente », « Directrice ») ;
  3. listes locales de prénoms (deviner_sexe dans app.py) : aucun appel réseau ;
  4. API genderize.io (pays MA), résultats mis en cache dans genre_cache.json.
Prénom inconnu ou API injoignable -> '' (le diplôme reste au masculin, signalé dans RAPPORT.txt).

Variables d'environnement :
  GENDERIZE_API_KEY  clé genderize.io (facultative : sans clé, 100 prénoms / jour)
  GENRE_API=0        désactive l'appel réseau (listes locales + cache uniquement)
"""
import json, os, threading, unicodedata, urllib.parse, urllib.request

API = 'https://api.genderize.io'
SEUIL = 0.80                       # probabilité minimale pour accepter la réponse
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'genre_cache.json')
_verrou = threading.Lock()


def cle(prenom):
    """« Fatima-Zahra » -> 'fatima' (premier prénom, minuscules, sans accents)."""
    s = ''.join(c for c in unicodedata.normalize('NFD', str(prenom or '')) if unicodedata.category(c) != 'Mn')
    mots = s.lower().replace('-', ' ').split()
    return mots[0] if mots else ''


def _lire_cache():
    try:
        with open(CACHE, encoding='utf-8') as f: return json.load(f)
    except Exception: return {}


def _ecrire_cache(c):
    try:
        tmp = CACHE + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f: json.dump(c, f, ensure_ascii=False, indent=0, sort_keys=True)
        os.replace(tmp, CACHE)
    except Exception: pass


def _api(noms):
    """{prenom: 'M'|'F'|''} pour au plus 10 prénoms ; {} si l'API ne répond pas."""
    q = [('name[]', n) for n in noms] + [('country_id', 'MA')]
    if os.environ.get('GENDERIZE_API_KEY'): q.append(('apikey', os.environ['GENDERIZE_API_KEY']))
    try:
        with urllib.request.urlopen(f'{API}/?{urllib.parse.urlencode(q)}', timeout=5) as rep:
            data = json.load(rep)
    except Exception:
        return {}
    out = {}
    for x in data if isinstance(data, list) else [data]:
        g, p = x.get('gender'), x.get('probability') or 0
        out[cle(x.get('name'))] = ('F' if g == 'female' else 'M' if g == 'male' else '') if p >= SEUIL else ''
    return out


def genres_api(prenoms):
    """{clé du prénom: 'M'|'F'|''} — cache d'abord, puis un seul appel groupé pour les prénoms nouveaux."""
    cles = [k for k in dict.fromkeys(cle(p) for p in prenoms) if k]
    with _verrou:
        cache = _lire_cache()
        nouveaux = [k for k in cles if k not in cache]
        if nouveaux and os.environ.get('GENRE_API', '1') != '0':
            trouves = {}
            for i in range(0, len(nouveaux), 10):
                trouves.update(_api(nouveaux[i:i + 10]))
            if trouves:                              # on ne met en cache que les réponses réellement reçues
                cache.update(trouves); _ecrire_cache(cache)
    return {k: cache.get(k, '') for k in cles}


def sexes_jury(membres, local):
    """membres : liste de (prenom, sexe_saisi, role) ; local : fonction prenom -> 'M'|'F'|''.
    Retourne [(sexe, source)] avec source dans {'excel', 'rôle', 'liste', 'api', ''}."""
    res = [None] * len(membres); a_demander = []
    for n, (pre, sx, role) in enumerate(membres):
        sx = str(sx or '').strip().upper()[:1]
        r = ''.join(c for c in unicodedata.normalize('NFD', str(role or '')) if unicodedata.category(c) != 'Mn').lower()
        if sx in ('M', 'F'): res[n] = (sx, 'excel')
        elif 'presidente' in r or 'directrice' in r: res[n] = ('F', 'rôle')
        elif local(pre): res[n] = (local(pre), 'liste')
        elif cle(pre): a_demander.append(n)
        else: res[n] = ('', '')
    if a_demander:
        g = genres_api([membres[n][0] for n in a_demander])
        for n in a_demander:
            s = g.get(cle(membres[n][0]), ''); res[n] = (s, 'api' if s else '')
    return res