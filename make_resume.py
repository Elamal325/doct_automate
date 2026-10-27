"""Crée modeles/doctorat/Resume.docx (« Résumé des avis des rapporteurs », champs {{...}})
à partir de l'exemple réel OUBRAIME : modeles/sources/resume_exemple.docx
(= Resume_avis_rapporteurs_oubrame.docx).

Lancer :  python make_resume.py
Les valeurs de l'exemple sont remplacées par des champs. Les tableaux sont prévus pour
3 directeurs et 6 rapporteurs : les lignes inutilisées sont supprimées à la génération
grâce au marqueur {{??CLE}} (ligne de tableau supprimée si la valeur est vide, voir docfill.py).
"""
import copy, os, re
from docfill import process, replace_in_paragraph, P, _ts, _w

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'modeles', 'sources', 'resume_exemple.docx')
OUT = os.path.join(BASE, 'modeles', 'doctorat', 'Resume.docx')
NB_DIR, NB_RAP = 3, 6


def texte(el):
    return ''.join(t.text or '' for t in _ts(el))


def remplacer(el, regles, vus):
    """Chaque règle (regex, champ) est appliquée dans les paragraphes de el ; vus compte les règles appliquées."""
    for p in list(el.iter(P)):
        for i, (pat, new) in enumerate(regles):
            while replace_in_paragraph(p, pat, new):
                vus[i] = vus.get(i, 0) + 1


def tableau(root, entete):
    """Tableau dont la 1re ligne commence par `entete`."""
    return next(t for t in root.iter(_w('tbl')) if texte(t.find(_w('tr'))).startswith(entete))


def lignes_personnes(tbl, nb, prefixe, colonnes):
    """La 1re ligne de données sert de modèle : elle est répétée nb fois avec des champs, les lignes d'exemple sont retirées."""
    rows = tbl.findall(_w('tr'))
    modele, exemples = rows[1], rows[1:]
    for k in range(1, nb + 1):
        tr = copy.deepcopy(modele)
        for c, (tc, nom) in enumerate(zip(tr.findall(_w('tc')), colonnes)):
            p = tc.find(P)
            champ = ('{{??%s_%d}}' % (prefixe, k) if c == 0 else '') + '{{%s_%d_%s}}' % (prefixe, k, nom)
            replace_in_paragraph(p, r'^.*$', champ)
        modele.addprevious(tr)
    for tr in exemples: tbl.remove(tr)


def construire():
    vus = {}
    R = [
        (r'OUBRAIME, Mohamed', '{{RES_NOM}}'), (r'JB362327', '{{CIN}}'), (r'04-22026721', '{{CNE}}'),
        (r"Élaboration d'algorithmes.*?complexes", '{{TITRE}}'),
        (r'^DCED$', '{{RES_TYPE}}'), (r'2022/2023', '{{RES_ANNEE_INSCRIPTION}}'),
        (r'^MIA$', '{{RES_FORMATION}}'), (r'06/07/2026', '{{DATE_SOUTENANCE}}'),
    ]

    def fn(root):
        if root.find(_w('body')) is None:                     # en-têtes / pieds de page : rien à remplacer
            return
        lignes_personnes(tableau(root, 'Directeur de thèse'), NB_DIR, 'DIR', ['NOM', 'GRADE', 'DISC', 'ETAB'])
        lignes_personnes(tableau(root, 'Rapporteur'), NB_RAP, 'RAP', ['NOM', 'GRADE', 'DISC', 'ETAB', 'AVIS'])
        remplacer(root, R, vus)

    data = process(SRC, fn)
    return data, [R[i][0] for i in range(len(R)) if i not in vus]


if __name__ == '__main__':
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    data, manque = construire()
    open(OUT, 'wb').write(data)
    print('créé :', OUT, '- règles non appliquées :', manque or 'aucune')