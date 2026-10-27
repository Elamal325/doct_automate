"""Crée modeles/doctorat/Attestation.docx et modeles/doctorat/Diplome.docx (champs {{...}})
à partir des exemples réels OUBRAI (modeles/sources/*.docx, convertis depuis les RTF fournis).

Lancer :  python make_attestation_diplome.py
Les valeurs de l'exemple sont remplacées par des champs ; le jury est prévu pour 10 membres
(les lignes inutilisées sont supprimées à la génération grâce au marqueur {{?JURY_k}}).
"""
import copy, os, re, sys
from docfill import process, replace_in_paragraph, P, _ts, _w

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'modeles', 'sources')
OUT = os.path.join(BASE, 'modeles', 'doctorat')
NB_JURY = 10


def pua(s):                       # texte « Barcode 3 of 9 » tel que Word l'enregistre (U+F000 + code ASCII)
    return ''.join(chr(0xF000 + ord(c)) for c in s)


BAR_CNE = re.escape(pua('*04-22026721*'))
BAR_CIN = re.escape(pua('*JB362327*'))


def texte(p):
    return ''.join(t.text or '' for t in _ts(p))


def appliquer(root, regles, stats):
    """Chaque règle (regex, remplacement) est appliquée partout, dans l'ordre donné (les plus longues d'abord)."""
    for p in list(root.iter(P)):
        for i, (pat, new) in enumerate(regles):
            while replace_in_paragraph(p, pat, new):
                stats[i] = stats.get(i, 0) + 1


def paragraphes_feuilles(root):
    return [p for p in root.iter(P) if p.find('.//' + P) is None]


def fixer_interligne(p, ligne):
    """Interligne exact (en 20e de point) pour un paragraphe."""
    ppr = p.find(_w('pPr'))
    if ppr is None:
        from lxml import etree
        ppr = etree.Element(_w('pPr')); p.insert(0, ppr)
    sp = ppr.find(_w('spacing'))
    if sp is None:
        from lxml import etree
        sp = etree.SubElement(ppr, _w('spacing'))
        # ordre du schéma : spacing doit précéder ind/jc/rPr
        for nom in ('ind', 'contextualSpacing', 'mirrorIndents', 'jc', 'textDirection', 'textAlignment', 'outlineLvl', 'rPr'):
            e = ppr.find(_w(nom))
            if e is not None:
                e.addprevious(sp); break
    sp.set(_w('lineRule'), 'exact'); sp.set(_w('line'), str(ligne))


def finir_attestation(root):
    """Mise en page : tout sur 1 page + espace pour la signature sous « Le Doyen »."""
    pars = [p for p in root.iter(P) if p.find('.//' + P) is None]
    for p in pars:
        if not texte(p).strip() and p.getparent() is not None and p.getparent().tag == _w('body'):
            fixer_interligne(p, 160 if p not in pars[:6] else 280)          # lignes vides : plus serrées
    dernier = [p for p in pars if texte(p).strip() == 'Le Doyen de la Faculté des Sciences'][-1]
    dernier.find(_w('pPr')).find(_w('spacing')).set(_w('after'), '1600')   # signature + cachet (ne crée jamais de 2e page)
    for m in root.iter(_w('pgMar')):
        m.set(_w('top'), '900'); m.set(_w('bottom'), '1300')


def construire_attestation():
    src = os.path.join(SRC, 'attestation_exemple.docx')
    stats = {}
    R = [
        (BAR_CNE, '{{BAR_CNE}}'), (BAR_CIN, '{{BAR_CIN}}'),
        (r'Monsieur', '{{CIVILITE_LONGUE}}'),
        (r'OUBRAIME Mohamed', '{{NOM_COMPLET}}'),
        (r'ابرايم محمد', '{{NOM_AR_COMPLET}}'),
        (r'né le', '{{NE}} le'),
        (r'05/07/1985', '{{DATE_NAISSANCE}}'),
        (r'Targua Ntouchka Chtouka-Ait Baha', '{{LIEU_NAISSANCE}}'),
        (r'تركا نتوشكا شتوكة-ايت باها', '{{LIEU_NAISSANCE_AR}}'),
        (r'04-22026721', '{{CNE}}'), (r'JB362327', '{{CIN}}'),
        (r'Mathématiques, Informatique et Applications', '{{FORMATION}}'),
        (r'Sciences et Techniques et Sciences Médicales', '{{CED}}'),
        (r'06/07/2026', '{{DATE_SOUTENANCE}}'),
        (r"Élaboration d'algorithmes.*?complexes", '{{TITRE}}'),
        (r'Informatique, sciences de données et intelligence artificielle', '{{SPECIALITE}}'),
        (r'Informatique', '{{DISCIPLINE}}'),            # il ne reste que les 2 « discipline »
        (r'Très Honorable', '{{MENTION}}'),
        (r'Agadir, le 08/10/2026', 'Agadir, le {{DATE_EDITION}}'),
    ]

    def fn(root):
        # 1) bloc jury : 6 membres d'exemple -> NB_JURY blocs de 2 paragraphes
        feuilles = paragraphes_feuilles(root)
        p_nom = next((p for p in feuilles if 'Nouh IZEM' in texte(p)), None)
        if p_nom is None:                      # en-têtes / pieds de page : rien à faire pour le jury
            appliquer(root, R, stats)
            return
        p_etab = p_nom.getnext()
        assert 'Président' in texte(p_etab), texte(p_etab)
        derniers = next(p for p in feuilles if 'Azidine GUEZZAZ' in texte(p)).getnext()
        sample = []
        cur = p_nom
        while True:
            sample.append(cur)
            if cur is derniers:
                break
            cur = cur.getnext()
        assert len(sample) == 12, len(sample)
        modele_nom, modele_etab = copy.deepcopy(p_nom), copy.deepcopy(p_etab)
        for k in range(1, NB_JURY + 1):
            a, b = copy.deepcopy(modele_nom), copy.deepcopy(modele_etab)
            replace_in_paragraph(a, r"Nouh IZEM, Professeur de l'Enseignement Supérieur, ",
                                 '{{?JURY_%d}}{{JURY_%d_NOM_GRADE}} ' % (k, k))
            replace_in_paragraph(b, r'Faculté des Sciences, Agadir', '{{?JURY_%d}}{{JURY_%d_ETABLISSEMENT}}' % (k, k))
            replace_in_paragraph(b, r'Président', '{{JURY_%d_ROLE}}' % k)
            p_nom.addprevious(a); p_nom.addprevious(b)
        for p in sample:
            p.getparent().remove(p)
        finir_attestation(root)
        appliquer(root, R, stats)
    data = process(src, fn)
    manque = [R[i][0] for i in range(len(R)) if i not in stats]
    return data, manque


def construire_diplome():
    src = os.path.join(SRC, 'diplome_exemple.docx')
    stats = {}
    R = [
        (BAR_CNE, '{{BAR_CNE}}'), (BAR_CIN, '{{BAR_CIN}}'),
        # dates du PV
        (r'06\s+يوليوز\s+2026', '{{DATE_SOUT_AR}}'), (r'06 juillet 2026', '{{DATE_SOUTENANCE_LONG}}'),
        # identité
        (r'ابرايم محمد', '{{NOM_AR_COMPLET}}'), (r'OUBRAIME Mohamed', '{{NOM_COMPLET}}'),
        (r'بتركا نتوشكا شتوكة-ايت باها', 'ب{{LIEU_NAISSANCE_AR}}'),
        (r'05\s+يوليوز\s+1985', '{{DATE_NAISSANCE_AR}}'),
        (r'Targua Ntouchka Chtouka-Ait Baha', '{{LIEU_NAISSANCE}}'),
        (r'05 juillet 1985', '{{DATE_NAISSANCE_LONG}}'),
        (r'JB362327', '{{CIN}}'), (r'04-22026721', '{{CNE}}'),
        (r'كلية العلوم بأگادير', '{{ETABLISSEMENT_AR}}'), (r'Faculté des Sciences, Agadir', '{{ETABLISSEMENT}}'),
        # cursus
        (r'الاعلاميات،\s+علوم المعلومات والذكاء الاصطناعي', '{{SPECIALITE_AR}}'),
        (r'Informatique, sciences de données et intelligence artificielle', '{{SPECIALITE}}'),
        (r'الإعلاميات', '{{DISCIPLINE_AR}}'), (r'Informatique', '{{DISCIPLINE}}'),
        (r'العلوم والتقنيات والعلوم الطبية', '{{CED_AR}}'), (r'Sciences et Techniques et Sciences Médicales', '{{CED}}'),
        (r"Elaboration d'algorithmes.*?complexes", '{{TITRE}}'),
        (r'مشرف جدا', '{{MENTION_AR}}'), (r'Très Honorable', '{{MENTION}}'),
        # date d'édition
        (r'08\s+اكتوبر\s+2026', '{{DATE_EDITION_AR}}'), (r'08 octobre 2026', '{{DATE_EDITION_LONG}}'),
    ]

    def fn(root):
        feuilles = paragraphes_feuilles(root)
        p0 = next((p for p in feuilles if 'إزام نوح' in texte(p)), None)
        if p0 is None:
            appliquer(root, R, stats)
            return
        dernier = next(p for p in feuilles if 'گزاز عز الدين' in texte(p))
        sample, cur = [], p0
        while True:
            sample.append(cur)
            if cur is dernier:
                break
            cur = cur.getnext()
        assert len(sample) == 6, len(sample)
        modele = copy.deepcopy(p0)
        for k in range(1, NB_JURY + 1):
            q = copy.deepcopy(modele)
            replace_in_paragraph(q, r'إزام نوح', '{{?JURY_%d}}{{JURY_%d_NOM_AR}}' % (k, k))
            replace_in_paragraph(q, r'أستاذ التعليم العالي ، ', '{{JURY_%d_GRADE_AR}} ' % k)
            replace_in_paragraph(q, r'رئيسا', '{{JURY_%d_ROLE_AR}}' % k)
            p0.addprevious(q)
        for p in sample:
            p.getparent().remove(p)
        for m in root.iter(_w('pgMar')):       # marge basse réduite : le jury (jusqu'à 10 membres) tient sur 1 page
            m.set(_w('bottom'), '500'); m.set(_w('top'), '600')
        appliquer(root, R, stats)
    data = process(src, fn)
    manque = [R[i][0] for i in range(len(R)) if i not in stats]
    return data, manque


# ------------------------------------------------------------------------------------------ HABILITATION
BAR_CNE_HU = re.escape(pua('*99-97177999*'))
BAR_CIN_HU = re.escape(pua('*IA72220*'))


def bloc_consecutif(feuilles, premier, dernier, n):
    p0 = next(p for p in feuilles if premier in texte(p))
    p_fin = next(p for p in feuilles if dernier in texte(p))
    bloc, cur = [], p0
    while True:
        bloc.append(cur)
        if cur is p_fin:
            break
        cur = cur.getnext()
    assert len(bloc) == n, (len(bloc), n)
    return p0, bloc


def construire_attestation_hu():
    src = os.path.join(SRC, 'attestation_hu_exemple.docx')
    stats = {}
    R = [
        (BAR_CNE_HU, '{{BAR_CNE}}'), (BAR_CIN_HU, '{{BAR_CIN}}'),
        (r'Monsieur', '{{CIVILITE_LONGUE}}'),
        (r'AJALLOUDA Lahbib', '{{NOM_COMPLET}}'), (r'اجلودة الحبيب', '{{NOM_AR_COMPLET}}'),
        (r'né le', '{{NE}} le'), (r'15/10/1979', '{{DATE_NAISSANCE}}'),
        (r'Kasba Tadla', '{{LIEU_NAISSANCE}}'), (r'قصبة تادلة', '{{LIEU_NAISSANCE_AR}}'),
        (r'99-97177999', '{{CNE}}'), (r'IA72220', '{{CIN}}'),
        (r'25/09/2026', '{{DATE_SOUTENANCE}}'),
        (r'Informatique et intelligence artificielle', '{{SPECIALITE}}'),
        (r'Informatique', '{{DISCIPLINE}}'),
        (r'Agadir, le 28/09/2026', 'Agadir, le {{DATE_EDITION}}'),
    ]

    def fn(root):
        feuilles = paragraphes_feuilles(root)
        if not any('Mustapha MACHKOUR' in texte(p) for p in feuilles):
            appliquer(root, R, stats)
            return
        p_nom, sample = bloc_consecutif(feuilles, 'Mustapha MACHKOUR', 'Mohamed LAZAAR', 5)   # 3 x 2 paragraphes - le dernier est lu plus bas
        sample = sample + [sample[-1].getnext()]
        assert len(sample) == 6 and 'Rapporteur' in texte(sample[-1])
        modele_nom, modele_etab = copy.deepcopy(sample[0]), copy.deepcopy(sample[1])
        for k in range(1, NB_JURY + 1):
            a, b = copy.deepcopy(modele_nom), copy.deepcopy(modele_etab)
            replace_in_paragraph(a, r"Mustapha MACHKOUR, Professeur de l'Enseignement Supérieur, ",
                                 '{{?JURY_%d}}{{JURY_%d_NOM_GRADE}} ' % (k, k))
            replace_in_paragraph(b, r'Faculté des Sciences, Agadir', '{{?JURY_%d}}{{JURY_%d_ETABLISSEMENT}}' % (k, k))
            replace_in_paragraph(b, r'Rapporteur et Président', '{{JURY_%d_ROLE}}' % k)
            p_nom.addprevious(a); p_nom.addprevious(b)
        for p in sample:
            p.getparent().remove(p)
        finir_attestation(root)
        appliquer(root, R, stats)
    data = process(src, fn)
    return data, [R[i][0] for i in range(len(R)) if i not in stats]


def construire_diplome_hu():
    src = os.path.join(SRC, 'diplome_hu_exemple.docx')
    stats = {}
    R = [
        (BAR_CNE_HU, '{{BAR_CNE}}'), (BAR_CIN_HU, '{{BAR_CIN}}'),
        (r'25\s+شتنبر\s+2026', '{{DATE_SOUT_AR}}'), (r'28\s+شتنبر\s+2026', '{{DATE_EDITION_AR}}'),
        (r'اجلودة الحبيب', '{{NOM_AR_COMPLET}}'), (r'AJALLOUDA Lahbib', '{{NOM_COMPLET}}'),
        (r'15\s+اكتوبر\s+1979', '{{DATE_NAISSANCE_AR}}'), (r'15/10/1979', '{{DATE_NAISSANCE}}'),
        (r'IA72220', '{{CIN}}'),
        (r'المعلوميات والذكاء الاصطناعي', '{{SPECIALITE_AR}}'), (r'Informatique et intelligence artificielle', '{{SPECIALITE}}'),
    ]

    def fn(root):
        feuilles = paragraphes_feuilles(root)
        if not any('مشكور مصطفى' in texte(p) for p in feuilles):
            appliquer(root, R, stats)
            return
        p0, sample = bloc_consecutif(feuilles, 'مشكور مصطفى', 'الأزعر محمد', 3)
        modele = copy.deepcopy(p0)
        for k in range(1, NB_JURY + 1):
            q = copy.deepcopy(modele)
            replace_in_paragraph(q, r'مشكور مصطفى', '{{?JURY_%d}}{{JURY_%d_NOM_AR}}' % (k, k))
            replace_in_paragraph(q, r'أستاذ التعليم العالي ، كلية العلوم ، أگادير', '{{JURY_%d_GRADE_ETAB_AR}}' % k)
            replace_in_paragraph(q, r'رئيسا', '{{JURY_%d_ROLE_AR}}' % k)
            p0.addprevious(q)
        for p in sample:
            p.getparent().remove(p)
        for m in root.iter(_w('pgMar')):
            m.set(_w('bottom'), '500'); m.set(_w('top'), '600')
        appliquer(root, R, stats)
    data = process(src, fn)
    return data, [R[i][0] for i in range(len(R)) if i not in stats]


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    HU = os.path.join(BASE, 'modeles', 'habilitation')
    for dossier, nom, f in ((OUT, 'Attestation.docx', construire_attestation), (OUT, 'Diplome.docx', construire_diplome),
                            (HU, 'Attestation.docx', construire_attestation_hu), (HU, 'Diplome.docx', construire_diplome_hu)):
        data, manque = f()
        open(os.path.join(dossier, nom), 'wb').write(data)
        print(os.path.basename(dossier), nom, 'créé -', 'règles non appliquées :', manque or 'aucune')
