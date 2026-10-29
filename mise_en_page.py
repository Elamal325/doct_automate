"""Alignement du diplôme généré (bilingue arabe / français).

aligner(docx) corrige, à chaque génération, deux défauts des modèles convertis depuis RTF :
  1. Tableau du préambule (« بناءً على القانون… » / « Vu la loi… ») : largeur exacte de la zone
     de texte, sans retrait ni marges de cellule. Le texte du préambule commence et finit
     ainsi au même bord que « يشهد رئيـــس جامعة ابن زهر أن » et « le Président… ».
  2. Partie française des lignes « arabe <tab> français » : dans une ligne arabe (droite à gauche),
     un signe final neutre saute au début du français (« : atteste que », « ……… : option »
     quand la valeur est vide). Une marque invisible de gauche à droite (U+200E) le garde à sa place.
  3. Codes-barres CNE / CIN : accrochés à la 1re ligne du préambule, dans la marge, comme sur le diplôme imprimé.
  4. En-tête : les lignes vides laissées par les codes-barres sont réduites, « Vu la loi… » suit de près « Diplôme de Doctorat »,
     et le titre est centré par l'alignement du paragraphe (plus par des espaces).
"""
import io, re, zipfile
from lxml import etree
from Codebarre import modules_code39

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(nom): return f'{{{W}}}{nom}'
LRM = '‎'
LATIN = re.compile(r'[A-Za-zÀ-ÿ]')


def _largeur_utile(root):
    """Largeur de la zone de texte (twips) d'après la 1re section : page - marges."""
    for s in root.iter(q('sectPr')):
        sz, mar = s.find(q('pgSz')), s.find(q('pgMar'))
        if sz is not None and mar is not None:
            g = lambda e, k: int(e.get(q(k)) or 0)
            return g(sz, 'w') - g(mar, 'left') - g(mar, 'right') - g(mar, 'gutter')
    return None


def _ajuster_tableau(tbl, largeur):
    pr = tbl.find(q('tblPr'))
    if pr is None: return False
    w = pr.find(q('tblW'))
    if w is None or (w.get(q('type')) or 'dxa') != 'dxa': return False
    if abs(int(w.get(q('w')) or 0) - largeur) > largeur * 0.03: return False   # pas un tableau pleine largeur : on n'y touche pas
    w.set(q('w'), str(largeur))
    ind = pr.find(q('tblInd'))
    if ind is not None: ind.set(q('w'), '0')
    mar = pr.find(q('tblCellMar'))
    if mar is not None:
        for c in mar:
            if c.tag in (q('left'), q('right'), q('start'), q('end')): c.set(q('w'), '0')
    for tc_mar in list(tbl.iter(q('tcMar'))):                                  # marges de cellule (108) côté arabe ou français
        tc_mar.getparent().remove(tc_mar)
    cols = tbl.find(q('tblGrid')).findall(q('gridCol')) if tbl.find(q('tblGrid')) is not None else []
    if cols:                                                                   # la 1re colonne absorbe la différence
        ecart = largeur - sum(int(c.get(q('w')) or 0) for c in cols)
        cols[0].set(q('w'), str(int(cols[0].get(q('w')) or 0) + ecart))
        for tr in tbl.findall(q('tr')):
            tcs = tr.findall(q('tc'))
            tcw = tcs[0].find(f"{q('tcPr')}/{q('tcW')}") if tcs else None
            if tcw is not None and (tcw.get(q('type')) or 'dxa') == 'dxa':
                tcw.set(q('w'), str(int(tcw.get(q('w')) or 0) + ecart))
    return True


def _deux_points(root):
    n = 0
    for p in root.iter(q('p')):
        if p.find(f"{q('pPr')}/{q('bidi')}") is None: continue
        # texte propre au paragraphe : on ignore celui des cadres ancrés dedans (codes-barres…)
        propres = [e for e in p.iter(q('t'), q('tab')) if next(a for a in e.iterancestors(q('p'))) is p]
        if not any(e.tag == q('tab') for e in propres): continue   # ligne « arabe <tab> français » seulement
        apres_tab = []                                       # partie française : après la dernière tabulation
        for e in propres:
            if e.tag == q('tab'): apres_tab = []
            elif (e.text or '').strip(): apres_tab.append(e)
        if not apres_tab or apres_tab[-1].text.endswith(LRM): continue
        if LATIN.search(''.join(t.text for t in apres_tab)):
            # finit par « : », « ……… » (valeur vide) ou autre signe neutre : la marque LTR le garde à droite
            apres_tab[-1].text = apres_tab[-1].text.rstrip() + LRM; n += 1
    return n


# Codes-barres CNE / CIN, comme sur le diplôme imprimé : accrochés à la 1re ligne du préambule
# (« Vu la loi… » / « بناء على… »), dans la marge de chaque côté du texte.
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
V = 'urn:schemas-microsoft-com:vml'
DECALAGE_CODES_BARRES = -68000     # EMU (1 mm = 36000) : < 0 remonte les codes-barres, > 0 les descend
DECALAGE_CNE = 0                   # CNE (à droite) par rapport à la CIN : 0 = même niveau (1 mm = 36000)
RAPPORT_CNE = 0.85                 # CNE : longueur x 0,85 (un peu plus court, reste plus long que la CIN)
MODULE_CODES_BARRES = 6500        # largeur d'une barre fine : 0,18 mm -> la longueur suit le nombre de caractères
LONGUEUR_MAX_CODES_BARRES = 1620000  # 4,5 cm au plus (barres affinées si le code est très long)
LONGUEUR_CODES_BARRES = 1260000    # 3,5 cm : longueur si la valeur codée n'est pas connue
LARGEUR_CODES_BARRES = 162000     # 0,45 cm : épaisseur du code-barres (cadre = image, sans marge)
ECART_TEXTE = 36000                # 1 mm entre le code-barres et le bord du texte
EMU_PAR_TWIP = 635


def _premier_paragraphe_preambule(body):
    """1er paragraphe du préambule : 1re cellule du tableau du préambule (doctorat) ou paragraphe « بناء على… » / « Vu… »."""
    for el in body:
        if el.tag == q('tbl'):
            return next(el.iter(q('p')), None)
        if el.tag == q('p'):
            texte = ''.join(t.text or '' for t in el.iter(q('t'))).strip()
            if re.match(r'(Vu |بناء)', texte): return el
    return None


def _page(root):
    s = next(root.iter(q('sectPr')), None)
    sz, mar = (s.find(q('pgSz')), s.find(q('pgMar'))) if s is not None else (None, None)
    if sz is None or mar is None: return None
    g = lambda e, k: int(e.get(q(k)) or 0)
    return g(sz, 'w'), g(mar, 'left'), g(mar, 'right')


def _placer_codes_barres(root):
    body = root.find(q('body')); cible = _premier_paragraphe_preambule(body); page = _page(root)
    if cible is None or page is None: return 0
    larg_page, marge_g, marge_d = (x * EMU_PAR_TWIP for x in page)
    n = 0
    for anc in list(root.iter(f'{{{WP}}}anchor')):
        if not any((d.get('name') or '').startswith('CodeBarre') for d in anc.iter(f'{{{WP}}}docPr')): continue
        ph, pv = anc.find(f'{{{WP}}}positionH'), anc.find(f'{{{WP}}}positionV')
        if ph is None or pv is None or ph.find(f'{{{WP}}}posOffset') is None: continue
        larg = LARGEUR_CODES_BARRES
        valeur = next((d.get('descr', '')[len('Code 39 : '):] for d in anc.iter(f'{{{WP}}}docPr')
                       if (d.get('descr') or '').startswith('Code 39 : ')), '')
        longueur = min(modules_code39(valeur) * MODULE_CODES_BARRES, LONGUEUR_MAX_CODES_BARRES) if valeur else LONGUEUR_CODES_BARRES
        gauche = int(ph.find(f'{{{WP}}}posOffset').text) < 0           # CIN à gauche (décalage négatif), CNE à droite
        decalage = DECALAGE_CODES_BARRES + (0 if gauche else DECALAGE_CNE)
        if not gauche: longueur = int(longueur * RAPPORT_CNE)
        x = marge_g - ECART_TEXTE - larg if gauche else larg_page - marge_d + ECART_TEXTE
        ph.set('relativeFrom', 'page'); ph.find(f'{{{WP}}}posOffset').text = str(x)
        pv.set('relativeFrom', 'paragraph'); pv.find(f'{{{WP}}}posOffset').text = str(decalage)
        anc.set('layoutInCell', '0')                                   # peut sortir de la cellule du tableau (mode Word 12)
        for ext in list(anc.iter(f'{{{WP}}}extent')) + list(anc.iter(f'{{{A}}}ext')):   # cadre et image : même hauteur
            if ext.get('cy'): ext.set('cy', str(longueur))
            if ext.get('cx'): ext.set('cx', str(LARGEUR_CODES_BARRES))       # cadre et image : même largeur
        run = next((a for a in anc.iterancestors(q('r'))), None)
        if run is None: continue
        # ancien format (mc:Fallback, VML) dans le même run : même position
        for sh in list(run.iter(f'{{{V}}}shape')) + list(run.iter(f'{{{V}}}rect')):
            st = sh.get('style') or ''
            st = re.sub(r'margin-left:-?[\d.]+pt', f'margin-left:{x / 12700:.2f}pt', st)
            st = re.sub(r'margin-top:-?[\d.]+pt', f'margin-top:{decalage / 12700:.2f}pt', st)
            st = re.sub(r'(?<![-\w])height:[\d.]+pt', f'height:{longueur / 12700:.2f}pt', st)
            st = re.sub(r'(?<![-\w])width:[\d.]+pt', f'width:{LARGEUR_CODES_BARRES / 12700:.2f}pt', st)
            st = re.sub(r'mso-position-horizontal-relative:\w+', 'mso-position-horizontal-relative:page', st)
            st = re.sub(r'mso-position-vertical-relative:\w+', 'mso-position-vertical-relative:text', st)
            sh.set('style', st)
            for d in sh.iter(f'{{{WP}}}extent'): d.set('cy', str(longueur)); d.set('cx', str(LARGEUR_CODES_BARRES))
            for d in sh.iter(f'{{{A}}}ext'): d.set('cy', str(longueur)); d.set('cx', str(LARGEUR_CODES_BARRES))
        if run.getparent() is not cible:                               # déplacé dans le 1er paragraphe du préambule
            # à sa place, un caractère invisible : l'en-tête garde sa mise en page (saut de colonne qui suit, etc.)
            if any(b.get(q('type')) == 'column' for b in run.getparent().iter(q('br'))):
                vide = etree.Element(q('r')); etree.SubElement(vide, q('t')).text = '​'
                run.addprevious(vide)
            run.getparent().remove(run); cible.append(run)
        n += 1
    return n


# En-tête en 3 colonnes (CIN | titre | CNE) : une fois les codes-barres déplacés dans la marge du préambule,
# les colonnes 1 et 3 ne contiennent plus que des lignes vides. La plus haute fixait pourtant la hauteur de l'en-tête,
# d'où un grand blanc entre « Diplôme de Doctorat » et « Vu la loi… » : ces lignes vides sont réduites à presque rien.
LIGNE_VIDE_ENTETE = 20             # twips (1 pt) : hauteur d'une ligne vide de l'en-tête
ESPACE_TITRE_PREAMBULE = 95        # twips (4 pt) entre « Diplôme de Doctorat » et « Vu la loi… » (200 dans le modèle)
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
INVISIBLE = re.compile(r'^[\s\u200b\u200e\u200f]*$')


def _texte_run(r): return ''.join(t.text or '' for t in r.findall(q('t')))
def _a_image(el): return el.find(f'.//{q("drawing")}') is not None or el.find(f'.//{q("pict")}') is not None or el.find(f'.//{{{MC}}}AlternateContent') is not None


def _espacement(p, **attrs):
    ppr = p.find(q('pPr'))
    if ppr is None: ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    sp = ppr.find(q('spacing'))
    if sp is None:
        sp = etree.Element(q('spacing'))
        suite = next((ppr.find(q(n)) for n in ('ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
                                               'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')
                      if ppr.find(q(n)) is not None), None)
        (suite.addprevious(sp) if suite is not None else ppr.append(sp))
    for k, val in attrs.items(): sp.set(q(k), str(val))


def _compacter_entete(root):
    body = root.find(q('body')); entete, sect = [], None
    for el in body:
        if el.tag != q('p'): return 0
        entete.append(el)
        sect = el.find(f"{q('pPr')}/{q('sectPr')}")
        if sect is not None: break
    cols = sect.find(q('cols')) if sect is not None else None
    if cols is None or int(cols.get(q('num')) or 1) < 2: return 0
    vide = lambda p: not _a_image(p) and all(INVISIBLE.match(_texte_run(r)) for r in p.iter(q('r')))
    n = 0
    for i, p in enumerate(entete):
        runs = p.findall(q('r'))
        k = next((j for j, r in enumerate(runs) if any(b.get(q('type')) == 'column' for b in r.findall(q('br')))), None)
        # « ​ <saut de colonne> titre » : la 1re ligne (invisible) restait dans la colonne précédente avec l'interligne du titre
        if k and i and vide(entete[i - 1]) and all(INVISIBLE.match(_texte_run(r)) and not _a_image(r) for r in runs[:k]):
            for r in runs[:k]: p.remove(r)
            entete[i - 1].append(runs[k])
            _espacement(p, before=0); n += 1
    for p in entete:
        if vide(p): _espacement(p, before=0, after=0, line=LIGNE_VIDE_ENTETE, lineRule='exact'); n += 1
    titre = [p for p in entete if not vide(p)]
    for p in titre: n += _centrer(p)
    if titre: _espacement(titre[-1], after=ESPACE_TITRE_PREAMBULE)
    return n


def _centrer(p):
    """Titre centré par le paragraphe et non par des espaces : « ⎵⎵⎵…⎵شهادة الدكتوراه » débordait de la colonne
    (le titre passait sur 2 lignes ou se décalait selon la police)."""
    ts = [t for r in p.findall(q('r')) for t in r.findall(q('t'))]
    for t in ts:                                                  # espaces en tête
        t.text = (t.text or '').lstrip(' \u00a0\u200b')
        if t.text: break
    for t in reversed(ts):                                        # espaces en fin
        t.text = (t.text or '').rstrip(' \u00a0')
        if t.text: break
    for t in ts: t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    ppr = p.find(q('pPr'))
    if ppr is None: ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    for tab in ppr.findall(q('tabs')): ppr.remove(tab)
    jc = ppr.find(q('jc'))
    if jc is None:
        jc = etree.Element(q('jc'))
        suite = next((ppr.find(q(n)) for n in ('textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId',
                                               'cnfStyle', 'rPr', 'sectPr', 'pPrChange') if ppr.find(q(n)) is not None), None)
        (suite.addprevious(jc) if suite is not None else ppr.append(jc))
    jc.set(q('val'), 'center')
    return 1


def aligner(docx):
    """docx (bytes) -> docx (bytes) aligné. Document inchangé si rien à corriger."""
    zin = zipfile.ZipFile(io.BytesIO(docx))
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    root = etree.fromstring(files['word/document.xml'])
    largeur = _largeur_utile(root)
    body = root.find(q('body'))
    modif = sum(_ajuster_tableau(t, largeur) for t in body.findall(q('tbl'))) if largeur else 0
    modif += _deux_points(root)
    modif += _placer_codes_barres(root)
    modif += _compacter_entete(root)
    if not modif: return docx
    files['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in files.items(): z.writestr(k, v)
    return out.getvalue()

# ------------------------------------------------------------------ accord au féminin
AR = '؀-ۿ'
FEMININ = [                                   # (motif dans le modèle, texte au féminin) ; le modèle est au masculin
    (r'l’étudiant(?!e)', 'l’étudiante'), (r"l'étudiant(?!e)", "l'étudiante"),
    (r'L’étudiant(?!e)', 'L’étudiante'), (r"L'étudiant(?!e)", "L'étudiante"),
    (r'(?<![A-Za-zÀ-ÿ])né à', 'née à'), (r'(?<![A-Za-zÀ-ÿ])Né à', 'Née à'),   # « …1992né à » : chiffre ou arabe avant, ok
    (rf'(?<![{AR}])الطالب(?![{AR}])', 'الطالبة'),          # « للطالب » (الرقم الوطني للطالب) reste inchangé
    (r'المزداد(?!ة)', 'المزدادة'),
    (r'أحرز(?!ت)', 'أحرزت'),
    (rf'(?<![{AR}])السيد(?![{AR}])', 'السيدة'),
]


def accorder(docx, sexe):
    """Met le diplôme au féminin si sexe == 'F' (« l'étudiante », « née à », « الطالبة », « المزدادة », « أحرزت »…).
    Pour 'M' ou un sexe inconnu, le document reste tel quel (modèle au masculin)."""
    if (sexe or '').upper()[:1] != 'F': return docx
    from docfill import replace_in_paragraph
    zin = zipfile.ZipFile(io.BytesIO(docx))
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    root = etree.fromstring(files['word/document.xml'])
    for p in root.iter(q('p')):
        if p.find('.//' + q('p')) is not None: continue           # paragraphe qui contient un cadre : on traite ceux du cadre
        for motif, fem in FEMININ:
            n = 0
            while n < 20 and replace_in_paragraph(p, motif, fem): n += 1
    files['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in files.items(): z.writestr(k, v)
    return out.getvalue()