import re, zipfile, io, copy
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
VML = 'urn:schemas-microsoft-com:vml'

T = '{%s}t' % W
P = '{%s}p' % W
R = '{%s}r' % W
BR = '{%s}br' % W
XML = '{http://www.w3.org/XML/1998/namespace}space'

ARABE = re.compile(r'[\u0600-\u06FF]')
LIGNE_EMU = 26 * 12700          # hauteur ajoutee par ligne de jury supplementaire (26 pt)
PARTS = re.compile(r'word/(document|header\d*|footer\d*)\.xml$')


def _w(name):
    return '{%s}%s' % (W, name)


def _ts(p):
    return [t for t in p.iter(T)]


# ---------------------------------------------------------------- mise en forme arabe
def _insert_after(parent, new, after_names):
    """Insere `new` apres le dernier enfant dont le nom figure dans after_names (sinon en tete)."""
    last = None
    for c in parent:
        if c.tag in [_w(n) for n in after_names]:
            last = c
    if last is not None:
        last.addnext(new)
    else:
        parent.insert(0, new)


def _rtl(t):
    """Marque le run contenant `t` comme texte complexe (arabe) : rtl, police cs, bCs, szCs."""
    r = t.getparent()
    if r is None or r.tag != R:
        return
    rpr = r.find(_w('rPr'))
    if rpr is None:
        rpr = etree.Element(_w('rPr'))
        r.insert(0, rpr)
    # taille complexe = taille latine si absente
    sz, szcs = rpr.find(_w('sz')), rpr.find(_w('szCs'))
    if sz is not None and szcs is None:
        e = etree.Element(_w('szCs')); e.set(_w('val'), sz.get(_w('val')))
        sz.addnext(e)
    # gras complexe = gras latin si absent
    b, bcs = rpr.find(_w('b')), rpr.find(_w('bCs'))
    if b is not None and bcs is None:
        e = etree.Element(_w('bCs'))
        if b.get(_w('val')) is not None:
            e.set(_w('val'), b.get(_w('val')))
        b.addnext(e)
    # police complexe = police latine si absente
    rf = rpr.find(_w('rFonts'))
    if rf is not None and rf.get(_w('cs')) is None and rf.get(_w('ascii')):
        rf.set(_w('cs'), rf.get(_w('ascii')))
    # <w:rtl/> (ordre du schema : apres vertAlign, avant cs/em/lang)
    if rpr.find(_w('rtl')) is None:
        e = etree.Element(_w('rtl'))
        suite = [c for c in rpr if c.tag in [_w(n) for n in ('cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath')]]
        if suite:
            suite[0].addprevious(e)
        else:
            rpr.append(e)


# ---------------------------------------------------------------- remplacement de texte
def replace_in_paragraph(p, pattern, new):
    """Remplace la 1re occurrence (regex) en conservant la mise en forme des runs.
    Gere l'arabe (rtl) et les retours a la ligne (\\n -> saut de ligne Word)."""
    ts = _ts(p)
    full = ''.join(t.text or '' for t in ts)
    m = re.search(pattern, full)
    if not m:
        return False
    i, j = m.span()
    pos = 0
    first = True
    for t in ts:
        s = t.text or ''
        a, b = pos, pos + len(s)
        pos = b
        if b <= i or a >= j:
            continue
        lo = max(i, a) - a
        hi = min(j, b) - a
        if first:
            parts = new.split('\n')
            t.text = s[:lo] + parts[0] + (s[hi:] if len(parts) == 1 else '')
            t.set(XML, 'preserve')
            if ARABE.search(new):
                _rtl(t)
            if len(parts) > 1:
                run = t.getparent()
                prev = t
                for k, part in enumerate(parts[1:], 1):
                    br = etree.Element(BR)
                    prev.addnext(br)
                    nt = etree.Element(T)
                    nt.set(XML, 'preserve')
                    nt.text = part + (s[hi:] if k == len(parts) - 1 else '')
                    br.addnext(nt)
                    prev = nt
            first = False
        else:
            t.text = s[:lo] + s[hi:]
            t.set(XML, 'preserve')
    return True


# ---------------------------------------------------------------- zone de texte du jury
def _a_jury(el):
    return any('JURY_AR_3' in (t.text or '') for t in el.iter(T))


def _agrandir_drawingml(root, extras):
    """Zones de texte DrawingML (Word moderne) : augmente la hauteur selon le nombre de lignes ajoutees."""
    ajout = len(extras) * LIGNE_EMU
    deja = set()
    for wsp in root.iter('{%s}wsp' % WPS):
        if not _a_jury(wsp):
            continue
        # etendue de l'ancre / du inline
        anc = wsp.getparent()
        while anc is not None and anc.tag not in ('{%s}anchor' % WP, '{%s}inline' % WP):
            anc = anc.getparent()
        if anc is not None and id(anc) not in deja:
            ext = anc.find('{%s}extent' % WP)
            if ext is not None and ext.get('cy'):
                ext.set('cy', str(int(ext.get('cy')) + ajout))
            deja.add(id(anc))
        # etendue de la forme
        sp = wsp.find('{%s}spPr' % 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
        if sp is None:
            for c in wsp:
                if c.tag.endswith('}spPr'):
                    sp = c
        if sp is not None:
            ext = sp.find('.//{%s}ext' % A)
            if ext is not None and ext.get('cy'):
                ext.set('cy', str(int(ext.get('cy')) + ajout))


def _agrandir_vml(root, extras):
    """Formes VML (ancien format / mc:Fallback)."""
    n = len(extras) + 3
    for sh in root.iter('{%s}rect' % VML, '{%s}shape' % VML):
        if _a_jury(sh) or any(extras[0] in (t.text or '') for t in sh.iter(T)):
            style = sh.get('style')
            if style and re.search(r'(?<![-\w])height:[\d.]+pt', style):
                sh.set('style', re.sub(r'(?<![-\w])height:[\d.]+pt', 'height:%dpt' % max(105, 60 + n * 26), style))


# ---------------------------------------------------------------- traitement du .docx
def process(template, fn):
    zin = zipfile.ZipFile(template)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if PARTS.match(it.filename):
                root = etree.fromstring(data)
                fn(root)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(it, data)
    return out.getvalue()


def fill(template, values, extras=None):
    return process(template, lambda root: _remplir(root, values, extras))


def fill_lettres(template, series):
    """Un seul .docx : le corps du modèle est répété une fois par série de valeurs
    (saut de page avant chaque nouvelle lettre). En-têtes/pieds : remplis avec la 1re série."""
    def fn(root):
        body = root.find(_w('body'))
        if body is None:                                   # en-tête / pied de page
            if series: _remplir(root, series[0])
            return
        sect = body.find(_w('sectPr'))
        modele = [el for el in body if el is not sect]
        for el in modele: body.remove(el)
        for k, values in enumerate(series):
            tmp = etree.Element(_w('body'))
            for el in modele: tmp.append(copy.deepcopy(el))
            _remplir(tmp, values)
            if k:
                p1 = tmp.find(P)
                if p1 is not None:
                    ppr = p1.find(_w('pPr'))
                    if ppr is None:
                        ppr = etree.Element(_w('pPr')); p1.insert(0, ppr)
                    if ppr.find(_w('pageBreakBefore')) is None:
                        e = etree.Element(_w('pageBreakBefore'))
                        suite = [c for c in ppr if c.tag not in (_w('pStyle'), _w('keepNext'), _w('keepLines'))]
                        (suite[0].addprevious(e) if suite else ppr.append(e))
            for el in list(tmp):
                (sect.addprevious(el) if sect is not None else body.append(el))
        for n, dp in enumerate(root.iter('{%s}docPr' % WP), 1):   # identifiants d'images uniques
            dp.set('id', str(n))
    return process(template, fn)


def _remplir(root, values, extras=None):
    if extras:
        # 1) agrandir les zones de texte AVANT de dupliquer les paragraphes
        _agrandir_drawingml(root, extras)
        # 2) dupliquer le paragraphe {{JURY_AR_3}} pour chaque membre supplementaire
        for p in list(root.iter(P)):
            if p.find('.//' + P) is None and '{{JURY_AR_3}}' in ''.join(t.text or '' for t in _ts(p)):
                for line in reversed(extras):
                    q = copy.deepcopy(p)
                    replace_in_paragraph(q, re.escape('{{JURY_AR_3}}'), line)
                    p.addnext(q)
        # 3) formes VML
        _agrandir_vml(root, extras)
    # 3b) paragraphes optionnels : {{?CLE}} -> le paragraphe est supprimé si values[CLE] est vide
    for p in list(root.iter(P)):
        if p.find('.//' + P) is not None:
            continue
        m = re.search(r'\{\{\?(\w+)\}\}', ''.join(t.text or '' for t in _ts(p)))
        if not m:
            continue
        if values.get(m.group(1)):
            while replace_in_paragraph(p, r'\{\{\?\w+\}\}', ''):
                pass
        elif p.getparent() is not None:
            p.getparent().remove(p)
    # 4) remplacer tous les {{CLE}}
    for p in list(root.iter(P)):
        while True:
            full = ''.join(t.text or '' for t in _ts(p))
            m = re.search(r'\{\{(\w+)\}\}', full)
            if not m:
                break
            replace_in_paragraph(p, re.escape(m.group(0)), values.get(m.group(1), ''))


def make_template(src, rules):
    """rules: liste ordonnee (regex, nouveau texte) dans l'ordre du document."""
    left = [(re.sub(r' +', r'\\s+', a.replace('\\ ', ' ')), b) for a, b in rules]

    def fn(root):
        for p in root.iter(P):
            while left and replace_in_paragraph(p, left[0][0], left[0][1]):
                left.pop(0)
    data = process(src, fn)
    return data, [r[0] for r in left]