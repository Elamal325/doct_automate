"""Codes-barres CNE / CIN insérés en images dans les documents générés.

Principe :
  1. valeurs() met dans {{BAR_CNE}} et {{BAR_CIN}} un simple repère : [[BAR_CNE]], [[BAR_CIN]].
  2. Après fill(), inserer_codes_barres(docx, {'BAR_CNE': cne, 'BAR_CIN': cin}) remplace
     chaque repère par une vraie image Code 39 (PNG) fabriquée à partir de la valeur.

Aucune police « Barcode 3 of 9 » n'est nécessaire : le code-barres s'affiche et
s'imprime sur n'importe quel ordinateur, et se lit avec une douchette standard.
"""
import io, re, zipfile
from lxml import etree
from PIL import Image, ImageDraw

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
V = 'urn:schemas-microsoft-com:vml'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

C39 = {
 '0':'nnnwwnwnn','1':'wnnwnnnnw','2':'nnwwnnnnw','3':'wnwwnnnnn','4':'nnnwwnnnw',
 '5':'wnnwwnnnn','6':'nnwwwnnnn','7':'nnnwnnwnw','8':'wnnwnnwnn','9':'nnwwnnwnn',
 'A':'wnnnnwnnw','B':'nnwnnwnnw','C':'wnwnnwnnn','D':'nnnnwwnnw','E':'wnnnwwnnn',
 'F':'nnwnwwnnn','G':'nnnnnwwnw','H':'wnnnnwwnn','I':'nnwnnwwnn','J':'nnnnwwwnn',
 'K':'wnnnnnnww','L':'nnwnnnnww','M':'wnwnnnnwn','N':'nnnnwnnww','O':'wnnnwnnwn',
 'P':'nnwnwnnwn','Q':'nnnnnnwww','R':'wnnnnnwwn','S':'nnwnnnwwn','T':'nnnnwnwwn',
 'U':'wwnnnnnnw','V':'nwwnnnnnw','W':'wwwnnnnnn','X':'nwnnwnnnw','Y':'wwnnwnnnn',
 'Z':'nwwnwnnnn','-':'nwnnnnwnw','.':'wwnnnnwnn',' ':'nwwnnnwnn','*':'nwnnwnwnn',
 '$':'nwnwnwnnn','/':'nwnwnnnwn','+':'nwnnnwnwn','%':'nnnwnwnwn'}

# Cadres du modèle Diplome.docx : 28,3 pt x 128,1 pt (360045 x 1626870 EMU) -> l'image tient dedans
LARGEUR_EMU, HAUTEUR_EMU = 252000, 1520000          # ~0,7 x 4,2 cm, centré : ~1,5 mm d'écart avec le texte


def repere(cle): return f'[[{cle}]]'                 # « [[BAR_CNE]] », mis par valeurs() à la place du code-barres


def nettoyer(valeur):
    """'j 515 056' -> 'J515056' : seuls les caractères que le Code 39 sait coder."""
    s = '' if valeur is None else str(valeur)
    if re.fullmatch(r'\d+\.0', s): s = s[:-2]                     # CNE lu comme nombre par Excel
    return re.sub(r'[^0-9A-Z\-.$/+%]', '', s.upper())


def modules_code39(valeur):
    """Nombre de modules (barre la plus fine) du code-barres, marges de silence comprises : (n + 2 étoiles) x 16 + 20."""
    return (len(valeur) + 2) * 16 + 20


def code39_png(valeur, longueur_px=2000, epaisseur_px=400):
    """Image PNG verticale (lecture de bas en haut, comme le texte vert270 du modèle)."""
    mods = []
    for ch in '*' + valeur + '*':
        for i, c in enumerate(C39[ch]):
            mods.append((i % 2 == 0, 3 if c == 'w' else 1))
        mods.append((False, 1))
    u = longueur_px / (sum(w for _, w in mods) + 20)              # 10 modules de silence de chaque côté
    img = Image.new('L', (longueur_px, epaisseur_px), 255)
    d = ImageDraw.Draw(img); x = 10 * u
    for barre, w in mods:
        if barre: d.rectangle([round(x), 0, round(x + w * u) - 1, epaisseur_px], fill=0)
        x += w * u
    b = io.BytesIO(); img.rotate(90, expand=True).save(b, 'PNG')
    return b.getvalue()


def _run_image(rid, pid, valeur=''):
    xml = (f'<w:r xmlns:w="{W}"><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">'
           f'<wp:extent cx="{LARGEUR_EMU}" cy="{HAUTEUR_EMU}"/><wp:docPr id="{pid}" name="CodeBarre{pid}" descr="Code 39 : {valeur}"/>'
           '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           f'<pic:nvPicPr><pic:cNvPr id="{pid}" name="codebarre{pid}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
           f'<pic:blipFill><a:blip xmlns:r="{R}" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
           f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{LARGEUR_EMU}" cy="{HAUTEUR_EMU}"/></a:xfrm>'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
           '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    return etree.fromstring(xml)


def _cadre_horizontal(el):
    """L'image est déjà tournée : la zone de texte qui la contient ne doit plus tourner son contenu."""
    for anc in el.iterancestors():
        if anc.tag == f'{{{WPS}}}txbx':
            bp = anc.getparent().find(f'{{{WPS}}}bodyPr')
            if bp is not None and (bp.get('vert') or '').startswith('vert'): bp.set('vert', 'horz')
            return
        if anc.tag == f'{{{V}}}textbox':
            st = anc.get('style') or ''
            anc.set('style', st.replace('layout-flow:vertical', 'layout-flow:horizontal').replace(';mso-layout-flow-alt:bottom-to-top', ''))
            return


def inserer_codes_barres(docx, codes):
    """docx (bytes), codes = {'BAR_CNE': 'J515056', 'BAR_CIN': 'JB123456'}
    -> (bytes, [valeurs encodées]). Chaque repère [[BAR_…]] devient une image ;
    si la valeur est vide, le repère est simplement effacé."""
    codes = {cle: nettoyer(val) for cle, val in codes.items()}
    zin = zipfile.ZipFile(io.BytesIO(docx))
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    if not any(repere(c).encode() in files['word/document.xml'] for c in codes):
        return docx, []
    root = etree.fromstring(files['word/document.xml'])
    rels = etree.fromstring(files['word/_rels/document.xml.rels'])

    images, vues, pid = {}, [], 9000                                 # valeur -> rId
    for run in list(root.iter(f'{{{W}}}r')):
        # texte propre au run seulement : un run qui porte un cadre contient aussi le texte du cadre
        textes = run.findall(f'{{{W}}}t')
        texte = ''.join(t.text or '' for t in textes)
        cle = next((c for c in codes if repere(c) in texte), None)
        if cle is None: continue
        val = codes[cle]
        if not val:                                                   # pas de CNE / CIN : on efface le repère
            for t in textes: t.text = (t.text or '').replace(repere(cle), '')
            continue
        if val not in images:
            n = len(images) + 1
            images[val] = f'rIdCodeBarre{n}'
            files[f'word/media/codebarre{n}.png'] = code39_png(val)
            etree.SubElement(rels, f'{{{PKG}}}Relationship', Id=images[val], Type=IMG, Target=f'media/codebarre{n}.png')
            vues.append(val)
        pid += 1
        img = _run_image(images[val], pid, val)
        p = run.getparent()
        p.replace(run, img)
        _cadre_horizontal(img)
        if p.tag == f'{{{W}}}p':                                      # image centrée dans son cadre
            ppr = p.find(f'{{{W}}}pPr')
            jc = ppr.find(f'{{{W}}}jc') if ppr is not None else None
            if jc is not None: jc.set(f'{{{W}}}val', 'center')

    if images:
        ct = etree.fromstring(files['[Content_Types].xml'])
        if not any((d.get('Extension') or '').lower() == 'png' for d in ct.findall(f'{{{CT}}}Default')):
            ct.insert(0, etree.Element(f'{{{CT}}}Default', Extension='png', ContentType='image/png'))
        files['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)
        files['word/_rels/document.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in files.items(): z.writestr(k, v)
    return out.getvalue(), vues