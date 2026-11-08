"""Crée modeles/<programme>/Attestation_rapporteur.docx (« Attestation de rapporteur », champs {{...}})
à partir de modeles/doctorat/Demande_rapport.docx : on garde le logo de la faculté et la mise en page,
le corps de la lettre est remplacé par le texte de l'attestation.

Lancer :  python make_attestation_rapporteur.py
Une attestation est générée par rapporteur (rôle du jury contenant « Rapporteur ») ; voir attestations_rapporteurs() dans app.py.
Champs propres au modèle (ATR_*) :
  ATR_DEST       « Monsieur le Professeur Abdellah BOULOUZ »
  ATR_GRADE      « Professeur de l'Enseignement Supérieur »          (paragraphe supprimé si vide)
  ATR_ETAB       « Faculté des Sciences, Agadir »                    (paragraphe supprimé si vide)
  ATR_DESIGNE    désigné / désignée
  ATR_NATURE     « de la thèse de Doctorat » / « du dossier d'Habilitation Universitaire »
  ATR_PRESENTE   présentée / présenté
  ATR_SUITE      « , intitulée : » (ou « . » si le titre est absent)
  ATR_TITRE      titre (paragraphe supprimé si vide)
  ATR_PRONOM     Il / Elle
  ATR_INTERESSE  l'intéressé / l'intéressée
"""
import os
from lxml import etree
from docfill import process, _w, P

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'modeles', 'doctorat', 'Demande_rapport.docx')
OUTS = [os.path.join(BASE, 'modeles', p, 'Attestation_rapporteur.docx') for p in ('doctorat', 'habilitation')]
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def para(morceaux=(), jc='both', taille=24, avant=0, apres=0, retrait=0, interligne=360):
    """Paragraphe Arial ; morceaux = [(texte, gras)]."""
    p = etree.Element(P)
    ppr = etree.SubElement(p, _w('pPr'))
    sp = etree.SubElement(ppr, _w('spacing'))
    sp.set(_w('before'), str(avant)); sp.set(_w('after'), str(apres))
    sp.set(_w('line'), str(interligne)); sp.set(_w('lineRule'), 'auto')
    if retrait:
        ind = etree.SubElement(ppr, _w('ind')); ind.set(_w('firstLine'), str(retrait))
    etree.SubElement(ppr, _w('jc')).set(_w('val'), jc)
    for texte, gras in morceaux:
        r = etree.SubElement(p, _w('r'))
        rpr = etree.SubElement(r, _w('rPr'))
        f = etree.SubElement(rpr, _w('rFonts'))
        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'): f.set(_w(a), 'Arial')
        if gras:
            etree.SubElement(rpr, _w('b')); etree.SubElement(rpr, _w('bCs'))
        etree.SubElement(rpr, _w('sz')).set(_w('val'), str(taille))
        etree.SubElement(rpr, _w('szCs')).set(_w('val'), str(taille))
        t = etree.SubElement(r, _w('t')); t.text = texte; t.set(XML_SPACE, 'preserve')
    return p


def corps():
    N, G = False, True
    return [
        para(), para(),
        para([('ATTESTATION', G)], jc='center', taille=36, apres=120),
        para([('DE RAPPORTEUR', G)], jc='center', taille=26),
        para(), para(),
        para([("Le Doyen de la Faculté des Sciences d'Agadir atteste que :", N)], retrait=700, apres=240),
        para([('{{ATR_DEST}}', G)], jc='center'),
        para([('{{?ATR_GRADE}}{{ATR_GRADE}}', N)], jc='center'),
        para([('{{?ATR_ETAB}}{{ATR_ETAB}}', N)], jc='center'),
        para(),
        para([('a été {{ATR_DESIGNE}} rapporteur {{ATR_NATURE}} {{ATR_PRESENTE}} par ', N),
              ('{{CIVILITE}} {{PRENOM}} {{NOM}}', G), ('{{ATR_SUITE}}', N)], apres=120),
        para([('{{?ATR_TITRE}}« {{ATR_TITRE}} »', G)], jc='center', apres=120),
        para([('{{ATR_PRONOM}} a établi un rapport scientifique sur ce travail et a siégé au jury de soutenance, '
               'qui s’est tenue le ', N), ('{{DATE_SOUTENANCE_LONG}}', G), (" à la Faculté des Sciences d'Agadir.", N)],
             retrait=700, apres=120),
        para([('La présente attestation est délivrée à {{ATR_INTERESSE}} pour servir et valoir ce que de droit.', N)],
             retrait=700),
        para(), para(),
        para([('Agadir, le {{DATE_EDITION}}', N)], jc='right'),
        para(),
        para([('Le Doyen de la Faculté des Sciences', G)], jc='right'),
    ]


def construire():
    def fn(root):
        body = root.find(_w('body'))
        if body is None: return
        sect = body.find(_w('sectPr'))
        logo = next(p for p in body.iter(P) if p.find('.//' + _w('drawing')) is not None)
        for el in [e for e in body if e is not sect and e is not logo]: body.remove(el)
        for el in corps(): sect.addprevious(el)
    return process(SRC, fn)


if __name__ == '__main__':
    data = construire()
    for out in OUTS:
        os.makedirs(os.path.dirname(out), exist_ok=True)
        open(out, 'wb').write(data)
        print('créé :', out)
