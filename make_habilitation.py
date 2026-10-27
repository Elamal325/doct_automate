"""Brouillon des PV d'habilitation : copie des PV de doctorat avec le titre adapté (à valider avec le modèle officiel)."""
from docfill import process, replace_in_paragraph, P
def fn(root):
    for p in root.iter(P):
        for a,b in (('DOCTORAT','HABILITATION UNIVERSITAIRE'),('Doctorat','Habilitation universitaire')):
            while replace_in_paragraph(p,a,b): pass
for n in range(3):
    open('modeles/habilitation/Page_%d.docx'%n,'wb').write(process('modeles/doctorat/Page_%d.docx'%n,fn))
