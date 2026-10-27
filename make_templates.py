"""Transforme les PV déjà remplis (Page_0/1/2) en modèles avec {{CHAMPS}}."""
import re
from docfill import make_template
U='/mnt/user-data/uploads/Page_%d_PV_DEFAOUI.docx'
TIT=r'Détection des piétons.*?Deep Learning'
AFF=['PES, Faculté des Sciences, Agadir','PES, Faculté des Sciences, Agadir','MCH, École Nationale de Commerce et de Gestion, Marrakech','MCH, École Supérieure de Technologie, Essaouira','MCH, Faculté Polydisciplinaire, Ouarzazate','PES, Faculté des Sciences, Meknès','MC, Faculté des Sciences, Agadir']
COURT=['A. BOULOUZ','L. KOUTTI','H. EL BAHI','A. GUEZZAZ','R. DAKIR','M. EL ANSARI','R. LAHMYED']
NOMS=[('BOULOUZ','Abdellah'),('KOUTTI','Lahcen'),('EL BAHI','Hassan'),('GUEZZAZ','Azidine'),('DAKIR','Rachid'),('EL ANSARI','Mohamed'),('LAHMYED','Redouan')]
ELL=r'[….]{2,}'
def jury():
    r=[]
    for i in range(7):
        r+= [(re.escape(COURT[i]),'{{J%d_COURT}}'%(i+1)),(re.escape(AFF[i]),'{{J%d_AFF}}'%(i+1))]
    return r
R0=[(ELL,'{{NUM_PV}}'),(r'Année : 2026','Année : {{ANNEE}}'),(r'Mme\.?\s+DEFAOUI\s+Mahassine','{{CIVILITE}} {{NOM}} {{PRENOM}}'),
    (r'Samedi 17 Octobre 2026 à 9h00','{{DATE_LONGUE}}'),(r'Amphi CRS, Faculté des Sciences, Agadir','{{LIEU_SOUTENANCE}}'),(r'2025/2026','{{ANNEE_UNIV}}')]
R1=[(ELL,'{{NUM_PV}}'),(r'DEFAOUI Mahassine','{{NOM}} {{PRENOM}}'),(r'18/02/1994 à Meknès','{{DATE_NAISSANCE}} à {{LIEU_NAISSANCE}}'),
    (r'1210051843','{{CNE}}'),(r'SH\d+','{{CIN}}'),(r'Sciences et Techniques & Sciences Médicales','{{FORMATION}}'),(TIT,'{{TITRE}}'),
    (r'Mathématiques et Informatique','{{SPECIALITE}}'),(r'17/10/2026','{{DATE_SOUTENANCE}}'),(r'Amphi CRS, Faculté des Sciences, Agadir','{{LIEU_SOUTENANCE}}')]+jury()
R2=[(ELL,'{{NUM_PV}}'),(r'DEFAOUI Mahassine','{{NOM}} {{PRENOM}}'),(r'18/02/1994 à Meknès','{{DATE_NAISSANCE}} à {{LIEU_NAISSANCE}}'),
    (r'1210051843','{{CNE}}'),(r'SH\d+','{{CIN}}'),(r'Sciences et Techniques & Sciences Médicales','{{FORMATION}}'),(r'Informatique','{{DISCIPLINE}}'),
    (r'Samedi 17 Octobre 2026 à 9h00','{{DATE_LONGUE}}'),(r'Amphi CRS, Faculté des Sciences, Agadir','{{LIEU_SOUTENANCE}}')]+jury()+[
    (TIT,'{{TITRE}}'),(r'Agadir, le 17/10/2026','Agadir, le {{DATE_SOUTENANCE}}'),(r'A\. BOULOUZ','{{J1_COURT}}'),
    (r'Par Mme DEFAOUI Mahassine','Par {{CIVILITE}} {{NOM}} {{PRENOM}}'),(TIT,'{{TITRE}}'),(r'-{8,}','{{SPECIALITE}}')]
for n,v in enumerate([n for pair in NOMS for n in pair]):
    k=n//2+1; R2.append((re.escape(v),'{{J%d_%s}}'%(k,'NOM' if n%2==0 else 'PRENOM')))
for i,r in enumerate([R0,R1,R2]):
    data,left=make_template(U%i,r)
    open('modeles/doctorat/Page_%d.docx'%i,'wb').write(data); print('Page',i,'non appliquées:',left)
