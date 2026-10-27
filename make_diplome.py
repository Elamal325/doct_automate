"""Crée modeles/doctorat/Diplome.docx (champs {{...}}) à partir du modèle شهادة_الدكتوراه.docx."""
from docfill import make_template
R=[(r'procès-verbal du jury de soutenance en date du\s*:','procès-verbal du jury de soutenance en date du : {{DATE_SOUTENANCE}}'),
 (r'L’étudiant\(e\):\s*…+','L’étudiant(e): {{NOM_COMPLET}}'),(r'…+','{{CIN}}'),(r'…+','{{CNE}}'),
 (r'Etablissement\s+Universitaire\s*:','Etablissement Universitaire : {{ETABLISSEMENT}}'),
 (r'A obtenu le Doctorat','A obtenu le Doctorat en {{DISCIPLINE}}'),(r'Option\s*:','Option : {{SPECIALITE}}'),
 (r'Titre de thèse\s*:','Titre de thèse : {{TITRE}}'),(r'Mention\s*:','Mention : {{MENTION}}'),
 (r'بتاريخ\s*\.+','بتاريخ {{DATE_SOUT_AR}}'),(r'\.{5,}','{{NOM_AR}}'),
 (r'\.{5,}','{{LIEU_NAISSANCE_AR}}'),(r'\.{5,}','{{DATE_NAISSANCE_AR}}'),
 (r'\.{3,}','{{CIN}}'),(r'\.{3,}','{{CNE}}'),(r'المؤسسة\s+الجامعية\s*:','المؤسسة الجامعية: {{ETABLISSEMENT_AR}}'),
 (r'الاداب\s*/\s*الفيزياء','{{DISCIPLINE_AR}}'),(r'تحصص\s*:','تخصص: {{SPECIALITE_AR}}'),(r'…+','{{FORMATION_AR}}'),
 (r'عنوان الأطروحة\s*:','عنوان الأطروحة : {{TITRE_AR}}'),(r'\.{3,}','{{MENTION_AR}}'),
 (r'\.{5,}',' {{LIEU_REDACTION_AR}} '),(r'في\s*$','في {{DATE_SOUT_AR}}'),
 (r'^.*رئيسا\s*$','{{JURY_AR_1}}'),(r'^.*مشرفا\s*$','{{JURY_AR_2}}'),(r'^.*عضوا\s*$','{{JURY_AR_3}}')]
d,left=make_template('/mnt/user-data/uploads/شهادة_الدكتوراه.docx',R)
open('modeles/doctorat/Diplome.docx', 'wb').write(d); print('non appliquées:', left)
