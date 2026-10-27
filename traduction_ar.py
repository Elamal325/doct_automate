"""Traduction automatique vers l'arabe des champs que l'Excel ne fournit pas en arabe.
Dictionnaires modifiables ; on peut aussi ajouter ses propres entrées dans traductions_ar.json
(format : {"lieux": {"taroudant": "تارودانت"}, "disciplines": {...}, "noms": {...}, "prenoms": {...}, "mentions": {...}}).
Principe : une valeur arabe saisie dans l'Excel n'est jamais écrasée ; si un mot est inconnu, le champ reste vide
(mieux vaut « ……… » qu'un nom propre faux sur un diplôme)."""
import json, os, re, unicodedata

def _sa(s): return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn')
def cle(s):                      # « El Mousadik » -> « elmousadik »
    return re.sub(r'[^a-z0-9]+', '', _sa(s).lower())
def mots(s):
    return [m for m in re.split(r'[^a-z0-9]+', _sa(s).lower()) if m]

LIEUX = {  # mots / villes (clé sans accents, en minuscules)
 'agadir': 'أگادير', 'idaoutanane': 'إداوتنان', 'idaoutanane': 'إداوتنان', 'marrakech': 'مراكش', 'taroudant': 'تارودانت',
 'lagfifat': 'لگفيفات', 'ouledteima': 'أولاد تايمة', 'melloul': 'ملول', 'aitmelloul': 'أيت ملول', 'inezgane': 'إنزكان',
 'meknes': 'مكناس', 'casablanca': 'الدار البيضاء', 'rabat': 'الرباط', 'sale': 'سلا', 'fes': 'فاس', 'tanger': 'طنجة',
 'tetouan': 'تطوان', 'kenitra': 'القنيطرة', 'oujda': 'وجدة', 'laayoune': 'العيون', 'dakhla': 'الداخلة', 'guelmim': 'كلميم',
 'safi': 'آسفي', 'essaouira': 'الصويرة', 'eljadida': 'الجديدة', 'benimellal': 'بني ملال', 'khouribga': 'خريبكة',
 'ouarzazate': 'ورزازات', 'tiznit': 'تيزنيت', 'berrechid': 'برشيد', 'mohammedia': 'المحمدية', 'nador': 'الناظور',
 'settat': 'سطات', 'taza': 'تازة', 'errachidia': 'الرشيدية', 'khemisset': 'الخميسات', 'larache': 'العرائش',
 'ifrane': 'إفران', 'biougra': 'بيوكرى', 'tafraout': 'تافراوت', 'taliouine': 'تالوين', 'tata': 'طاطا', 'zagora': 'زاكورة',
 'chtouka': 'شتوكة', 'aitbaha': 'أيت باها', 'ait': 'أيت', 'oulad': 'أولاد', 'ouled': 'أولاد', 'sidi': 'سيدي',
 'ifni': 'إفني', 'sidiifni': 'سيدي إفني', 'tantan': 'طانطان', 'smara': 'السمارة', 'azrou': 'أزرو', 'khenifra': 'خنيفرة',
 'elkelaa': 'القلعة', 'kelaa': 'القلعة', 'sraghna': 'السراغنة', 'chichaoua': 'شيشاوة', 'alhoceima': 'الحسيمة',
 'berkane': 'بركان', 'taourirt': 'تاوريرت', 'beni': 'بني', 'mellal': 'ملال', 'el': 'ال', 'al': 'ال',
 'maroc': 'المغرب', 'france': 'فرنسا', 'paris': 'باريس', 'algerie': 'الجزائر', 'tunisie': 'تونس', 'senegal': 'السنغال',
 'mauritanie': 'موريتانيا', 'mali': 'مالي', 'cameroun': 'الكاميرون', 'gabon': 'الغابون', 'guinee': 'غينيا',
}
LIEUX_FRASES = {'agadir ida outanane': 'أگادير إداوتنان', 'agadir melloul': 'أگادير ملول', 'agadir melloul taroudant': 'أگادير ملول، تارودانت',
                'ouled teima': 'أولاد تايمة', 'lagfifat taroudant': 'لگفيفات، تارودانت', 'ait melloul': 'أيت ملول',
                'agadir agadir ida outanane': 'أگادير إداوتنان', 'sidi ifni': 'سيدي إفني', 'el jadida': 'الجديدة',
                'beni mellal': 'بني ملال', 'al hoceima': 'الحسيمة', 'laayoune': 'العيون'}

DISCIPLINES = {
 'informatique': 'الإعلاميات', 'informatiques': 'الإعلاميات', 'biologie': 'البيولوجيا', 'chimie': 'الكيمياء', 'physique': 'الفيزياء',
 'mathematiques': 'الرياضيات', 'mathematique': 'الرياضيات', 'geologie': 'الجيولوجيا', 'mathematiques appliquees': 'الرياضيات التطبيقية',
 'mathematiques et informatique': 'الرياضيات والإعلاميات', 'mathematiques, informatique et applications': 'الرياضيات والإعلاميات وتطبيقاتها',
 'informatique et intelligence artificielle': 'الإعلاميات والذكاء الاصطناعي', 'securite informatique': 'الأمن المعلوماتي',
 'sciences de la vie et ressources naturelles': 'علوم الحياة والموارد الطبيعية', 'sciences de la vie': 'علوم الحياة',
 'chimie environnementale et valorisation des dechets agroalimentaires': 'الكيمياء البيئية وتثمين النفايات الصناعية الغذائية',
 'chimie organique': 'الكيمياء العضوية', 'chimie analytique': 'الكيمياء التحليلية', 'physique appliquee': 'الفيزياء التطبيقية',
 'physique de la matiere': 'فيزياء المادة', 'energies renouvelables': 'الطاقات المتجددة', 'geosciences': 'علوم الأرض',
 'sciences de la terre': 'علوم الأرض', 'environnement': 'البيئة', 'ecologie': 'علم البيئة', 'biotechnologie': 'التكنولوجيا الحيوية',
 'biochimie': 'الكيمياء الحيوية', 'microbiologie': 'علم الأحياء الدقيقة', 'statistique': 'الإحصاء', 'sciences des materiaux': 'علوم المواد',
 'mathematiques et applications': 'الرياضيات وتطبيقاتها', 'mathematiques appliquees et informatique': 'الرياضيات التطبيقية والإعلاميات',
 'intelligence artificielle': 'الذكاء الاصطناعي', 'reseaux et telecommunications': 'الشبكات والاتصالات',
 'biologie appliquee': 'البيولوجيا التطبيقية', 'valorisation des ressources naturelles': 'تثمين الموارد الطبيعية',
 'biologie appliquee et valorisation des ressources naturelles': 'البيولوجيا التطبيقية وتثمين الموارد الطبيعية',
 'physique de la matiere et du rayonnement': 'فيزياء المادة والإشعاع', 'chimie organique et environnement': 'الكيمياء العضوية والبيئة',
 'sciences de donnees': 'علوم المعطيات', 'informatique, sciences de donnees et intelligence artificielle': 'الإعلاميات، علوم المعطيات والذكاء الاصطناعي',
 'genie des procedes': 'هندسة الطرائق', 'electronique': 'الإلكترونيات', 'mecanique': 'الميكانيك', 'agronomie': 'الهندسة الزراعية',
 'genie logiciel': 'هندسة البرمجيات', 'informatique, genie logiciel et intelligence artificielle': 'الإعلاميات وهندسة البرمجيات والذكاء الاصطناعي',
 'ecologie et environnement': 'الإيكولوجيا والبيئة', 'analyse fonctionnelle': 'التحليل الدالي', 'analyse numerique': 'التحليل العددي',
 'algebre': 'الجبر', 'probabilites et statistique': 'الاحتمالات والإحصاء', 'recherche operationnelle': 'البحث العملياتي',
 'sciences de la vie et ressources naturelles': 'علوم الحياة والموارد الطبيعية', 'vision par ordinateur': 'الرؤية الحاسوبية',
 'chimie physique': 'الكيمياء الفيزيائية', 'physique theorique': 'الفيزياء النظرية', 'geologie appliquee': 'الجيولوجيا التطبيقية',
}
MENTIONS = {'tres honorable': 'مشرف جدا', 'honorable': 'مشرف', 'tres honorable avec felicitations du jury': 'مشرف جدا مع تهنئة اللجنة',
            'passable': 'مقبول', 'assez bien': 'حسن', 'bien': 'جيد', 'tres bien': 'جيد جدا'}
FORMATIONS = {'sciences et techniques & sciences medicales': 'العلوم والتقنيات والعلوم الطبية',
              'sciences et techniques et sciences medicales': 'العلوم والتقنيات والعلوم الطبية'}

NOMS = {  # noms de famille (clé = cle(nom))
 'elmousadik': 'الموساديق', 'moreira': 'موريرا', 'lagzouli': 'لگزولي', 'eddabra': 'الدبرة', 'elblidi': 'البليدي', 'fleury': 'فلوري',
 'alouani': 'العلواني', 'jebbar': 'جبار', 'boumhandi': 'بومهندي', 'hachimi': 'هاشمي', 'machkour': 'مشكور', 'elansari': 'الأنصاري',
 'elmezouary': 'المزواري', 'amellal': 'أملال', 'amazal': 'أمزال', 'wakrim': 'وكريم', 'taik': 'التايك', 'fakhar': 'فخار',
 'benmoussa': 'بنموسى', 'sadik': 'صديق', 'benahmed': 'بن أحمد', 'amghar': 'أمغار', 'asimi': 'عاصمي', 'asimu': 'عاصمي',
 'darouichi': 'درويشي', 'sadqi': 'صدقي', 'akharraz': 'أخراز', 'elouali': 'الوالي', 'yafia': 'يافع', 'ouaderhman': 'وادرحمان',
 'ouyassine': 'أوياسين', 'chamlal': 'شملال', 'elqorachi': 'القرشي', 'boumazgour': 'بومزكور', 'ajebbar': 'أجبار', 'elharfi': 'الحرفي',
 'chahbi': 'الشهبي', 'elfihri': 'الفهري', 'maragh': 'مراغ', 'korry': 'كري', 'idbraim': 'إدبراهيم', 'guelzim': 'كلزيم',
 'elboushaki': 'البوسحاقي', 'elbahi': 'الباهي', 'charfi': 'شرفي', 'aitbaddi': 'أيت بادي', 'bouazza': 'بوعزة', 'abbassi': 'عباسي',
 'aourir': 'أورير', 'aboualid': 'أبو وليد', 'abououalid': 'أبو وليد', 'taheri': 'الطاهري', 'elmorjani': 'المرجاني',
 'boumandi': 'بومندي', 'alaoui': 'العلوي', 'idrissi': 'الإدريسي', 'elidrissi': 'الإدريسي', 'bennani': 'بناني', 'tazi': 'التازي',
 'elamrani': 'العمراني', 'benali': 'بنعلي', 'elfassi': 'الفاسي', 'elhassani': 'الحسني', 'ouahbi': 'وهبي', 'elghazi': 'الغازي',
 'belmouden': 'بلمودن', 'elalami': 'العلمي', 'amrani': 'عمراني', 'ouali': 'والي', 'zahid': 'زاهد', 'elkhattabi': 'الخطابي',
 'ibnouzahir': 'ابن زهير', 'lamrani': 'لمراني', 'benchekroun': 'بنشقرون', 'berrada': 'برادة', 'chraibi': 'الشرايبي',
}
PRENOMS = {
 'mohamed': 'محمد', 'mohammed': 'محمد', 'mohamad': 'محمد', 'ahmed': 'أحمد', 'abdelhamid': 'عبد الحميد', 'david': 'دافيد', 'rkia': 'رقية',
 'altaf': 'ألطاف', 'yannick': 'يانيك', 'naima': 'نعيمة', 'mustapha': 'مصطفى', 'ali': 'علي', 'hicham': 'هشام', 'fatimaazzahra': 'فاطمة الزهراء',
 'fatima': 'فاطمة', 'hassan': 'حسن', 'rachid': 'رشيد', 'khalid': 'خالد', 'elhassan': 'الحسن', 'elhoucein': 'الحسين', 'omar': 'عمر',
 'abdelhadi': 'عبد الهادي', 'abdellatif': 'عبد اللطيف', 'kamal': 'كمال', 'fouad': 'فؤاد', 'solaymane': 'سليمان', 'soulaiman': 'سليمان',
 'soufiane': 'سفيان', 'ibrahim': 'إبراهيم', 'abdessamad': 'عبد الصمد', 'said': 'سعيد', 'ghita': 'غيثة', 'abdellah': 'عبد الله',
 'jaouad': 'جواد', 'zineelabidine': 'زين العابدين', 'mourad': 'مراد', 'radouane': 'رضوان', 'redouan': 'رضوان', 'tayeb': 'الطيب',
 'hasna': 'حسناء', 'aziz': 'عزيز', 'yassin': 'ياسين', 'ismail': 'إسماعيل', 'younes': 'يونس', 'hamid': 'حميد', 'karim': 'كريم',
 'youssef': 'يوسف', 'abdelaziz': 'عبد العزيز', 'abdelkader': 'عبد القادر', 'abderrahim': 'عبد الرحيم', 'abderrahmane': 'عبد الرحمن',
 'abdelmajid': 'عبد المجيد', 'abdelghani': 'عبد الغني', 'abdelilah': 'عبد الإله', 'abdelouahed': 'عبد الواحد', 'brahim': 'إبراهيم',
 'driss': 'إدريس', 'najib': 'نجيب', 'noureddine': 'نور الدين', 'nabil': 'نبيل', 'samir': 'سمير', 'tarik': 'طارق', 'taoufik': 'توفيق',
 'zakaria': 'زكرياء', 'yahya': 'يحيى', 'ayoub': 'أيوب', 'amine': 'أمين', 'othmane': 'عثمان', 'mehdi': 'مهدي', 'lahcen': 'لحسن',
 'larbi': 'العربي', 'hakim': 'حكيم', 'bilal': 'بلال', 'fouzia': 'فوزية', 'imane': 'إيمان', 'zineb': 'زينب', 'soumia': 'سمية',
 'khadija': 'خديجة', 'samira': 'سميرة', 'latifa': 'لطيفة', 'meryem': 'مريم', 'mariam': 'مريم', 'nadia': 'نادية', 'asmae': 'أسماء',
 'hafsa': 'حفصة', 'houda': 'هدى', 'hanane': 'حنان', 'hajar': 'هاجر', 'ikram': 'إكرام', 'kawtar': 'كوثر', 'loubna': 'لبنى',
 'nezha': 'نزهة', 'rachida': 'رشيدة', 'saida': 'سعيدة', 'sara': 'سارة', 'sanaa': 'سناء', 'siham': 'سهام', 'souad': 'سعاد',
 'wafae': 'وفاء', 'zahra': 'الزهراء', 'aicha': 'عائشة', 'bouchra': 'بشرى', 'malika': 'مليكة', 'mouna': 'منى', 'rajae': 'رجاء',
 'soukaina': 'سكينة', 'yasmine': 'ياسمين', 'lamia': 'لمياء', 'fadwa': 'فدوى', 'ghizlane': 'غزلان', 'hayat': 'حياة', 'najat': 'نجاة',
 'nabila': 'نبيلة', 'sofia': 'صوفيا', 'zakia': 'زكية', 'widad': 'وداد', 'mohcine': 'محسن', 'slimane': 'سليمان', 'reda': 'رضى',
 'abdelhak': 'عبد الحق', 'fouzi': 'فوزي', 'jamal': 'جمال', 'moulay': 'مولاي', 'hafid': 'حفيظ', 'lahoucine': 'الحسين', 'tarek': 'طارق',
 'yassir': 'ياسر', 'anas': 'أنس', 'adil': 'عادل', 'hamza': 'حمزة', 'mounir': 'منير',
}

def _perso():
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'traductions_ar.json')
    if not os.path.exists(p): return {}
    try:
        with open(p, encoding='utf-8') as f: return json.load(f)
    except Exception: return {}

def _dico(nom, base):
    d = dict(base); d.update({cle(k) if nom in ('noms', 'prenoms') else _norm(k): v for k, v in _perso().get(nom, {}).items()}); return d
def _norm(s): return re.sub(r'\s+', ' ', _sa(s).lower().replace('’', "'")).strip(' .')

def lieu(s):
    s = _norm(s)
    if not s: return ''
    L = _dico('lieux', {**LIEUX, **{cle(k): v for k, v in LIEUX_FRASES.items()}})
    if cle(s) in L: return L[cle(s)]
    out, i, w = [], 0, mots(s)
    while i < len(w):                         # plus long groupe de mots connu d'abord (« ouled teima », « ait melloul »)
        for n in (3, 2, 1):
            k = ''.join(w[i:i+n])
            if i + n <= len(w) and k in L: out.append(L[k]); i += n; break
        else: return ''
    res = ' '.join(out)
    return re.sub(r'\bال (?=\S)', 'ال', res)

def _parties(s, d):
    """Cherche la phrase entière, sinon découpe sur « , » et « et » et traduit chaque morceau."""
    s = _norm(s)
    if s in d: return d[s]
    morceaux = [m.strip() for m in re.split(r',|\bet\b|&', s) if m.strip()]
    if len(morceaux) < 2: return ''
    t = [d.get(m) for m in morceaux]
    if not all(t): return ''
    return t[0] + ' ' + ' '.join('و' + x for x in t[1:])

def discipline(s): return _parties(s, _dico('disciplines', DISCIPLINES))
specialite = discipline
def mention(s): return _dico('mentions', MENTIONS).get(_norm(s), '')
def formation(s): return _dico('formations', FORMATIONS).get(_norm(s), '')

def nom_complet(nom, prenom):
    """« EL MOUSADIK » + « Abdelhamid » -> « الموساديق عبد الحميد » ; '' si un des deux est inconnu."""
    N, P = _dico('noms', NOMS), _dico('prenoms', PRENOMS)
    n = N.get(cle(nom))
    if not n:
        w = mots(nom)
        n = N.get(''.join(w[:-1])) if len(w) > 1 else None          # « BEN-AHMED El Hassan » mal découpé
    ps = [P.get(cle(x)) for x in re.split(r'\s+', _sa(prenom).strip()) if x] if prenom else []
    p = P.get(cle(prenom)) or (' '.join(ps) if ps and all(ps) else None)
    return f"{n} {p}" if n and p else ''