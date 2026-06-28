# -*- coding: utf-8 -*-
"""Genere Rapport_Deep_Learning.docx (python-docx) a partir des sorties reelles des 3
notebooks executes. Style inspire de rapport_final.pdf (identite EMSI). TOC + listes des
figures et des tableaux en vrais champs Word."""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import os
ASSETS = Path('report_assets')
OUT_DOCX = os.environ.get('OUT_DOCX', 'Rapport_Deep_Learning.docx')
doc = Document()

# ---- Palette (inspiree EMSI) ----
GREEN = RGBColor(0x1E, 0x84, 0x49)
BLUE = RGBColor(0x1F, 0x4E, 0x79)
GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN_HEX = '1E8449'

# ---- Styles ----
normal = doc.styles['Normal']
normal.font.name = 'Calibri'; normal.font.size = Pt(11)
doc.styles['Heading 1'].font.size = Pt(15); doc.styles['Heading 1'].font.color.rgb = BLUE
doc.styles['Heading 2'].font.size = Pt(12.5); doc.styles['Heading 2'].font.color.rgb = GREEN
doc.styles['Heading 3'].font.size = Pt(11.5); doc.styles['Heading 3'].font.color.rgb = GREEN

# ----------------------------------------------------------------------
# Helpers bas niveau
# ----------------------------------------------------------------------
def _field(paragraph, instr, placeholder=''):
    run = paragraph.add_run()
    for kind, attr, val in [('begin', 'w:fldCharType', 'begin')]:
        pass
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = placeholder
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
    for el in (b, it, sep, t, e): run._r.append(el)

def _shade(paragraph, fill):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    pPr.append(shd)

def _shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def _bottom_border(p, color=GREEN_HEX, sz=18):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), '2'); bottom.set(qn('w:color'), color)
    pbdr.append(bottom); pPr.append(pbdr)

def color_bar(fill=GREEN_HEX, height_pt=12):
    p = doc.add_paragraph(); _shade(p, fill); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    r = p.add_run(' '); r.font.size = Pt(height_pt)
    return p

def h1(t): return doc.add_heading(t, level=1)
def h2(t): return doc.add_heading(t, level=2)
def h3(t): return doc.add_heading(t, level=3)

def para(t, bold=False, italic=False, color=None, size=None, align=None):
    p = doc.add_paragraph()
    r = p.add_run(t); r.bold = bold; r.italic = italic
    if color is not None: r.font.color.rgb = color
    if size is not None: r.font.size = Pt(size)
    if align is not None: p.alignment = align
    p.paragraph_format.space_after = Pt(8)
    return p

def bullet(t):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(t); p.paragraph_format.space_after = Pt(2)
    return p

def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# ---- Legendes numerotees (alimentent les listes) ----
def add_caption(text, kind='Figure'):
    p = doc.add_paragraph(style='Caption'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(kind + ' '); run.font.color.rgb = BLUE; run.bold = True
    _field(p, f'SEQ {kind} \\* ARABIC', '0')
    r2 = p.add_run(' — ' + text); r2.font.color.rgb = BLUE; r2.font.size = Pt(9)
    return p

def add_figure(img_name, caption, width=6.0):
    path = ASSETS / img_name
    if not path.exists():
        para(f'[figure manquante : {img_name}]'); return
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Inches(width))
    add_caption(caption, 'Figure')

def add_table(headers, rows, caption=None, font=9.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, htxt in enumerate(headers):
        c = t.rows[0].cells[j]; _shade_cell(c, GREEN_HEX)
        c.paragraphs[0].text = ''
        run = c.paragraphs[0].add_run(str(htxt)); run.bold = True
        run.font.color.rgb = WHITE; run.font.size = Pt(font)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for j, v in enumerate(row):
            if ri % 2 == 1: _shade_cell(cells[j], 'EAF5EE')
            cells[j].paragraphs[0].text = ''
            run = cells[j].paragraphs[0].add_run(str(v)); run.font.size = Pt(font)
    if caption: add_caption(caption, 'Table')
    doc.add_paragraph()
    return t

def add_code(code, fill='EFEFEF'):
    p = doc.add_paragraph(); _shade(p, fill)
    p.paragraph_format.space_after = Pt(6); p.paragraph_format.space_before = Pt(6)
    r = p.add_run(); r.font.name = 'Consolas'; r.font.size = Pt(8.5)
    r._element.rPr.rFonts.set(qn('w:cs'), 'Consolas')
    for i, ln in enumerate(code.strip('\n').split('\n')):
        if i: r.add_break()
        r.add_text(ln)

def add_output(text, label='Sortie reelle du notebook :'):
    if label:
        lp = doc.add_paragraph()
        lr = lp.add_run(label); lr.italic = True; lr.font.size = Pt(9); lr.font.color.rgb = GRAY
        lp.paragraph_format.space_after = Pt(2)
    p = doc.add_paragraph(); _shade(p, 'F4F4F4')
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(); r.font.name = 'Consolas'; r.font.size = Pt(8.5)
    r._element.rPr.rFonts.set(qn('w:cs'), 'Consolas')
    for i, ln in enumerate(text.strip('\n').split('\n')):
        if i: r.add_break()
        r.add_text(ln)

def interpret(text):
    p = doc.add_paragraph(); _shade(p, 'FBF6E9')
    p.paragraph_format.space_after = Pt(8)
    r0 = p.add_run('Interpretation. '); r0.bold = True
    r = p.add_run(text)
    return p

# ----------------------------------------------------------------------
# EN-TETE / PIED DE PAGE (sauf page de garde)
# ----------------------------------------------------------------------
section = doc.sections[0]
section.different_first_page_header_footer = True
hp = section.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hr = hp.add_run('Projet Deep Learning — MLP · CNN · RNN/LSTM/GRU & Seq2Seq')
hr.italic = True; hr.font.size = Pt(8); hr.font.color.rgb = GRAY
fp = section.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = fp.add_run('EMSI Casablanca — Deep Learning — 2025/2026    |    page ')
fr.font.size = Pt(8); fr.font.color.rgb = GRAY
_field(fp, 'PAGE', '1')

# ----------------------------------------------------------------------
# PAGE DE GARDE
# ----------------------------------------------------------------------
color_bar(GREEN_HEX, 10)
doc.add_paragraph()
lg = doc.add_paragraph(); lg.alignment = WD_ALIGN_PARAGRAPH.CENTER
lg.paragraph_format.space_after = Pt(4)
lg.add_run().add_picture(str(ASSETS / 'logo_emsi.png'), width=Inches(1.9))
para('Ecole Marocaine des Sciences de l\'Ingenieur — EMSI Casablanca', bold=True,
     color=GREEN, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
para('Filiere : Informatique', italic=True, color=GRAY, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
pmod = doc.add_paragraph(); pmod.alignment = WD_ALIGN_PARAGRAPH.CENTER
rmod = pmod.add_run('Module : Deep Learning'); rmod.font.size = Pt(11); rmod.font.color.rgb = GRAY
_bottom_border(pmod)
doc.add_paragraph()
para('Projet de fin de module', bold=True, color=GRAY, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
para('Deep Learning : MLP, CNN et RNN / LSTM / GRU & Seq2Seq', bold=True, color=GREEN,
     size=20, align=WD_ALIGN_PARAGRAPH.CENTER)
para('Conception, implementation, comparaison et analyse critique de modeles de deep learning '
     'pour donnees tabulaires, images et sequences', italic=True, size=11.5,
     align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
pr = doc.add_paragraph(); pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
pr.add_run('Realise par : ').bold = True
pr.add_run('HABACHI Anass').font.size = Pt(12)
pe = doc.add_paragraph(); pe.alignment = WD_ALIGN_PARAGRAPH.CENTER
pe.add_run('Encadre par : ').bold = True
pe.add_run('Pr. Zineb Hdila').font.size = Pt(12)
doc.add_paragraph()
para('Annee universitaire : 2025 / 2026', bold=True, color=GREEN, size=12,
     align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
color_bar(GREEN_HEX, 10)
# saut de section -> la page de garde occupe exactement une page
sb = doc.add_paragraph(); sb.add_run().add_break(WD_BREAK.PAGE)

# ----------------------------------------------------------------------
# RESUME
# ----------------------------------------------------------------------
h1('Resume')
para("Ce rapport presente la conception, l'implementation et l'analyse critique de trois familles "
     "de modeles de deep learning sous PyTorch, chacune adaptee a une structure de donnees : un "
     "perceptron multicouche (MLP) pour des donnees tabulaires (Adult Income), un reseau "
     "convolutionnel (CNN) pour des images (CIFAR-10), et des reseaux recurrents (RNN, LSTM, GRU) "
     "ainsi qu'un systeme encodeur-decodeur (Seq2Seq) pour des sequences (modele de langage sur "
     "The Time Machine et traduction anglais-francais). Resultats reels obtenus : MLP — exactitude "
     "de test 0,857 (F1 0,671) ; CNN — 0,696 pour le modele de reference, porte a 0,839 par un CNN "
     "ameliore (+14,3 points) contre 0,520 pour un MLP equivalent ; sequences — perplexite finale "
     "LSTM 2,78 < GRU 2,99 < RNN 3,46, et systeme Seq2Seq atteignant un BLEU de 0,45 sur "
     "l'entrainement (beam search) mais 0,03 en test, illustrant le surapprentissage d'un petit "
     "reseau sans attention. Tous les chiffres, tableaux et figures sont extraits des notebooks "
     "reellement executes de bout en bout ; aucune valeur n'est inventee.")
page_break()

# ----------------------------------------------------------------------
# ABREVIATIONS
# ----------------------------------------------------------------------
h1('Abreviations')
add_table(['Sigle', 'Signification'],
    [['MLP', 'Multi-Layer Perceptron (perceptron multicouche)'],
     ['CNN', 'Convolutional Neural Network (reseau convolutionnel)'],
     ['RNN', 'Recurrent Neural Network (reseau recurrent)'],
     ['LSTM', 'Long Short-Term Memory'],
     ['GRU', 'Gated Recurrent Unit'],
     ['Seq2Seq', 'Sequence-to-Sequence (encodeur-decodeur)'],
     ['BPTT', 'Backpropagation Through Time (retropropagation dans le temps)'],
     ['ReLU', 'Rectified Linear Unit (fonction d\'activation)'],
     ['PPL', 'Perplexite (metrique des modeles de langage)'],
     ['BLEU', 'Bilingual Evaluation Understudy (qualite de traduction)'],
     ['OHE', 'One-Hot Encoding (encodage disjonctif)'],
     ['F1', 'Moyenne harmonique precision / rappel'],
     ['SGD', 'Stochastic Gradient Descent']],
    caption='Liste des abreviations.')
page_break()

# ----------------------------------------------------------------------
# TABLE DES MATIERES + LISTE DES FIGURES + LISTE DES TABLEAUX
# ----------------------------------------------------------------------
def index_field(title, instr):
    h1(title)
    p = doc.add_paragraph()
    _field(p, instr, '  [ Mettre a jour : Ctrl+A puis F9 ]')

index_field('Table des matieres', 'TOC \\o "1-3" \\h \\z \\u')
page_break()
index_field('Liste des figures', 'TOC \\h \\z \\c "Figure"')
page_break()
index_field('Liste des tableaux', 'TOC \\h \\z \\c "Table"')
page_break()

# ----------------------------------------------------------------------
# 1. INTRODUCTION
# ----------------------------------------------------------------------
h1('Introduction et mise en contexte')
para("L'objectif de ce projet est de montrer comment le choix d'une architecture de deep learning "
     "encode un biais inductif adapte a la nature des donnees, et d'en mesurer experimentalement les "
     "consequences. Trois cadres complementaires sont traites : les donnees tabulaires (MLP), les "
     "images (CNN) et les sequences (RNN/LSTM/GRU et Seq2Seq). Chaque partie suit la meme demarche "
     "scientifique : rappel theorique, preparation rigoureuse des donnees, implementation PyTorch, "
     "etude experimentale comparative et analyse critique des resultats.")
para('Bonnes pratiques methodologiques appliquees :', bold=True)
bullet("graines aleatoires fixees (reproductibilite) ;")
bullet("separation stricte apprentissage / validation / test ;")
bullet("normalisation ajustee sur le seul ensemble d'apprentissage (anti-fuite) ;")
bullet("bascule model.train() / model.eval() et torch.no_grad() en validation et test ;")
bullet("sauvegarde du meilleur modele (state_dict) et chaque figure enregistree via plt.savefig().")

h2('Environnement technique')
add_table(['Element', 'Valeur'],
    [['Langage', 'Python 3.12'],
     ['Framework', 'PyTorch 2.12 (execution CPU)'],
     ['Device', 'selection automatique cuda / mps / cpu (ici : cpu)'],
     ['Bibliotheques', 'numpy, pandas, scikit-learn, matplotlib'],
     ['Partie I — dataset', 'Adult Income (UCI) — 48 842 exemples, 14 variables mixtes'],
     ['Partie II — dataset', 'CIFAR-10 — 60 000 images couleur 32x32, 10 classes'],
     ['Partie III — corpus', 'The Time Machine (langage) + fra-eng / Tatoeba (traduction)'],
     ['Reproductibilite', 'graine = 42 ; splits et initialisations deterministes']],
    caption='Environnement technique du projet.')
page_break()

# ======================================================================
# 2. PARTIE I — MLP (Adult Income)
# ======================================================================
h1('1. Partie I — MLP et classification tabulaire')
para("Theme : classification supervisee binaire sur le dataset reel Adult Income (UCI). On predit "
     "si le revenu annuel d'un individu depasse 50 000 dollars a partir de 14 variables mixtes "
     "(numeriques et categorielles). Cette partie demontre la maitrise de l'ingenierie PyTorch sur "
     "donnees tabulaires.")

h2('1.1 Fondements theoriques')
para("Un MLP empile des couches entierement connectees (nn.Linear) separees par des non-linearites "
     "(ReLU). La composition de transformations affines et de non-linearites approche des frontieres "
     "de decision complexes. Concepts PyTorch mobilises :")
bullet("nn.Module : classe de base encapsulant parametres et methode forward (propagation avant) ;")
bullet("parametres : tenseurs appris (poids W, biais b), exposes par named_parameters() ;")
bullet("gradient : calcule par autograd lors de loss.backward() (retropropagation) ;")
bullet("state_dict : dictionnaire serialisable des parametres (sauvegarde/rechargement) ;")
bullet("device : CPU ou GPU sur lequel residents tenseurs et modele, qui doivent etre coherents.")

h2('1.2 Chargement et apercu du dataset')
para("Le dataset comporte 32 561 exemples d'apprentissage et 16 281 de test, chacun decrit par 14 "
     "variables plus la cible. Les fichiers sont charges localement (execution reproductible).")
add_output(
"..\\data\\adult\\adult.data - 3974305 octets\n"
"..\\data\\adult\\adult.test - 2003153 octets\n"
"adult.data : (32561, 15)\n"
"adult.test : (16281, 15)", label='Sortie reelle (chargement) :')
para("Le tableau ci-dessous decrit le type des variables (6 numeriques, 8 categorielles, 1 cible).")
add_table(['Variable', 'Type'],
    [['age, fnlwgt, education_num', 'numerique'],
     ['capital_gain, capital_loss, hours_per_week', 'numerique'],
     ['workclass, education, marital_status, occupation', 'categorielle'],
     ['relationship, race, sex, native_country', 'categorielle'],
     ['income', 'cible (binaire <=50K / >50K)']],
    caption='Dictionnaire (resume) des variables du dataset Adult Income.')

h2('1.3 Qualite des donnees et desequilibre des classes')
para("Trois variables categorielles contiennent des valeurs manquantes (codees \"?\"), qui seront "
     "imputees. Surtout, la cible est fortement desequilibree — point determinant pour l'analyse.")
add_output(
"--- Valeurs manquantes (train + test) ---\n"
"workclass        2799\n"
"occupation       2809\n"
"native_country    857\n\n"
"--- Distribution de la cible (apprentissage) ---\n"
"<=50K   24720   (75.9 %)\n"
">50K     7841   (24.1 %)", label='Sortie reelle (qualite des donnees) :')
add_table(['Classe', 'Apprentissage', 'Proportion', 'Test', 'Proportion'],
    [['<=50K', '24 720', '75,9 %', '12 435', '76,4 %'],
     ['>50K', '7 841', '24,1 %', '3 846', '23,6 %']],
    caption='Distribution de la variable cible (desequilibre 76 / 24).')
add_figure('p1_fig1.png', "Exploration : distribution de la cible et de variables explicatives.", 6.2)
add_figure('p1_fig2.png', "Exploration : distributions numeriques et categorielles d'Adult Income.", 6.2)
interpret("la classe majoritaire (<=50K) represente ~76 % des exemples. Ce desequilibre, visible des "
     "l'exploration, conditionnera la lecture de la matrice de confusion et l'ecart de rappel entre "
     "les deux classes.")

h2('1.4 Preparation des donnees')
para("Les variables numeriques sont imputees (mediane) puis standardisees (StandardScaler ajuste "
     "sur le seul ensemble d'apprentissage, anti-fuite) ; les categorielles sont imputees (modalite "
     "la plus frequente) puis encodees en one-hot. Apres encodage, la dimension d'entree passe a 105. "
     "L'apprentissage est divise en train (27 676) et validation (4 885), le test officiel (16 281) "
     "restant intact.")
add_output(
"--- Resultat du decoupage ---\n"
"Apprentissage : 27676 lignes   Validation : 4885 lignes   Test : 16281 lignes\n"
"Dimension d'entree apres preprocessing : 105", label='Sortie reelle (preparation) :')

h2('1.5 Deux implementations du MLP')
para("Deux versions equivalentes sont construites : une avec nn.Sequential, une sous forme de classe "
     "personnalisee (nn.Module) offrant un controle fin du forward et un Dropout. Topologie commune "
     "105 -> 128 -> 64 -> 2.")
add_code(
"def build_sequential_mlp(in_features, hidden1, hidden2, num_classes=2):\n"
"    return nn.Sequential(\n"
"        nn.Linear(in_features, hidden1), nn.ReLU(),\n"
"        nn.Linear(hidden1, hidden2), nn.ReLU(),\n"
"        nn.Linear(hidden2, num_classes))\n\n"
"class AdultIncomeMLP(nn.Module):\n"
"    def __init__(self, in_features, hidden1=128, hidden2=64, num_classes=2, dropout=0.30):\n"
"        super().__init__()\n"
"        self.fc1 = nn.Linear(in_features, hidden1)\n"
"        self.fc2 = nn.Linear(hidden1, hidden2)\n"
"        self.fc3 = nn.Linear(hidden2, num_classes)\n"
"        self.dropout = nn.Dropout(p=dropout)\n"
"    def forward(self, x):\n"
"        x = self.dropout(F.relu(self.fc1(x)))\n"
"        x = self.dropout(F.relu(self.fc2(x)))\n"
"        return self.fc3(x)")
add_output(
"Sequential( (0): Linear(105, 128)  (1): ReLU()  (2): Linear(128, 64)\n"
"            (3): ReLU()  (4): Linear(64, 2) )\n"
"AdultIncomeMLP( (fc1): Linear(105, 128)  (fc2): Linear(128, 64)\n"
"                (fc3): Linear(64, 2)  (dropout): Dropout(p=0.3) )",
    label='Sortie reelle (architectures) :')

h2('1.6 Inspection des parametres')
para("L'inspection via named_parameters() et state_dict() confirme la structure et denombre 21 954 "
     "parametres apprenables.")
add_output(
"nom        shape       numel\n"
"fc1.weight (128, 105)  13440\n"
"fc1.bias   (128,)        128\n"
"fc2.weight (64, 128)    8192\n"
"fc2.bias   (64,)          64\n"
"fc3.weight (2, 64)       128\n"
"fc3.bias   (2,)            2\n"
"Total parametres apprenables : 21954", label='Sortie reelle (named_parameters / state_dict) :')

h2('1.7 Comparaison de trois initialisations')
para("Trois strategies d'initialisation sont comparees a architecture et hyperparametres identiques : "
     "gaussienne (N(0, 0.01)), constante (poids a 1) et Xavier (xavier_uniform_). La constante est "
     "incluse pour illustrer le probleme de symetrie : tous les neurones d'une couche calculent la "
     "meme chose et recoivent le meme gradient.")
add_table(['Initialisation', 'Fonction', 'Principe'],
    [['Gaussienne', 'init_gaussian', 'poids ~ N(0, 0.01), biais = 0'],
     ['Constante', 'init_constant', 'poids = 1, biais = 0'],
     ['Xavier', 'init_xavier', 'xavier_uniform_, biais = 0']],
    caption='Strategies d\'initialisation comparees.')
para("L'effet est net sur les premieres epoques : la constante demarre avec une perte aberrante "
     "(7,06) et converge lentement, alors que Xavier demarre deja vers 0,37.")
add_output(
"Entrainement : Custom_Constante\n"
"Epoch 001/35 | train_loss=7.0574 train_acc=0.6010 | val_acc=0.7593\n"
"Epoch 010/35 | train_loss=0.6075 train_acc=0.7358 | val_acc=0.7112\n"
"Epoch 035/35 | train_loss=0.3362 train_acc=0.8478 | val_acc=0.8553\n\n"
"Entrainement : Custom_Xavier\n"
"Epoch 001/35 | train_loss=0.3699 train_acc=0.8265 | val_acc=0.8606\n"
"Epoch 035/35 | train_loss=0.2760 train_acc=0.8731 | val_acc=0.8631",
    label='Sortie reelle (logs d\'entrainement) :')
add_figure('p1_fig3.png', "Courbes de perte et d'exactitude (apprentissage vs validation) des 4 experiences MLP.", 6.4)
add_table(['Experience', 'Ep. opt.', 'Acc. val', 'Acc. test', 'Precision', 'Rappel', 'F1'],
    [['Custom_Xavier', '9', '0,8663', '0,8570', '0,735', '0,617', '0,671'],
     ['Custom_Gaussienne', '14', '0,8657', '0,8562', '0,744', '0,596', '0,662'],
     ['Sequential_Xavier', '3', '0,8653', '0,8552', '0,733', '0,609', '0,665'],
     ['Custom_Constante', '31', '0,8563', '0,8524', '0,731', '0,593', '0,655']],
    caption='Performances comparees des quatre experiences (jeu de test officiel).')
interpret("Xavier offre le meilleur compromis (test 0,857) en convergeant le plus vite (meilleure "
     "epoque des la 9e). La constante, bien qu'elle finisse par converger grace au Dropout et a Adam "
     "qui brisent partiellement la symetrie, reste la moins performante et la plus lente (epoque 31) "
     "— illustration concrete du probleme de symetrie.")

h2('1.8 Sauvegarde, rechargement et evaluation finale')
para("Le meilleur modele (Custom_Xavier) est sauvegarde via torch.save(state_dict) puis recharge et "
     "reevalue, confirmant une exactitude de test identique.")
add_output(
"Modele sauvegarde : ..\\outputs\\mlp\\best_mlp_adult_income_final.params\n"
"Rechargement reussi.  Loss test : 0.3091   Accuracy test : 0.8569",
    label='Sortie reelle (sauvegarde / rechargement) :')
add_output(
"Accuracy : 0.85695   Precision : 0.734757   Recall : 0.617265   F1 : 0.670906\n\n"
"              precision    recall  f1-score   support\n"
"     <=50K       0.89      0.93      0.91     12435\n"
"      >50K       0.73      0.62      0.67      3846\n"
"  accuracy                           0.86     16281\n"
" macro avg       0.81      0.77      0.79     16281\n"
"weighted avg     0.85      0.86      0.85     16281", label='Sortie reelle (evaluation finale) :')
add_figure('p1_fig4.png', "Matrice de confusion du meilleur MLP sur le jeu de test officiel (16 281 exemples).", 4.6)
interpret("l'exactitude globale de 86 % masque un comportement asymetrique : le modele excelle sur "
     "la classe majoritaire <=50K (rappel 0,93) mais manque ~38 % des hauts revenus (rappel 0,62). "
     "C'est la consequence directe du desequilibre 76/24 ; le F1 ponderee (0,85) et le F1 de la "
     "classe >50K (0,67) sont des indicateurs plus honnetes que la seule exactitude.")

h2('1.9 Verification de conformite (Partie I)')
add_table(['Exigence du sujet', 'Couverte'],
    [['Concepts (nn.Module, gradient, state_dict, device)', 'Oui (1.1)'],
     ['Preparation (nettoyage, encodage, normalisation, splits)', 'Oui (1.2-1.4)'],
     ['Deux MLP (Sequential + classe)', 'Oui (1.5)'],
     ['named_parameters() et state_dict()', 'Oui (1.6)'],
     ['Trois initialisations (gaussienne/constante/Xavier)', 'Oui (1.7)'],
     ['Sauvegarde puis rechargement', 'Oui (1.8)'],
     ['Execution sur device + coherence', 'Oui (1.1, 1.8)'],
     ['Metriques (accuracy, precision, recall, F1, confusion)', 'Oui (1.8)']],
    caption='Conformite de la Partie I au cahier des charges.')

h2('1.10 Question de synthese — Partie I')
para("Dans quelle mesure un MLP bien parametre est-il pertinent pour la classification tabulaire, et "
     "quelles sont ses limites au regard de la structure statistique des donnees ?", bold=True, italic=True)
para("Theoriquement, un MLP est pertinent car le probleme admet une frontiere non lineaire : "
     "l'alternance Linear/ReLU approche des interactions complexes, et la preparation (standardisation, "
     "one-hot) produit une representation vectorielle homogene de dimension 105. Methodologiquement, le "
     "protocole (split train/val/test, preprocessing ajuste sur le train, selection sur la validation, "
     "evaluation sur un test intact) garantit des resultats non biaises ; la comparaison des "
     "initialisations demontre l'importance d'un bon point de depart. Cote resultats, le modele atteint "
     "85,7 % d'exactitude, mais la matrice de confusion revele la limite principale : le desequilibre "
     "degrade le rappel des hauts revenus (0,62). Au-dela, un MLP ne modelise pas nativement la structure "
     "des donnees tabulaires (interactions, variables ordinales) ; des methodes a arbres (Random Forest, "
     "XGBoost) y sont souvent plus competitives. Pistes : ponderation de la perte par classe, "
     "reequilibrage, variables croisees.")
page_break()

# ======================================================================
# 3. PARTIE II — CNN (CIFAR-10)
# ======================================================================
h1('2. Partie II — CNN et vision par ordinateur')
para("Theme : classification d'images sur CIFAR-10 (60 000 images couleur 32x32, 10 classes : avion, "
     "automobile, oiseau, chat, cerf, chien, grenouille, cheval, bateau, camion). Cette partie analyse "
     "les operations convolutionnelles et demontre la superiorite des CNN sur les MLP pour les donnees "
     "a structure spatiale.")
add_figure('p2_fig1.png', "Echantillon d'images CIFAR-10 avec leurs etiquettes.", 5.5)

h2('2.1 Pourquoi un MLP echoue sur les images')
para("Une image 32x32x3 represente 3 072 valeurs. Un MLP les aplatit, detruisant la structure spatiale "
     "et exigeant enormement de parametres, sans aucune invariance par translation. Les CNN repondent "
     "par trois idees fondatrices : la localite (chaque neurone regarde une region restreinte), le "
     "partage des poids (un meme filtre glisse sur toute l'image) et la hierarchie des representations "
     "(aretes -> textures -> objets).")

h2('2.2 Calculs manuels et verification contre PyTorch')
para("La correlation croisee 2D, le max-pooling et l'average-pooling sont d'abord implementes a la main "
     "(doubles boucles Python) puis valides contre les couches PyTorch. La formule de taille de sortie "
     "est H_out = (H + 2p - k) / s + 1.")
add_output(
"--- Correlation croisee 2D ---       --- nn.Conv2d ---\n"
"maison : [[19, 25],                  PyTorch : [[19, 25],\n"
"          [37, 43]]                            [37, 43]]\n"
"Equivalence validee : True\n\n"
"--- Max-pooling ---  maison [[4,5],[7,8]]  ==  nn.MaxPool2d [[4,5],[7,8]]\n"
"--- Average-pooling --- maison [[2,3],[5,6]] == nn.AvgPool2d [[2,3],[5,6]]  -> True\n"
"--- Conv padding=1, stride=2 --- maison [[0,8],[21,43]] == nn.Conv2d -> True\n"
"--- Conv multi-canaux --- maison [[1035,1101],[1233,1299]] == nn.Conv2d -> True",
    label='Sortie reelle (implementations maison vs PyTorch) :')
add_table(['Operation', 'H_in', 'k', 'p', 's', 'Formule', 'PyTorch', 'OK'],
    [['Conv 5x5 (LeNet C1)', '32', '5', '0', '1', '28', '28', 'OUI'],
     ['Conv 3x3 (same)', '32', '3', '1', '1', '32', '32', 'OUI'],
     ['Conv 3x3 (downsample)', '32', '3', '1', '2', '16', '16', 'OUI'],
     ['MaxPool 2x2 (apres C1)', '28', '2', '0', '2', '14', '14', 'OUI']],
    caption='Verification des calculs dimensionnels manuels contre PyTorch.', font=9)
interpret("l'equivalence (correlation, max/avg-pooling, padding+stride, multi-canaux) confirme que "
     "nn.Conv2d calcule bien une correlation croisee et que nos formules dimensionnelles sont exactes ; "
     "les couches PyTorch ne different que par leur vectorisation optimisee.")

h2('2.3 Architecture et entrainement du CNN de reference')
para("Un CNN de type LeNet (deux blocs convolution + pooling, puis couches denses) est implemente. Le "
     "modele de reference compte 53 770 parametres, contre 62 006 pour un LeNet-5 classique. Il est "
     "entraine 12 epoques sur 45 000 images (5 000 en validation).")
add_output(
"LeNet5 : 62 006 parametres   |   ConfigurableCNN (defaut) : 53 770 parametres\n"
"Train: 45000 | Validation: 5000 | Test: 10000\n\n"
"CNN final : 53 770 parametres\n"
"Epoch  1/12 | loss 1.589 | train_acc 0.421 | val_acc 0.516\n"
"Epoch  6/12 | loss 1.094 | train_acc 0.608 | val_acc 0.634\n"
"Epoch 12/12 | loss 0.943 | train_acc 0.665 | val_acc 0.692", label='Sortie reelle (entrainement CNN) :')
add_figure('p2_fig2.png', "Courbes d'apprentissage du CNN de reference (perte et exactitude).", 6.2)
add_output(
"Accuracy : 0.6963   Precision (macro) : 0.6946   Recall : 0.6963   F1 : 0.6930\n\n"
"              precision  recall  f1-score  support\n"
"    airplane     0.81     0.67     0.73     1000\n"
"  automobile     0.77     0.85     0.81     1000\n"
"        bird     0.62     0.48     0.54     1000\n"
"         cat     0.52     0.51     0.51     1000\n"
"        ship     0.79     0.81     0.80     1000", label='Sortie reelle (evaluation test du CNN) :')
add_figure('p2_fig3.png', "Matrice de confusion du CNN de reference sur CIFAR-10.", 4.8)
interpret("le CNN atteint 69,6 % d'exactitude de test. Les classes d'animaux (chat 0,51, oiseau 0,48) "
     "sont nettement plus difficiles que les objets manufactures (automobile 0,85, bateau 0,81), "
     "confusion attendue car ces categories partagent des textures et formes proches.")

h2('2.4 Etude experimentale (ablation)')
para("Sept configurations sont comparees pour isoler l'influence de chaque choix (padding, stride, type "
     "de pooling, nombre de filtres, convolution 1x1). Chaque variante est entrainee sur peu d'epoques ; "
     "les valeurs absolues sont donc bruitees, mais les tendances relatives restent informatives.")
add_table(['Configuration', 'Acc. test', 'Parametres', 'Temps (s)'],
    [['Sans padding (valid)', '0,640', '53 770', '307'],
     ['Beaucoup de filtres (64)', '0,623', '142 986', '641'],
     ['Avec convolution 1x1', '0,622', '59 178', '506'],
     ['Stride=2 (conv)', '0,622', '53 770', '134'],
     ['Pooling moyen', '0,612', '53 770', '321'],
     ['Baseline (32f, 3x3, same, max)', '0,606', '53 770', '440'],
     ['Peu de filtres (16)', '0,547', '22 986', '227']],
    caption='Comparaison des variantes architecturales du CNN (etude ablative).')
add_figure('p2_fig4.png', "Influence des choix architecturaux sur l'exactitude de test.", 6.0)
interpret("le nombre de filtres est le facteur le plus determinant (32->16 fait chuter de 0,606 a "
     "0,547). Le stride=2 divise le temps par trois pour une exactitude comparable. La convolution 1x1 "
     "et le max-pooling apportent des gains modestes. Ces tendances sont coherentes avec la theorie.")

h2('2.5 Visualisation des cartes de caracteristiques')
para("A l'aide de register_forward_hook, on capture les activations internes : 32 cartes 32x32 pour la "
     "premiere convolution, 64 cartes 16x16 pour la seconde.")
add_output(
"Forme des activations conv1 : (1, 32, 32, 32)\n"
"Forme des activations conv2 : (1, 64, 16, 16)", label='Sortie reelle (cartes de caracteristiques) :')
add_figure('p2_fig5.png', "Cartes de caracteristiques des deux couches convolutionnelles.", 6.4)
interpret("les cartes de conv1 reagissent a des motifs simples et locaux (aretes, transitions de "
     "couleur) ; celles de conv2, plus petites et plus abstraites, s'activent sur des combinaisons de "
     "motifs. On visualise concretement la hierarchie des representations.")

h2('2.6 Comparaison MLP vs CNN et modele ameliore')
para("Sur le meme jeu d'images, un MLP simple (1,7 M de parametres) plafonne a 52,0 % tandis que le CNN "
     "(53 770 parametres) atteint 69,6 %. Avec 32 fois moins de parametres, le CNN fait nettement mieux.")
add_output(
"MLP simple : 1 707 274 parametres\n"
"Epoch 12/12 | loss 1.120 | train_acc 0.602 | val_acc 0.522\n\n"
"Modele        Parametres   Test Acc\n"
"MLP simple     1 707 274     0.5202\n"
"CNN (final)       53 770     0.6963", label='Sortie reelle (MLP vs CNN) :')
add_figure('p2_fig6.png', "Comparaison directe MLP simple vs CNN sur CIFAR-10.", 5.5)
para("Un CNN ameliore (reseau plus profond 64->128->256, BatchNorm2d, Dropout, weight decay, "
     "augmentation de donnees et scheduler cosine) est ensuite entraine 22 epoques.")
add_output(
"ImprovedCNN : 1 423 114 parametres\n"
"Epoch  1/22 | loss 1.595 | train_acc 0.418 | val_acc 0.533 | lr 9.95e-04\n"
"Epoch 11/22 | loss 0.715 | train_acc 0.753 | val_acc 0.769 | lr 5.00e-04\n"
"Epoch 22/22 | loss 0.502 | train_acc 0.829 | val_acc 0.855 | lr 0.00e+00\n"
"Meilleure val_acc : 0.8552 | modele sauvegarde dans best_cnn_ameliore.params",
    label='Sortie reelle (CNN ameliore) :')
add_figure('p2_fig7.png', "Courbes d'apprentissage du CNN ameliore (22 epoques).", 6.2)
add_figure('p2_fig8.png', "Matrice de confusion du CNN ameliore sur CIFAR-10.", 4.8)
add_table(['Modele', 'Parametres', 'Acc. test', 'Augmentation'],
    [['MLP simple', '1 707 274', '0,520', 'Non'],
     ['CNN de reference', '53 770', '0,696', 'Non'],
     ['CNN ameliore', '1 423 114', '0,839', 'Oui']],
    caption='Bilan : MLP vs CNN de reference vs CNN ameliore.')
add_figure('p2_fig9.png', "Comparaison du CNN de reference et du CNN ameliore.", 5.5)
interpret("le modele ameliore atteint 83,9 % d'exactitude de test, soit +14,3 points sur la reference. "
     "Ce gain n'a de sens que compare a la baseline, d'ou l'interet d'entrainer les deux modeles. "
     "Trois facteurs l'expliquent : augmentation de donnees (moins de surapprentissage), profondeur "
     "accrue (hierarchie plus riche) et regularisation (BatchNorm, dropout, weight decay).")

h2('2.7 Question de synthese — Partie II')
para("Pourquoi un CNN est-il plus pertinent qu'un MLP pour la classification d'images, et comment les "
     "choix de padding, stride, pooling et profondeur influencent-ils les performances ?",
     bold=True, italic=True)
para("Un CNN encode deux biais inductifs adaptes aux images : la localite et l'invariance par "
     "translation. Cela donne une hierarchie de representations et une economie drastique de parametres, "
     "demontree experimentalement (53 770 parametres battent 1,7 M de 17 points). Les calculs "
     "dimensionnels montrent l'action de chaque choix : le padding preserve la taille spatiale (et "
     "l'information des bords), le stride sous-echantillonne (cout et taille reduits), le pooling agrege "
     "localement (invariance). L'etude ablative confirme que le nombre de filtres et la profondeur sont "
     "les leviers les plus puissants : le modele ameliore gagne 14,3 points. Les cartes de "
     "caracteristiques illustrent enfin la hierarchie aretes -> motifs abstraits qui fait la force des CNN.")
page_break()

# ======================================================================
# 4. PARTIE III — RNN / LSTM / GRU & Seq2Seq
# ======================================================================
h1('3. Partie III — RNN, LSTM, GRU et Seq2Seq')
para("Theme : modelisation de sequences. Le volet A compare RNN, LSTM et GRU sur un modele de langage "
     "caractere par caractere (corpus The Time Machine) ; le volet B construit un systeme encodeur-"
     "decodeur (Seq2Seq) de traduction anglais -> francais (corpus fra-eng / Tatoeba).")

h2('3.1 Fondements theoriques')
para("Un modele de langage factorise la probabilite d'une sequence par la regle de chaine : "
     "P(x1..xT) = produit de P(xt | x<t). Un RNN materialise cette factorisation via un etat cache ht "
     "transmis de pas en pas : ht = phi(Wxh xt + Whh ht-1 + b). La qualite se mesure par la perplexite "
     "PPL = exp(- moyenne des log P) ; plus elle est faible, meilleur est le modele. L'apprentissage par "
     "retropropagation a travers le temps (BPTT) fait apparaitre un produit de jacobiens, source "
     "d'evanouissement ou d'explosion du gradient ; les portes du LSTM (oubli, entree, sortie + cellule) "
     "et du GRU (mise a jour, reinitialisation) creent un chemin de gradient stable.")

h2('3.2 Cellule RNN maison contre PyTorch')
para("Le calcul d'un RNN est d'abord reconstruit a la main en deroulant la recurrence avec les poids "
     "d'un nn.RNN, puis compare a la sortie de PyTorch.")
add_output("Stats corpus : 173 427 caracteres nettoyes | vocabulaire = 28 caracteres\n"
"Ecart maximal entre RNN maison et nn.RNN : 1.19e-07", label='Sortie reelle (preparation + verification) :')
interpret("l'ecart de l'ordre de la precision machine confirme que le calcul du RNN est parfaitement "
     "compris et reproduit.")

h2('3.3 BPTT et gradient clipping')
para("Pour illustrer le gradient clipping, le meme RNN est entraine en SGD a pas eleve (lr = 2.0), "
     "reglage propice a l'explosion, avec puis sans ecretage de la norme au seuil 1.")
add_output(
"AVEC clipping  : norme mediane=0.59  max=2.85   | perplexite finale =    5.52\n"
"SANS clipping  : norme mediane=2.87  max=24.95  | perplexite finale = 1159.77",
    label='Sortie reelle (effet du clipping) :')
add_figure('p3_fig1.png', "Effet du gradient clipping : norme du gradient (gauche) et perplexite (droite).", 6.4)
interpret("sans ecretage, la norme du gradient explose (pics > 20) et la perplexite diverge (1159) ; "
     "avec un seuil de 1, la norme reste maitrisee et la perplexite converge (5,52). Le clipping ne "
     "change pas la direction du gradient, seulement sa magnitude : garantie de stabilite peu couteuse.")

h2('3.4 Comparaison RNN / LSTM / GRU')
para("Les trois cellules sont entrainees a budget strictement egal (40 epoques, Adam, clipping) sur le "
     "modele de langage caractere.")
add_output(
">> RNN  : epoque 40/40 | perplexite = 3.46 | 109 596 parametres | 97s\n"
">> LSTM : epoque 40/40 | perplexite = 2.78 | 406 044 parametres | 226s\n"
">> GRU  : epoque 40/40 | perplexite = 2.99 | 307 228 parametres | 267s",
    label='Sortie reelle (entrainement des 3 cellules) :')
add_table(['Modele', 'Perplexite finale', 'Parametres', 'Temps total (s)', 'Temps/epoque (s)'],
    [['RNN', '3,46', '109 596', '96,5', '2,41'],
     ['LSTM', '2,78', '406 044', '225,8', '5,64'],
     ['GRU', '2,99', '307 228', '267,3', '6,68']],
    caption='Comparaison RNN / LSTM / GRU (modele de langage caractere).')
add_figure('p3_fig2.png', "Convergence de la perplexite : RNN vs LSTM vs GRU.", 5.8)
para("La generation de texte a partir d'une meme amorce confirme qualitativement la hierarchie :")
add_output(
"[RNN ] the time traveller i saw the same the sun i saw the same the sun i saw...\n"
"[LSTM] the time traveller with a sudden thought i had started with the same sensation of the sun\n"
"[GRU ] the time traveller s face i saw a second the sphinx and the splutter were decay...",
    label='Sortie reelle (generation de texte) :')
interpret("le LSTM (2,78) et le GRU (2,99) atteignent une perplexite nettement plus basse que le RNN "
     "simple (3,46), confirmant que leurs portes ameliorent la memorisation du contexte. Le cout suit "
     "la complexite (LSTM le plus lourd). Le RNN tombe vite dans des boucles repetitives, tandis que le "
     "LSTM produit un anglais bien plus coherent — preuve de sa meilleure memoire a long terme.")

h2('3.5 Systeme Seq2Seq de traduction')
para("Le volet B construit un encodeur-decodeur recurrent (GRU) pour traduire de l'anglais vers le "
     "francais sur 6 000 paires courtes du corpus fra-eng. Les donnees sont preparees avec les tokens "
     "speciaux <pad>, <bos>, <eos>, <unk>, padding et masquage.")
add_output(
"Vocabulaire source (anglais) : 1706   |   cible (francais) : 2102   |   lots/epoque : 43\n"
"Exemple de lot : 'tom doesn't want to take a walk now . <eos>'\n"
"             -> 'tom ne veut pas se promener maintenant . <eos> <pad>'\n\n"
"Parametres du Seq2Seq : 2 264 630\n"
"epoque   1/150 | perte/token = 0.4793\n"
"epoque  76/150 | perte/token = 0.0121\n"
"epoque 150/150 | perte/token = 0.0148   (temps total 1910 s)",
    label='Sortie reelle (preparation + entrainement Seq2Seq) :')
add_figure('p3_fig3.png', "Courbe d'entrainement du systeme Seq2Seq (perte masquee par token).", 5.8)

h2('3.6 Decodage glouton, beam search et BLEU')
para("Deux strategies de decodage sont implementees a la main : le decodage glouton et le beam search "
     "(k meilleures hypotheses, normalisation de longueur). La qualite est mesuree par un BLEU egalement "
     "implemente a la main, sur l'entrainement et sur le test.")
add_table(['Methode', 'BLEU (entrainement)', 'BLEU (test)'],
    [['Glouton', '0,432', '0,032'],
     ['Beam k=2', '0,450', '0,032'],
     ['Beam k=3', '0,448', '0,032'],
     ['Beam k=5', '0,450', '0,032']],
    caption='BLEU : decodage glouton vs beam search, entrainement et test.')
add_figure('p3_fig4.png', "BLEU : glouton vs beam search, sur l'entrainement et le test.", 5.8)
add_table(['Source (EN)', 'Reference (FR)', 'Traduction (glouton)'],
    [['i will wait .', "j'attendrai .", '<unk> .'],
     ["don't be such a fool .", 'ne soyez pas si idiote !', 'ne sois pas si <unk> difficile !'],
     ['tom is pretty decisive .', 'tom est plutot categorique .', 'tom est plutot <unk> .'],
     ['tell me where you got this .', 'dis-moi ou tu as trouve ca .', 'dis-moi ou tu es fait ceci .']],
    caption='Exemples de traduction produits par le systeme Seq2Seq.', font=9)
interpret("les resultats separent apprentissage et generalisation. Sur l'entrainement, le BLEU est "
     "eleve (0,43) et le beam search depasse le glouton (0,45 vs 0,43), validant la theorie du decodage. "
     "Sur le test, le BLEU s'effondre (0,03) : un petit GRU sans attention, sur un corpus reduit, "
     "memorise mais generalise mal (frequents <unk> et non-emission du token <eos>). Ce fosse "
     "entrainement/test motive directement le mecanisme d'attention puis les Transformers.")

h2('3.7 Question de synthese — Partie III')
para("Dans quelle mesure les architectures recurrentes modelisent-elles une sequence reelle, et comment "
     "justifier le passage d'un RNN vers un LSTM/GRU puis vers un schema encodeur-decodeur ?",
     bold=True, italic=True)
para("Une sequence se factorise par la regle de chaine ; le RNN materialise cette factorisation via son "
     "etat cache, qui leve la sparsite et l'horizon borne des n-grammes. Nos perplexites (bien en dessous "
     "du hasard de 28) prouvent que le modele capte la structure du texte. Le passage du RNN au LSTM/GRU "
     "est justifie empiriquement (perplexite 3,46 -> 2,78, generation plus coherente) et theoriquement "
     "(les portes stabilisent le gradient face a la BPTT, le clipping traitant l'explosion mais pas "
     "l'evanouissement). Enfin, generer librement ne suffit pas pour traduire : il faut conditionner la "
     "sortie sur une entree, d'ou l'architecture encodeur-decodeur, avec teacher forcing, perte masquee et "
     "decodage (glouton/beam) mesure par BLEU. La limite observee — compression de tout le contexte dans "
     "un seul vecteur, d'ou un fort surapprentissage — est precisement ce que l'attention resout.")
page_break()

# ======================================================================
# 5. QUESTION TRANSVERSALE
# ======================================================================
h1('4. Question transversale finale')
para("Comment le deep learning adapte-t-il ses architectures a la structure des donnees — tabulaire, "
     "image et sequentielle — et pourquoi un meme paradigme d'apprentissage supervise doit-il etre "
     "decline differemment selon la geometrie, la dependance locale, la temporalite et la representation "
     "des donnees ?", bold=True, italic=True)
para("Les trois parties partagent un meme paradigme : un reseau differentiable optimise par "
     "retropropagation et descente de gradient. Ce qui change, c'est le biais inductif impose par "
     "l'architecture, choisi pour epouser la structure des donnees.")
bullet("Donnees tabulaires (MLP) : attributs sans geometrie ni ordre privilegie -> couches entierement "
       "connectees. Limite : aucune modelisation des correlations structurelles.")
bullet("Donnees images (CNN) : structure spatiale locale et invariante par translation -> convolutions "
       "(localite, partage des poids) et hierarchie de representations. Resultat : 32x moins de "
       "parametres que le MLP pour 17 points d'exactitude en plus.")
bullet("Donnees sequentielles (RNN/LSTM/GRU/Seq2Seq) : temporalite et longueur variable -> etat recurrent "
       "partage dans le temps, raffine par les portes, puis schema encodeur-decodeur pour relier deux "
       "sequences.")
para("Conclusion transversale : le paradigme supervise est unique, mais son efficacite depend de "
     "l'alignement entre le biais inductif de l'architecture et la structure des donnees. Absence de "
     "structure -> connexions denses ; dependance locale et invariance spatiale -> convolutions ; "
     "temporalite et ordre -> recurrence et portes. Choisir une architecture, c'est encoder a priori la "
     "geometrie du probleme.")
page_break()

# ======================================================================
# 6. CONCLUSION
# ======================================================================
h1('Conclusion generale')
para("Ce projet a couvert les trois grandes familles d'architectures du deep learning supervise, en les "
     "confrontant a des donnees reelles et en interpretant chaque resultat. Le tableau de synthese "
     "ci-dessous resume les performances obtenues.")
add_table(['Partie', 'Tache / dataset', 'Modele retenu', 'Metrique principale'],
    [['I — MLP', 'Adult Income (tabulaire)', 'MLP Xavier', 'Exactitude test 0,857 ; F1 0,671'],
     ['II — CNN', 'CIFAR-10 (images)', 'CNN ameliore', 'Exactitude test 0,839 (+14,3 pts)'],
     ['III — LM', 'The Time Machine (caracteres)', 'LSTM', 'Perplexite 2,78 (< GRU < RNN)'],
     ['III — Seq2Seq', 'fra-eng (traduction)', 'Encodeur-decodeur GRU', 'BLEU entrainement 0,45 (beam)']],
    caption='Synthese des resultats des trois parties.')
para("L'enseignement central est methodologique : la performance depend de l'adequation entre le biais "
     "inductif du modele et la structure des donnees. Le MLP convient au tabulaire, le CNN exploite la "
     "localite spatiale des images, les architectures recurrentes capturent la temporalite des sequences. "
     "Chaque partie a aussi mis en evidence des limites concretes — desequilibre de classes lu sur la "
     "matrice de confusion, bruit des etudes ablatives a faible budget, surapprentissage du Seq2Seq sans "
     "attention — qui tracent les pistes d'amelioration et la transition vers les architectures modernes "
     "a base d'attention.")

h1('Bibliographie')
bullet("Paszke, A. et al. (2019). PyTorch: An Imperative Style, High-Performance Deep Learning Library. "
       "Advances in Neural Information Processing Systems, 32.")
bullet("LeCun, Y., Bottou, L., Bengio, Y. & Haffner, P. (1998). Gradient-based learning applied to "
       "document recognition. Proceedings of the IEEE, 86(11), 2278-2324.")
bullet("Hochreiter, S. & Schmidhuber, J. (1997). Long Short-Term Memory. Neural Computation, 9(8).")
bullet("Cho, K. et al. (2014). Learning Phrase Representations using RNN Encoder-Decoder for Statistical "
       "Machine Translation. EMNLP.")
bullet("Sutskever, I., Vinyals, O. & Le, Q. (2014). Sequence to Sequence Learning with Neural Networks. "
       "NeurIPS.")
bullet("Zhang, A., Lipton, Z. C., Li, M. & Smola, A. J. (2023). Dive into Deep Learning. "
       "https://d2l.ai")
bullet("Pedregosa, F. et al. (2011). Scikit-learn: Machine Learning in Python. JMLR, 12, 2825-2830.")
bullet("Dua, D. & Graff, C. (2019). UCI Machine Learning Repository — Adult Data Set. "
       "https://archive.ics.uci.edu/")

# ---- Mise a jour des champs (TOC + listes) a l'ouverture ----
settings = doc.settings.element
uf = OxmlElement('w:updateFields'); uf.set(qn('w:val'), 'true')
settings.append(uf)

doc.save(OUT_DOCX)
print('RAPPORT COMPLET ecrit dans', OUT_DOCX, '| Blocs:', len(doc.element.body))
