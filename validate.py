#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate.py - controle technique des sources VBA de ClassicOutlook-Toolkit.

A LANCER AVANT DE COLLER QUOI QUE CE SOIT DANS LE VBE.

    python validate.py               # tout le dossier courant
    python validate.py modUi.txt     # un ou plusieurs fichiers
    python validate.py --quiet       # ne sort que les anomalies
    python validate.py --strict      # code retour 1 si AVERTISSEMENT aussi
    python validate.py --config config.ini   # cles inconnues du code (typos)
    python validate.py --gen-versions        # Select Case de ModuleVersionOf
    python validate.py --gen-config          # config_ini.example genere du code

Pack v3 (2026-09-02) : marqueurs d'accent etendus (~x ~c ~u ~U ~A ~y),
detection des marqueurs inconnus et des "e circonflexe" ecrits ~e, Dim en
double dans une procedure, collections COM re-evaluees dans une boucle,
controle des cles de config.ini, generateurs (versions, config example).

Lecture seule. Pas de COM, pas d'Outlook, pas de droits admin : le script
n'ouvre que des fichiers texte. Rien qui puisse alerter un EDR.

Chaque controle correspond a un incident REEL du projet ; la reference est
citee dans son en-tete. Ne pas ajouter de controle "par principe" : un
validateur qui crie pour rien finit ignore, et c'est ainsi qu'on laisse
passer le suivant.

Codes retour : 0 = aucune ERREUR, 1 = au moins une ERREUR (ou, avec
--strict, au moins un AVERTISSEMENT). Utilisable dans un hook git.
"""

import os
import re
import sys

# =============================================================================
# Parametres
# =============================================================================

VBA_MAX_LINE = 1023          # limite dure d'une ligne logique VBA
WARN_LINE = 900              # marge d'alerte avant la limite
VBA_MAX_CONT = 25            # continuations " _" par instruction logique

# Fichiers du depot qui ne sont PAS du source VBA.
NOT_VBA = {
    "CHANGELOG.txt",
    "config_ini_ADD_2026-08-22.txt",
    "Suivi_des_mails__Feuille_Suivi.txt",
    "modMailAnalytics__README.txt",
}

# Modules dont le nom de fichier ne porte pas de MOD_NAME (formulaires,
# fichiers de code de UserForm du projet Excel).
NO_MOD_NAME_OK = {
    "frmConfig.txt",
    "frmProgress.txt",
    "modMailAnalytics__frmConfig_Code.txt",
    "modMailAnalytics__frmProgress_Code.txt",
    "modMailAnalytics.txt",
    "modDebugLog.txt",       # le logger lui-meme
}

# Wrappers globaux de modConfig qu'une variable locale ne doit pas masquer.
# VBA est insensible a la casse : Dim t et la fonction T() se disputent la
# meme portee des lors que la procedure appelle le wrapper.
SHADOWING = {"t": "T", "a": "A"}

# Identifiants autorises A L'INTERIEUR d'un T() : ce sont des constantes de
# mise en forme, pas des donnees. Tout le reste traverse A() et peut etre
# corrompu (un sujet de mail ou un chemin contenant ~e, ~a, ~o).
T_SAFE = {
    "vbcrlf", "vblf", "vbtab", "vbnewline", "vbcr",
    "chr", "chrw", "chrw$", "chr$", "string", "space", "byval", "as",
}
# NOTE : Chr()/ChrW() sont surs a l'interieur de T(). Ils produisent un
# caractere deja accentue, que A() laisse passer tel quel - A() ne decode
# que les marqueurs "~X".
# Table des marqueurs de modConfig.A() (pack v3) :
#   ~e e aigu, ~f e grave, ~a a grave, ~E E aigu, ~i i circ, ~o o circ,
#   ~x e circ, ~c c cedille, ~u u circ, ~U u grave, ~A a circ, ~y i trema,
#   ~I I circ.
# Tout autre "~lettre" dans un T()/A() est un marqueur INCONNU (affiche tel
# quel a l'ecran) -> erreur MARKER.
KNOWN_MARKERS = "eafEioxcuUAyI"

# Incident A1 (pack v3) : 25 libelles "fen~etre / m~eme / ~etre / en-t~ete"
# ecrits avec ~e (e aigu) faute de marqueur e-circonflexe. Le marqueur ~x
# existe desormais ; ces formes sont signalees (avertissement ECIRC).
ECIRC_WORDS = re.compile(
    r"(fen~etre|m~eme|(?<![a-z])~etre|t~ete|arr~et|(?<![a-z])pr~et(?![a-z])"
    r"|for~et|f~ete|b~ete|qu~ete|r~eve(?![a-z])|m~ele(?![a-z])|gr~ele|ch~ene|cr~epe|d~ep~eche"
    r"|ench~ere|(?<![a-z])~etes(?![a-z])|apr~es|tr~es|acc~es|succ~es|proc~es)",
    re.I)

# Signatures publiques de modUi (nom -> position attendue du titre).
# Detecte l'inversion (buttons, title) -> (title, buttons), erreur deja
# commise une fois lors de la bascule de modQuickMove.
MODUI_TITLE_ARG = {"Ask": 1, "Inform": 1, "Fail": 1, "AskText": 1}

VB_CONST = re.compile(
    r"\bvb(YesNo|OKCancel|YesNoCancel|OKOnly|Critical|Exclamation|Information"
    r"|Question|DefaultButton\d|AbortRetryIgnore|RetryCancel"
    r"|MsgBoxSetForeground|SystemModal|ApplicationModal)\b"
)

RE_PROC = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I)
RE_END_PROC = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.I)
RE_DECL = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
    r"(Const|Declare|Type|Enum|Dim|WithEvents)\b", re.I)
RE_DECL_TYPED = re.compile(
    r"^\s*(?:Public|Private)\s+(?:WithEvents\s+)?\w+\s+As\b", re.I)


# =============================================================================
# Utilitaires
# =============================================================================

class Report:
    """Collecte les constats et les rend lisibles."""

    def __init__(self):
        self.errors = []
        self.warns = []

    def error(self, path, line, code, msg):
        self.errors.append((path, line, code, msg))

    def warn(self, path, line, code, msg):
        self.warns.append((path, line, code, msg))

    @property
    def clean(self):
        return not self.errors and not self.warns


def read_source(path):
    """Renvoie (texte_normalise_LF, octets_bruts) ou leve UnicodeDecodeError."""
    raw = open(path, "rb").read()
    try:
        txt = raw.decode("utf-8")
    except UnicodeDecodeError:
        txt = raw.decode("cp1252")
    return txt.replace("\r\n", "\n"), raw


def logical_lines(txt):
    """
    Recompose les lignes de continuation VBA (' _' en fin de ligne).
    Renvoie [(numero_1re_ligne, texte_recompose)].
    Indispensable : la moitie des controles est fausse sur des lignes
    physiques, une instruction VBA s'etalant souvent sur 6 lignes.
    """
    out, buf, start = [], "", None
    for i, line in enumerate(txt.split("\n"), 1):
        if start is None:
            start = i
        stripped = line.rstrip()
        if stripped.endswith(" _"):
            buf += stripped[:-1]
        else:
            out.append((start, buf + line))
            buf, start = "", None
    if buf:
        out.append((start or 1, buf))
    return out


def is_comment(line):
    return line.lstrip().startswith("'")


def strip_strings(s):
    """Retire les litteraux pour ne raisonner que sur le code."""
    return re.sub(r'"[^"]*"', '""', s)


def strip_T_calls(s):
    """Retire les appels T(...) et A(...) complets, imbrications comprises."""
    prev = None
    while prev != s:
        prev = s
        s = re.sub(
            r'\bT\(\s*"(?:[^"]|"")*"'
            r'(?:\s*&\s*(?:vbCrLf|vbLf|vbTab|"(?:[^"]|"")*"|[\w.$]+\([^()]*\)|[\w.$]+))*'
            r'\s*,\s*"(?:[^"]|"")*"'
            r'(?:\s*&\s*(?:vbCrLf|vbLf|vbTab|"(?:[^"]|"")*"|[\w.$]+\([^()]*\)|[\w.$]+))*\s*\)',
            "", s)
        s = re.sub(r'\bA\(\s*"[^"]*"\s*\)', "", s)
    return s


def call_span(line, start):
    """Contenu entre parentheses d'un appel demarrant a l'index start."""
    depth, k, in_str = 1, start, False
    while k < len(line) and depth > 0:
        c = line[k]
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
        k += 1
    return line[start:k - 1]


def split_args(inner):
    """Decoupe une liste d'arguments au niveau 0 de parentheses."""
    args, depth, in_str, cur = [], 0, False, ""
    for c in inner:
        if c == '"':
            in_str = not in_str
            cur += c
        elif not in_str and c == "(":
            depth += 1
            cur += c
        elif not in_str and c == ")":
            depth -= 1
            cur += c
        elif not in_str and c == "," and depth == 0:
            args.append(cur.strip())
            cur = ""
        else:
            cur += c
    args.append(cur.strip())
    return args


def procedures(txt):
    """Renvoie [(nom, 1re_ligne, corps)] pour chaque procedure."""
    out, name, buf, start = [], None, [], 0
    for i, line in enumerate(txt.split("\n"), 1):
        if is_comment(line):
            if name:
                buf.append(line)
            continue
        m = RE_PROC.match(line)
        if m and name is None:
            name, buf, start = m.group(2), [line], i
        elif name and RE_END_PROC.match(line):
            buf.append(line)
            out.append((name, start, "\n".join(buf)))
            name, buf = None, []
        elif name:
            buf.append(line)
    return out


# =============================================================================
# Controles
# =============================================================================

def check_encoding(path, txt, raw, rep):
    """ASCII strict, CRLF, fin de fichier. Incident : frmProgress/frmConfig."""
    for i, line in enumerate(txt.split("\n"), 1):
        bad = [c for c in line if ord(c) > 127]
        if bad:
            chars = " ".join(f"U+{ord(c):04X}" for c in sorted(set(bad)))
            rep.error(path, i, "ASCII",
                      f"caractere non-ASCII ({chars}) - "
                      f"utiliser A() ou Chr() ; accents interdits jusque dans "
                      f"les commentaires")
    if b"\r\n" not in raw:
        rep.error(path, 0, "CRLF", "fins de ligne LF - le VBE attend CRLF")
    elif not raw.endswith(b"\r\n"):
        rep.warn(path, 0, "CRLF", "pas de saut de ligne final")


def check_option_explicit(path, txt, rep):
    """Option Explicit en premiere ligne significative."""
    lines = txt.split("\n")
    for i, line in enumerate(lines, 1):
        s = line.strip()
        if not s:
            continue
        if s.lower() == "option explicit":
            if i != 1:
                rep.warn(path, i, "OPTEXPL",
                         f"Option Explicit en ligne {i} au lieu de 1 "
                         f"(convention de livraison du projet)")
            return
        if is_comment(line):
            continue
        rep.error(path, i, "OPTEXPL",
                  "Option Explicit absent avant la premiere instruction")
        return
    rep.error(path, 0, "OPTEXPL", "Option Explicit absent")


def check_no_attribute(path, txt, rep):
    """Le VBE regenere Attribute/VERSION : ces lignes ne doivent pas etre livrees."""
    for i, line in enumerate(txt.split("\n"), 1):
        if re.match(r"^\s*(Attribute|VERSION|BEGIN|END)\b", line):
            if line.strip() in ("END", "BEGIN"):
                continue
            rep.error(path, i, "ATTR",
                      f"ligne '{line.strip()[:40]}' - le VBE la regenere, "
                      f"elle ne doit pas figurer dans le .txt")


def check_decl_after_proc(path, txt, rep):
    """
    LE CONTROLE LE PLUS IMPORTANT.
    VBA exige toutes les declarations de niveau module AVANT la premiere
    procedure. Incident du 2026-08-22 : quatre modules non compilables
    (modInboxAlert, modBackupAuto, modBackupVBA, modReplySuspicion), crash
    d'Outlook 20-30 s apres le demarrage - un module non compilable rend
    modDefer non compilable, et l'echec survient DANS le callback SetTimer,
    hors de portee de tout gestionnaire d'erreur.
    """
    depth, first_proc = 0, None
    for i, line in enumerate(txt.split("\n"), 1):
        if is_comment(line) or not line.strip():
            continue
        if RE_PROC.match(line):
            if first_proc is None:
                first_proc = i
            depth = 1
        elif RE_END_PROC.match(line):
            depth = 0
        elif depth == 0 and first_proc is not None:
            if RE_DECL.match(line) or RE_DECL_TYPED.match(line):
                rep.error(path, i, "DECL",
                          f"declaration de niveau module apres la premiere "
                          f"procedure (ligne {first_proc}) - VBA ne compilera "
                          f"pas : '{line.strip()[:50]}'")


def check_balance(path, txt, rep):
    """Sub/Function/Property equilibres. Detecte un collage tronque."""
    body = "\n".join(l for l in txt.split("\n") if not is_comment(l))
    pairs = [
        ("Sub", r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+\w+",
         r"^\s*End\s+Sub\s*$"),
        ("Function", r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Function\s+\w+",
         r"^\s*End\s+Function\s*$"),
        ("Property", r"^\s*(?:(?:Public|Private|Friend)\s+)?Property\s+(?:Get|Let|Set)\s+\w+",
         r"^\s*End\s+Property\s*$"),
    ]
    for label, open_re, close_re in pairs:
        o = len(re.findall(open_re, body, re.M | re.I))
        c = len(re.findall(close_re, body, re.M | re.I))
        if o != c:
            rep.error(path, 0, "BALANCE",
                      f"{label} : {o} declaration(s) pour {c} 'End {label}' "
                      f"- fichier tronque ou bloc mal ferme")


def check_line_length(path, txt, rep):
    """
    Deux limites VBA distinctes, souvent confondues :
      - 1023 caracteres par ligne PHYSIQUE (et non logique) ;
      - 25 continuations " _" au maximum par instruction logique.
    Une instruction de 1900 caracteres etalee sur 12 lignes est donc
    parfaitement legale : c'est le cas de la moitie du depot.
    """
    for i, line in enumerate(txt.split("\n"), 1):
        n = len(line)
        if n > VBA_MAX_LINE:
            rep.error(path, i, "LONGLINE",
                      f"ligne physique de {n} caracteres (max {VBA_MAX_LINE})")
        elif n > WARN_LINE:
            rep.warn(path, i, "LONGLINE",
                     f"ligne physique de {n} caracteres (limite {VBA_MAX_LINE})")

    cont, start = 0, None
    for i, line in enumerate(txt.split("\n"), 1):
        if line.rstrip().endswith(" _"):
            if start is None:
                start = i
            cont += 1
        else:
            if cont > VBA_MAX_CONT:
                rep.error(path, start, "CONTLINE",
                          f"{cont} continuations ' _' (max {VBA_MAX_CONT})")
            cont, start = 0, None


def check_declare_ptrsafe(path, txt, rep):
    """
    64 bits : toute Declare exige PtrSafe.
    EXCEPTION : la branche 32 bits d'un bloc de compilation conditionnelle
    (#If VBA7 ... #Else ... #End If). PtrSafe y serait une erreur de syntaxe
    sur VBA6 - c'est justement la raison d'etre du bloc. Plusieurs modules
    du depot (modNotify, modFolderMerge, modMailCampaign) l'utilisent.
    """
    in_legacy = False
    for i, line in enumerate(txt.split("\n"), 1):
        s_ = line.strip()
        if re.match(r"#\s*If\b", s_, re.I):
            in_legacy = False
        elif re.match(r"#\s*Else\b", s_, re.I):
            in_legacy = True
        elif re.match(r"#\s*End\s+If\b", s_, re.I):
            in_legacy = False
        if is_comment(line) or in_legacy:
            continue
        if re.search(r"\bDeclare\s+(?!PtrSafe)", line):
            rep.error(path, i, "PTRSAFE",
                      "Declare sans PtrSafe - incompatible Outlook 64 bits")


def check_accent_markers(path, txt, rep):
    """
    Un marqueur d'accent (~e, ~a, ~o...) doit TOUJOURS etre dans T() ou A(),
    sinon il s'affiche tel quel a l'ecran.
    """
    for ln, line in logical_lines(txt):
        if is_comment(line):
            continue
        rest = strip_T_calls(line)
        for lit in re.findall(r'"([^"]*)"', rest):
            if re.search(r"~[A-Za-z]", lit):
                rep.error(path, ln, "ACCENT",
                          f"marqueur d'accent hors T()/A() : \"{lit[:45]}\"")
        # Pack v3 : a l'INTERIEUR des T()/A(), marqueur inconnu et formes
        # "e circonflexe" ecrites ~e (le projet analytics utilise le meme
        # A() de modConfig : "BO~ITES" y etait affiche tel quel).
        for lit in re.findall(r'"((?:[^"]|"")*)"', line):
            for m in re.finditer(r"~([A-Za-z])", lit):
                if m.group(1) not in KNOWN_MARKERS:
                    rep.error(path, ln, "MARKER",
                              f"marqueur d'accent inconnu ~{m.group(1)} "
                              f"(table A() : ~{' ~'.join(KNOWN_MARKERS)}) : "
                              f"\"{lit[:45]}\"")
            m2 = ECIRC_WORDS.search(lit)
            if m2:
                w = m2.group(0).lower()
                grave = any(w.endswith(x) for x in ("apr~es", "tr~es", "acc~es", "succ~es", "proc~es"))
                rep.warn(path, ln, "ECIRC",
                         f"\"{m2.group(0)}\" : accent probablement faux (~e = e aigu) "
                         f"- utiliser {'~f (e grave)' if grave else '~x (e circonflexe)'}")


def check_vars_in_T(path, txt, rep):
    """
    T() applique A() a ses DEUX branches. Toute donnee concatenee A
    L'INTERIEUR traverse donc le decodeur : un sujet de mail, un chemin ou
    un Err.Description contenant ~e / ~a / ~o est affiche FAUX.
    Incidents : ThisOutlookSession v10.6, modProjectTag v8.9,
    modMailResponseTracker v6, modReplySuspicion v5.
    """
    for ln, line in logical_lines(txt):
        if is_comment(line):
            continue
        if re.match(r"\s*(?:Public|Private|Friend)?\s*Function\s+T\s*\(", line, re.I):
            continue  # definition de T() elle-meme, pas un appel
        for m in re.finditer(r"(?<![.\w])T\(", line):
            inner = call_span(line, m.end())
            outside = re.sub(r'"[^"]*"', "", inner)
            ids = [x for x in re.findall(r"[A-Za-z_][\w.]*", outside)
                   if x.lower() not in T_SAFE]
            if ids:
                rep.error(path, ln, "TVAR",
                          f"variable(s) a l'interieur de T() : "
                          f"{', '.join(sorted(set(ids))[:4])} - les sortir de "
                          f"T() (A() corromprait leur contenu)")


def check_shadowing(path, txt, rep):
    """
    Dim t / Dim a dans une procedure qui appelle T() / A() : VBA etant
    insensible a la casse, la variable masque le wrapper global et leve
    "Declaration existante dans la portee".
    Le controle n'alerte QUE si le wrapper est reellement appele : sinon la
    declaration est seulement latente (17 cas dans le depot) et crier
    dessus rendrait le validateur inutilisable.
    """
    for name, start, body in procedures(txt):
        for var, wrapper in SHADOWING.items():
            if not re.search(rf"^\s*Dim\s+{var}\s+As\b", body, re.M | re.I):
                continue
            code = "\n".join(l for l in body.split("\n") if not is_comment(l))
            if re.search(rf"(?<![.\w]){wrapper}\(", code):
                rep.error(path, start, "SHADOW",
                          f"[{name}] Dim {var} masque le wrapper global "
                          f"{wrapper}() appele dans la meme procedure")
            else:
                rep.warn(path, start, "SHADOW",
                         f"[{name}] Dim {var} porte le nom du wrapper "
                         f"{wrapper}() - latent, a renommer si la procedure "
                         f"doit un jour l'appeler")


RE_OPEN_END = re.compile(r"[(&,+\-*/=]\s*$")


def check_unterminated_lines(path, txt, rep):
    """
    Pack v3 (incident modContactBirthday) : une ligne qui se termine par "(",
    "&", "," ou un operateur SANS " _" de continuation est une erreur de
    syntaxe VBA ("Erreur de syntaxe" a la compilation, statement rouge dans
    le VBE). Chaines et commentaires retires avant le test.
    """
    for ln, line in enumerate(txt.split("\n"), 1):
        s = line.rstrip()
        if not s or is_comment(line):
            continue
        core = strip_strings(s)
        core = re.sub(r"'.*$", "", core).rstrip()
        if core.endswith(" _") or core.lstrip().startswith("#"):
            continue
        if RE_OPEN_END.search(core):
            rep.error(path, ln, "UNTERM",
                      f"ligne non terminee sans ' _' de continuation : '{s.strip()[:60]}'")


def check_paren_balance(path, txt, rep):
    """
    2026-09-15 (incident modQuickStart) : une parenthese en trop ou en moins
    sur une instruction (continuations comprises) = "Erreur de syntaxe" a la
    compilation. Chaines et commentaires retires avant le comptage.
    """
    for ln, line in logical_lines(txt):
        s = line.strip()
        if not s or is_comment(line):
            continue
        core = strip_strings(s)
        core = re.sub(r"'.*$", "", core)
        if core.count("(") != core.count(")"):
            rep.error(path, ln, "PARENS",
                      f"parentheses desequilibrees ({core.count('(')} ouvrantes, "
                      f"{core.count(')')} fermantes) : '{s[:60]}'")


def check_dup_dim(path, txt, rep):
    """
    Pack v3 (D14). Deux "Dim x" dans une meme procedure = erreur de
    compilation "Declaration existante dans la portee" ; le module entier
    (et modDefer) devient non compilable. Les branches #If / #Else sont des
    portees distinctes (une seule est compilee).
    """
    for name, start, body in procedures(txt):
        seen = {}
        branch = 0
        for off, line in enumerate(body.split("\n")):
            s = line.strip()
            if is_comment(line):
                continue
            if re.match(r"#\s*(If|ElseIf)\b", s, re.I):
                branch += 1
                continue
            if re.match(r"#\s*Else\b", s, re.I):
                branch += 1
                continue
            if re.match(r"#\s*End\s+If\b", s, re.I):
                branch = 0
                continue
            m = re.match(r"\s*(?:Static\s+|Dim\s+)(.+)$", line, re.I)
            if not m:
                continue
            decl = strip_strings(m.group(1))
            decl = re.sub(r"'.*$", "", decl)
            decl = decl.split(":")[0]          # "Dim x As String: x = ..." -> declaration seule
            decl = re.sub(r"\([^)]*\)", "", decl)  # Dim arr(1 To 2, 1 To 3) : pas de virgule de liste
            for part in decl.split(","):
                var = re.match(r"\s*(\w+)", part)
                if not var:
                    continue
                key = var.group(1).lower()
                if key in seen and seen[key][1] == branch:
                    rep.error(path, start + off, "DUPDIM",
                              f"[{name}] Dim {var.group(1)} en double "
                              f"(deja ligne {seen[key][0]}) - ne compilera pas")
                else:
                    seen[key] = (start + off, branch)


RE_COM_LOOP = re.compile(r"\.(Folders|Items)\s*\.\s*(Item\s*\(|Count\b)", re.I)


def check_com_in_loop(path, txt, rep):
    """
    Pack v3 (B7). "parent.Folders.Item(i)" ou "x.Items.Count" evalue A
    CHAQUE tour d'une boucle For recree la collection COM (une descente MAPI
    par iteration, sur 1000+ dossiers). Poser "Set subs = parent.Folders"
    avant la boucle. Avertissement : un acces unique hors boucle est licite.
    """
    for name, start, body in procedures(txt):
        depth = 0
        for off, line in enumerate(body.split("\n")):
            if is_comment(line):
                continue
            s = line.strip()
            if re.match(r"For\b", s, re.I):
                depth += 1
                if RE_COM_LOOP.search(strip_strings(s)):
                    rep.warn(path, start + off, "COMLOOP",
                             f"[{name}] collection COM evaluee dans l'en-tete "
                             f"de boucle : '{s[:60]}'")
                continue
            if re.match(r"Next\b", s, re.I):
                depth = max(0, depth - 1)
                continue
            if depth > 0 and RE_COM_LOOP.search(strip_strings(s)):
                rep.warn(path, start + off, "COMLOOP",
                         f"[{name}] collection COM re-evaluee dans une boucle : "
                         f"'{s[:60]}'")


def check_modui_args(path, txt, rep):
    """
    modUi.Ask/Inform/Fail/AskText attendent (prompt, titre, ...).
    Detecte l'inversion avec l'ancienne signature MsgBox(prompt, buttons,
    titre) : erreur deja commise lors de la bascule de modQuickMove.
    """
    for ln, line in logical_lines(txt):
        if is_comment(line):
            continue
        for m in re.finditer(r"modUi\.(Ask|AskText|Inform|Fail|Toast)\b\s*\(", line):
            fn = m.group(1)
            args = split_args(call_span(line, m.end()))
            if fn == "Toast":
                if len(args) >= 2 and VB_CONST.search(args[1]):
                    rep.error(path, ln, "UIARGS",
                              "modUi.Toast : 2e argument = niveau/booleen, "
                              "pas une constante vb*")
                continue
            pos = MODUI_TITLE_ARG[fn]
            if len(args) <= pos:
                rep.error(path, ln, "UIARGS",
                          f"modUi.{fn} : titre manquant (2e argument obligatoire)")
            elif VB_CONST.search(args[pos]):
                rep.error(path, ln, "UIARGS",
                          f"modUi.{fn} : arguments inverses - le 2e doit etre "
                          f"le TITRE, les constantes vb* viennent en 3e")


def check_mod_name(path, txt, rep, known_short_names):
    """
    MOD_NAME present, et coherent avec MODULES_DEFAULT de modMaintenance.
    Un ecart rend le module invisible du filtre de log par module.
    Incident : ONProbe/OneNoteProbe, ReplyGreet/ReplyGreeting,
    StoreCtx/StoreContext.
    """
    base = os.path.basename(path)
    m = re.search(r'Private\s+Const\s+MOD_NAME\s+As\s+String\s*=\s*"([^"]+)"', txt)
    if not m:
        if base not in NO_MOD_NAME_OK:
            rep.warn(path, 0, "MODNAME",
                     "MOD_NAME absent - module invisible du filtre de log")
        return
    short = m.group(1)
    if known_short_names and short not in known_short_names:
        rep.error(path, 0, "MODNAME",
                  f"MOD_NAME \"{short}\" absent de MODULES_DEFAULT "
                  f"(modMaintenance) - le filtre de log ne l'atteindra pas")


def check_mod_version(path, txt, rep):
    """MOD_VERSION et ModVersion vont par paire."""
    base = os.path.basename(path)
    if base in NO_MOD_NAME_OK:
        return
    has_const = bool(re.search(r"Private\s+Const\s+MOD_VERSION\b", txt))
    has_func = bool(re.search(r"Public\s+Function\s+ModVersion\s*\(", txt))
    if has_const and not has_func:
        rep.error(path, 0, "MODVER",
                  "MOD_VERSION defini mais ModVersion() absent - la version "
                  "reelle est illisible depuis modMaintenance")
    elif has_func and not has_const:
        rep.error(path, 0, "MODVER", "ModVersion() sans constante MOD_VERSION")
    elif not has_const:
        rep.warn(path, 0, "MODVER",
                 "pas d'instrumentation de version (MOD_VERSION / ModVersion)")


def check_debug_print(path, txt, rep):
    """Debug.Print interdit en production cote Outlook (modDebugLog fait foi).
    Le projet Excel Analytics est un VBProject separe : exempte."""
    base = os.path.basename(path)
    if "Analytics" in base or base.startswith("frm"):
        return
    for i, line in enumerate(txt.split("\n"), 1):
        if is_comment(line):
            continue
        if re.search(r"(?<![.\w])Debug\.Print\b", line):
            rep.warn(path, i, "DBGPRINT",
                     "Debug.Print - utiliser modDebugLog (DTrace/DInfo/DError)")



def collect_public_symbols(folder, files):
    """
    Table des membres PUBLICS par module : procedures, constantes, variables
    et membres d'Enum. Sert au controle XREF ci-dessous.
    Cle = nom de module tel qu'on l'ecrit dans le code (sans .txt).
    """
    symbols, enum_members = {}, {}
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or "config_ini" in base or base.startswith("CHANGELOG"):
            continue
        mod = base[:-4] if base.endswith(".txt") else base
        try:
            txt, _ = read_source(os.path.join(folder, base)
                                 if not os.path.dirname(f) else f)
        except Exception:
            continue
        names, enums = set(), set()
        in_enum = False
        for line in txt.split("\n"):
            st = line.strip()
            if is_comment(st) or not st:
                continue
            m = re.match(r"Public\s+Enum\s+(\w+)", st, re.I)
            if m:
                in_enum = True
                names.add(m.group(1))
                continue
            if in_enum:
                if re.match(r"End\s+Enum", st, re.I):
                    in_enum = False
                    continue
                m = re.match(r"(\w+)", st)
                if m:
                    enums.add(m.group(1))
                    names.add(m.group(1))
                continue
            # En VBA, un Sub/Function SANS modificateur est PUBLIC par
            # defaut. Ne matcher que "Public ..." produirait de faux
            # positifs (cas reel : ThisOutlookSession.CleanOldFlags).
            m = re.match(
                r"(?:(Public|Private|Friend)\s+)?(?:Static\s+)?"
                r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
                st, re.I)
            if m:
                if (m.group(1) or "public").lower() != "private":
                    names.add(m.group(2))
                continue
            m = re.match(r"Public\s+(?:Const\s+|WithEvents\s+)?(\w+)", st, re.I)
            if m and m.group(1).lower() not in ("enum", "type", "declare"):
                names.add(m.group(1))
        symbols[mod] = names
        enum_members[mod] = enums
    return symbols, enum_members


def check_xref(path, txt, rep, symbols, all_enum_members):
    """
    References croisees entre modules. Deux controles :

    1. QUALIFIE - "modXxx.Membre" ou modXxx est un module connu du depot,
       mais Membre n'y est pas Public. C'est la cause typique d'un
       "Variable non definie" en cascade : un module a ete mis a jour, son
       appelant pas (ou l'inverse).

    2. ENUM NU - une constante d'Enum publique ("uiInfo", "uiError"...)
       utilisee alors qu'AUCUN module du depot ne la definit. Cas reel du
       2026-08-22 : modDevTools appelait modUi.Toast ..., uiInfo alors que
       le modUi charge etait encore en v1.2, sans l'Enum -> "Variable non
       definie" sur modDevTools, alors que le module fautif etait modUi.

    Ne remplace pas le compilateur : ne verifie ni les types, ni les
    arites, ni les references non qualifiees vers un autre module.
    """
    known = set(symbols)
    for ln, line in logical_lines(txt):
        if is_comment(line):
            continue
        code = strip_strings(line)

        for m in re.finditer(r"\b(mod\w+|cls\w+|ThisOutlookSession)\.(\w+)", code):
            mod, member = m.group(1), m.group(2)
            if mod not in known:
                continue
            if member not in symbols[mod]:
                rep.error(path, ln, "XREF",
                          f"{mod}.{member} : {member} n'est pas un membre "
                          f"Public de {mod} - module appelant ou appele pas "
                          f"a jour")

        for m in re.finditer(r"(?<![.\w])(ui[A-Z]\w*)", code):
            name = m.group(1)
            if name not in all_enum_members:
                rep.error(path, ln, "XREF",
                          f"{name} : constante d'Enum introuvable dans le "
                          f"depot - le module qui la definit n'est pas a jour")


def collect_short_names(folder):
    """Lit MODULES_DEFAULT dans modMaintenance.txt."""
    p = os.path.join(folder, "modMaintenance.txt")
    if not os.path.exists(p):
        return set()
    try:
        txt, _ = read_source(p)
    except Exception:
        return set()
    m = re.search(r"Private\s+Const\s+MODULES_DEFAULT\b", txt)
    if not m:
        return set()
    # La constante s'etale sur plusieurs lignes de continuation : on lit
    # jusqu'a la premiere ligne qui ne se termine pas par " _".
    block, i = "", m.start()
    for line in txt[i:].split("\n"):
        block += line + "\n"
        if not line.rstrip().endswith(" _"):
            break
    names = set()
    for lit in re.findall(r'"([^"]*)"', block):
        for x in lit.split("|"):
            x = x.strip()
            if x:
                names.add(x)
    return names


# =============================================================================
# Pack v3 : cles de configuration lues par le code, generateurs
# =============================================================================

RE_CFG = re.compile(r'\bCfg[LB]?\(\s*"([A-Za-z0-9_]+)"(?:\s*,\s*((?:"(?:[^"]|"")*")|-?\d+(?:\.\d+)?|True|False))?', re.I)
RE_CFG_DYN = re.compile(r'\bCfg[LB]?\(\s*"([A-Za-z0-9_]+)"\s*&', re.I)
# Prefixes de cles numerotees ("InboxAlertStore" & i) et cles construites
# par variable ("NewFolderParent" & category).
CFG_DYNAMIC_PREFIXES = ("InboxAlertStore", "InboxAlertColor", "InboxAlertSound",
                        "StatMailbox", "NewFolderParent", "OneNoteNotebook")


def collect_config_keys(folder, files):
    """Renvoie {cle: (module, defaut)} des Cfg*("Cle", defaut) du depot."""
    keys = {}
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or "config_ini" in base or base.startswith("CHANGELOG"):
            continue
        try:
            txt, _ = read_source(f)
        except Exception:
            continue
        for ln, line in logical_lines(txt):
            if is_comment(line):
                continue
            for m in RE_CFG.finditer(line):
                k = m.group(1)
                d = m.group(2) if m.group(2) is not None else ""
                if k not in keys:
                    keys[k] = (base[:-4], d)
    return keys


def config_file_keys(path):
    """Cles presentes dans un config.ini (hors commentaires / sections)."""
    out = []
    try:
        txt, _ = read_source(path)
    except Exception:
        return out
    for ln, line in enumerate(txt.split("\n"), 1):
        s = line.strip()
        if not s or s[0] in ";#[":
            continue
        if "=" in s:
            out.append((ln, s.split("=", 1)[0].strip()))
    return out


def check_config_file(cfg_path, folder, files, rep):
    """--config : cles du fichier inconnues du code (typos, cles obsoletes)."""
    known = {k.lower() for k in collect_config_keys(folder, files)}
    for ln, key in config_file_keys(cfg_path):
        kl = key.lower()
        if kl in known:
            continue
        if any(kl.startswith(p.lower()) for p in CFG_DYNAMIC_PREFIXES):
            continue
        rep.warn(cfg_path, ln, "CFGKEY",
                 f"cle '{key}' inconnue du code (typo ou cle obsolete ?)")


def gen_versions(folder, files):
    """--gen-versions : Select Case de modMaintenance.ModuleVersionOf."""
    cases = []
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or not (base.startswith("mod") or base == "ThisOutlookSession.txt"):
            continue
        if "MailAnalytics" in base:
            continue
        try:
            txt, _ = read_source(f)
        except Exception:
            continue
        m = re.search(r'Private\s+Const\s+MOD_NAME\s+As\s+String\s*=\s*"([^"]+)"', txt)
        short = m.group(1) if m else ("DebugLog" if base == "modDebugLog.txt" else None)
        if not short:
            continue
        if not re.search(r"^Public\s+Function\s+ModVersion\s*\(", txt, re.M):
            continue
        mod = base[:-4]
        call = "ModVersion()" if mod == "modMaintenance" else f"{mod}.ModVersion()"
        cases.append(f'        Case "{short}":{" " * max(1, 13 - len(short))}ModuleVersionOf = {call}')
    print("\n".join(sorted(cases)))
    print('        Case Else:        ModuleVersionOf = ""')


def gen_config_example(folder, files):
    """--gen-config : template config_ini.example genere depuis le code."""
    keys = collect_config_keys(folder, files)
    by_mod = {}
    for k, (mod, d) in keys.items():
        by_mod.setdefault(mod, []).append((k, d))
    print("; config_ini.example - GENERE par validate.py --gen-config")
    print("; Une cle par ligne : Cle=valeur. Valeurs par defaut = celles du code.")
    print("; Les cles numerotees (InboxAlertStore1, StatMailbox1...) se declinent en N.")
    print("[ClassicOutlookToolkit]")
    for mod in sorted(by_mod):
        print(f"\n; --- {mod} ---")
        for k, d in sorted(by_mod[mod]):
            v = d
            if v.startswith('"') and v.endswith('"'):
                v = v[1:-1].replace('""', '"')
            print(f"{k}={v}")


# =============================================================================
# Pilote
# =============================================================================

CHECKS_TXT = [
    check_encoding,       # signature (path, txt, raw, rep)
]


def validate_file(path, rep, known, symbols=None, enum_members=None):
    try:
        txt, raw = read_source(path)
    except Exception as exc:
        rep.error(path, 0, "READ", f"lecture impossible : {exc}")
        return

    check_encoding(path, txt, raw, rep)
    check_option_explicit(path, txt, rep)
    check_no_attribute(path, txt, rep)
    check_decl_after_proc(path, txt, rep)
    check_balance(path, txt, rep)
    check_line_length(path, txt, rep)
    check_declare_ptrsafe(path, txt, rep)
    check_accent_markers(path, txt, rep)
    check_vars_in_T(path, txt, rep)
    check_shadowing(path, txt, rep)
    check_dup_dim(path, txt, rep)
    check_unterminated_lines(path, txt, rep)
    check_paren_balance(path, txt, rep)
    check_com_in_loop(path, txt, rep)
    check_modui_args(path, txt, rep)
    check_mod_name(path, txt, rep, known)
    check_mod_version(path, txt, rep)
    check_debug_print(path, txt, rep)
    if symbols is not None:
        check_xref(path, txt, rep, symbols, enum_members)


def main(argv):
    quiet = "--quiet" in argv
    strict = "--strict" in argv
    cfg_path = None
    if "--config" in argv:
        i = argv.index("--config")
        if i + 1 < len(argv):
            cfg_path = argv[i + 1]
            argv = argv[:i] + argv[i + 2:]
    args = [a for a in argv if not a.startswith("--")]

    if args:
        files = args
        folder = os.path.dirname(os.path.abspath(args[0])) or "."
    else:
        folder = "."
        files = sorted(f for f in os.listdir(folder)
                       if f.endswith(".txt") and f not in NOT_VBA)

    if "--gen-versions" in argv:
        gen_versions(folder, [os.path.join(folder, f) if not os.path.isabs(f) and not os.path.exists(f) else f for f in files])
        return 0
    if "--gen-config" in argv:
        gen_config_example(folder, [os.path.join(folder, f) if not os.path.isabs(f) and not os.path.exists(f) else f for f in files])
        return 0

    known = collect_short_names(folder)
    symbols, enums_by_mod = collect_public_symbols(folder, files)
    all_enum_members = set()
    for v in enums_by_mod.values():
        all_enum_members |= v
    rep = Report()
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or "config_ini" in base or base.startswith("CHANGELOG"):
            continue
        validate_file(f, rep, known, symbols, all_enum_members)

    if cfg_path:
        check_config_file(cfg_path, folder, [os.path.join(folder, f) if not os.path.exists(f) else f for f in files], rep)

    def show(items, label):
        if not items:
            return
        print(f"\n{label} ({len(items)})")
        print("-" * 78)
        cur = None
        for path, line, code, msg in items:
            base = os.path.basename(path)
            if base != cur:
                print(f"  {base}")
                cur = base
            loc = f"L{line}" if line else "  -"
            print(f"     {loc:>6}  [{code}]  {msg}")

    show(rep.errors, "ERREURS - a corriger avant collage dans le VBE")
    if not quiet:
        show(rep.warns, "AVERTISSEMENTS")

    print()
    print("=" * 78)
    n = len([f for f in files if os.path.basename(f) not in NOT_VBA])
    if rep.clean:
        print(f"{n} fichier(s) - aucune anomalie.")
    else:
        print(f"{n} fichier(s) - {len(rep.errors)} erreur(s), "
              f"{len(rep.warns)} avertissement(s).")
    print("=" * 78)

    if rep.errors:
        return 1
    if strict and rep.warns:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
