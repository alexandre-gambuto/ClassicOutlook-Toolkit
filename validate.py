#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate.py - controle technique des sources VBA de ClassicOutlook-Toolkit.

A LANCER AVANT DE COLLER QUOI QUE CE SOIT DANS LE VBE.

    python validate.py               # tout Code/ (voir 2026-10-25 plus bas)
    python validate.py modUi.txt     # un ou plusieurs fichiers
    python validate.py --quiet       # ne sort que les anomalies
    python validate.py --strict      # code retour 1 si AVERTISSEMENT aussi
    python validate.py --config config.ini   # cles inconnues, en double, vides
    python validate.py --against <dossier>   # recul par rapport a une reference
    python validate.py --gen-versions        # Select Case de ModuleVersionOf
    python validate.py --gen-config          # config_ini.example genere du code

Pack v3 (2026-09-02) : marqueurs d'accent etendus (~x ~c ~u ~U ~A ~y),
detection des marqueurs inconnus et des "e circonflexe" ecrits ~e, Dim en
double dans une procedure, collections COM re-evaluees dans une boucle,
controle des cles de config.ini, generateurs (versions, config example).

2026-10-24 : code du classeur Campagnes.xlsm (Campagnes__ThisWorkbook,
Campagnes__modCampaignBook) exempte de MOD_NAME / MOD_VERSION comme le
classeur Suivi des mails (projet VBA Excel distinct).

2026-10-25 : sources rangees par application sous Code/ (Outlook/,
Excel/<classeur>/, Interface/Rubans/ et Interface/Formulaires/, Docs/,
Outils/), ce script dans Code/Outils/. Sans fichier en argument : tout
Code/, sous-dossiers compris, que le script soit lance depuis
Code/Outils, depuis Code/ ou depuis le dossier de l'outil ; un depot a
plat (GitHub) se controle comme avant, depuis sa racine. Erreur DOUBLON
si un meme fichier existe a deux endroits (livraison posee au mauvais
endroit : on risquerait de coller l'ancien).

2026-11-16 (audit 11/2026) :
  - VERSION (erreur, F0-02) : le MOD_VERSION d'un module doit etre >= au
    plus haut numero cite pour lui dans CHANGELOG.txt depuis
    VERSION_SINCE. Incident : quatre livraisons refaites sur une ancienne
    copie (modFolderSort v3.1 du 18/08 ecrasee par une "v2" du 22/08,
    modMaintenance, modSuiviTemps, modBackupAuto).
    REGLE DE LIVRAISON : toute modification part du fichier de Code\\
    (jamais d'une copie de l'espace du projet) et recoit un numero de
    version NOUVEAU, superieur a tout numero deja employe.
  - --against <dossier> (alias --installe) : RECUL (erreur) quand une
    entree d'historique "vN (AAAA-MM-JJ)" de l'en-tete de la reference a
    disparu du nouveau fichier, ou quand le contenu differe sans que
    MOD_VERSION augmente. Reference : dernier Code\\ livre, ou modules
    extraits de la sauvegarde du projet installe.
  - --config : CFGDUP (cle en double avec deux valeurs, la derniere gagne)
    et CFGEMPTY (cle presente mais vide alors que le code a un defaut non
    vide : Cfg() renvoie alors du vide). Cles : appels Cfg/CfgL/CfgB a DEUX
    arguments seulement (un dictionnaire local nomme cfg("Keyword") n'est
    plus pris pour une cle, F2-02).
  - MSGLOOP (avertissement, F2-01) : texte construit dans une boucle puis
    passe a une boite (MsgBox, InputBox, modUi, MsgBoxC) : risque de
    depasser 1 024 caracteres. Commentaire "MSGLOOP-OK" sur la ligne
    precedente pour une boucle bornee.
  - ITREMA (avertissement) : "~yt" (i trema) presque toujours mis pour un
    i circonflexe (~i) : "reconna~yt".
  - ACCENT : guillemets doubles "" acceptes dans A("...") (6 fausses
    erreurs dans modSuiviTemps).
  - Fichier dont l'en-tete porte "ARCHIVE (MM/AAAA)" (modules sortis du
    projet installe, F9-14) : non controle ; toute reference
    modArchive.Membre ailleurs est une ERREUR XREF (ne compilerait plus).
  - modMailAnalytics et ses formulaires ne sont plus dispenses (F9-12) ;
    seuls les fichiers de code des formulaires restent dispenses de
    MOD_VERSION. Debug.Print toleres seulement dans le code des classeurs.
  - --gen-config : cles dont le defaut du code n'est pas vide ecrites en
    commentaire (";Cle=defaut"), defauts construits ou constantes
    signales, cles numerotees documentees (F2-02).

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
    # 2026-10-05 : code du classeur Suivi des mails.xlsm (projet VBA Excel
    # distinct, sans modDebugLog ni modMaintenance).
    "Suivi_des_mails__ThisWorkbook.txt",
    "Suivi_des_mails__modMailTracker.txt",
    # 2026-10-24 : code du classeur Campagnes.xlsm (meme statut).
    "Campagnes__ThisWorkbook.txt",
    "Campagnes__modCampaignBook.txt",
    # 2026-10-25 : meme code, fichiers nommes d'apres le classeur
    # "Suivi des Campagnes.xlsm".
    "Suivi des Campagnes__ThisWorkbook.txt",
    "Suivi des Campagnes__modCampaignBook.txt",
    # 2026-11-16 : code du classeur Suivi de temps.xlsm (meme statut ; un
    # MOD_NAME y produirait une erreur "absent de MODULES_DEFAULT").
    "Suivi de temps__modSuiviTemps.txt",
    "Suivi de temps__ThisWorkbook.txt",
    "frmConfig.txt",
    "frmProgress.txt",
    # 2026-11-16 (F9-12) : modMailAnalytics.txt et les deux fichiers de code
    # des formulaires retires de cette liste : ils vivent dans le projet
    # Outlook (MOD_NAME "Analytics", modDebugLog).
    "modDebugLog.txt",       # le logger lui-meme
}

# 2026-11-16 (F9-12) : dispenses du seul controle MOD_VERSION / ModVersion.
# Un formulaire n'a pas de ModVersion() : frmConfig.ModVersion chargerait le
# formulaire et lancerait son Initialize. Sa version suit modMailAnalytics.
NO_MOD_VERSION_OK = {
    "modMailAnalytics__frmConfig_Code.txt",
    "modMailAnalytics__frmProgress_Code.txt",
}

# 2026-11-16 (F9-14) : en-tete d'un module archive (sorti du projet installe,
# garde dans le depot). Le fichier n'est pas controle ; une reference a ce
# module depuis un autre fichier est une erreur (ne compilerait plus).
RE_ARCHIVED = re.compile(r"^\s*'\s*ARCHIVE\s*\(\d{2}/\d{4}\)", re.M)
ARCHIVE_HEADER_LINES = 40

# 2026-11-16 (F0-02) : controle VERSION. Numeros de CHANGELOG.txt pris en
# compte a partir de cette date : avant, plusieurs modules ont ete
# renumerotes (ThisOutlookSession v12 -> v10.x en juillet, modQuickMove
# v15 -> v10.x en juin, modMaintenance v1.2 -> v1 en mai...).
VERSION_SINCE = "2026-08-01"
RE_CHANGELOG_DATE = re.compile(r"^\[(\d{4}-\d{2}-\d{2})\]")
RE_CHANGELOG_CITE = re.compile(r"\b(mod\w+|cls\w+|ThisOutlookSession)\s+[vV](\d+(?:\.\d+)*)\b")
RE_MOD_VERSION = re.compile(r'Private\s+Const\s+MOD_VERSION\s+As\s+String\s*=\s*"\s*[vV]?(\d+(?:\.\d+)*)')
# Entree d'historique d'en-tete : "' v4.5 (2026-10-25) : ..." (RECUL).
RE_HISTORY_ENTRY = re.compile(r"^\s*'\s*[vV](\d+(?:\.\d+)*)\s*\((\d{4}-\d{2}-\d{2})")

# 2026-11-16 (F2-01) : boites dont le texte est coupe vers 1 024 caracteres
# (MSGLOOP) : nom de l'appel -> le texte est le 1er argument.
MSG_CALLS = re.compile(
    r"(?<![.\w])(MsgBox|InputBox|MsgBoxC|Confirm3|modUi\.(?:Ask|AskText|AskTextEx|Inform|Fail)"
    r"|modMsgBoxCustom\.(?:MsgBoxC|Confirm3|Choose3))\b\s*(\(?)", re.I)
MSGLOOP_OK = "MSGLOOP-OK"

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
        # 2026-11-16 : guillemets doubles "" acceptes dans le litteral.
        s = re.sub(r'\bA\(\s*"(?:[^"]|"")*"\s*\)', "", s)
    return s


def is_archived(txt):
    """En-tete "ARCHIVE (MM/AAAA)" dans les premieres lignes (F9-14)."""
    head = "\n".join(txt.split("\n")[:ARCHIVE_HEADER_LINES])
    return bool(RE_ARCHIVED.search(head))


def version_tuple(s):
    """'4.5' -> (4, 5) ; None si illisible."""
    try:
        return tuple(int(x) for x in s.split("."))
    except (ValueError, AttributeError):
        return None


def module_key(path):
    """Nom de module tel qu'ecrit dans CHANGELOG : modX ; pour le code d'un
    classeur "Classeur__modY.txt" : modY. ThisWorkbook : None (ambigu)."""
    base = os.path.basename(path)
    if base.endswith(".txt"):
        base = base[:-4]
    key = base.split("__")[-1]
    if key.lower() == "thisworkbook":
        return None
    return key


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
        # 2026-11-16 : litteral avec guillemets doubles "" lu en entier.
        for lit in re.findall(r'"((?:[^"]|"")*)"', rest):
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
            # 2026-11-16 : "~yt" = i trema devant t, presque toujours mis
            # pour un i circonflexe (reconna~yt, appara~ytre).
            m3 = re.search(r"\w*(?<!u)~yt\w*", lit)     # ambigu~yte : juste
            if m3:
                rep.warn(path, ln, "ITREMA",
                         f"\"{m3.group(0)}\" : ~y = i trema ; i circonflexe probable "
                         f"- utiliser ~i")


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
    if base in NO_MOD_NAME_OK or base in NO_MOD_VERSION_OK:
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


def is_workbook_code(path):
    """Code d'un classeur Excel (projet VBA distinct, sans modDebugLog)."""
    base = os.path.basename(path)
    if base.startswith(("Suivi", "Campagnes")):
        return True
    # Code\Excel\<classeur>\fichier.txt : "Excel" = dossier grand-parent.
    parts = [p.lower() for p in os.path.normpath(os.path.abspath(path)).split(os.sep)]
    return "excel" in parts[-3:-1]


def check_debug_print(path, txt, rep):
    """Debug.Print interdit en production cote Outlook (modDebugLog fait foi).
    2026-11-16 (F9-12) : seul le code des classeurs Excel (projets VBA
    distincts, sans modDebugLog) est exempte. modMailAnalytics et ses
    formulaires vivent dans le projet Outlook : controles."""
    if is_workbook_code(path):
        return
    for i, line in enumerate(txt.split("\n"), 1):
        if is_comment(line):
            continue
        if re.search(r"(?<![.\w])Debug\.Print\b", line):
            rep.warn(path, i, "DBGPRINT",
                     "Debug.Print - utiliser modDebugLog (DTrace/DInfo/DError)")



RE_LOOP_OPEN = re.compile(r"^\s*(For|Do|While)\b", re.I)
RE_LOOP_CLOSE = re.compile(r"^\s*(Next|Loop|Wend)\b", re.I)
RE_GROW = re.compile(r"^\s*(\w+)\s*=\s*(\w+)\s*&", re.I)
# Borne de boucle For jugee limitee : constante en capitales (NB_BINDINGS),
# nombre, ou nom qui dit une page / un plafond (pageEnd, maxShown, iLast...).
RE_BOUNDED = re.compile(r"\b[A-Z][A-Z0-9_]{2,}\b|^\s*\d+\s*$"
                        r"|page|max|cap|limit|size|top|first|last|half|end|stop|shown", 0)


def for_is_bounded(stmt):
    """Boucle For dont la borne est visiblement limitee (voir RE_BOUNDED)."""
    m = re.match(r"^\s*For\s+Each\s+\w+\s+In\s+(.+)$", stmt, re.I)
    if m:
        bound = m.group(1)
    else:
        m = re.search(r"\bTo\b(.+?)(?:\bStep\b.*)?$", stmt, re.I)
        bound = m.group(1) if m else ""
    return bool(RE_BOUNDED.search(bound.strip())) if bound.strip() else False


def check_msg_loop(path, txt, rep):
    """
    2026-11-16 (F2-01). Une MsgBox comme une InputBox est coupee vers 1 024
    caracteres. Signale le texte d'une boite (1er argument de MsgBox,
    InputBox, modUi.Ask/AskText/AskTextEx/Inform/Fail, MsgBoxC, Confirm3)
    qui contient une variable allongee (x = x & ...) dans une boucle For
    sans borne visible (For Each sur une collection, For ... To n) de la
    meme procedure. Incidents : MaintModuleVersions (24 modules sur 58
    visibles), MaintFilterModules (32 sur 65). Non signale : boucle bornee
    par une constante en capitales ou un nom de page / plafond (pageEnd,
    maxShown...), boucles Do / While (dialogues repetes). Boucle bornee
    autrement : commentaire "MSGLOOP-OK" sur la ligne qui precede l'appel.
    """
    for name, start, body in procedures(txt):
        raw = body.split("\n")
        grown = {}
        stack = []                    # pour chaque boucle ouverte : None (Do/While) ou borne (For)
        # logical_lines renumerote a partir de 1 : decaler sur la procedure.
        for ln, line in logical_lines(body):
            if is_comment(line):
                continue
            # Litteraux vides de meme longueur (les positions restent celles
            # de line), commentaire de fin retire.
            code = re.sub(r'"[^"]*"', lambda mm: '"' + " " * (len(mm.group(0)) - 2) + '"', line)
            if "'" in code:
                code = code[:code.index("'")]
            for stmt in code.split(":"):
                mo = RE_LOOP_OPEN.match(stmt)
                if mo and not re.match(r"^\s*Do\s+While\b.*\bLoop\b", stmt, re.I):
                    if mo.group(1).lower() == "for":
                        stack.append(for_is_bounded(stmt))
                    else:
                        stack.append(None)
                elif RE_LOOP_CLOSE.match(stmt):
                    if stack:
                        stack.pop()
                m = RE_GROW.match(stmt)
                if m and m.group(1).lower() == m.group(2).lower():
                    fors = [b for b in stack if b is not None]
                    if fors and fors[-1] is False:
                        grown.setdefault(m.group(1).lower(), start + ln - 1)
            if not grown:
                continue
            for mc in MSG_CALLS.finditer(code):
                if mc.group(2) == "(":
                    args = split_args(call_span(line, mc.end()))
                else:
                    args = split_args(line[mc.end():len(code)])
                if not args:
                    continue
                first = re.sub(r'"(?:[^"]|"")*"', '""', args[0])
                ids = {x.lower() for x in re.findall(r"[A-Za-z_]\w*", first)}
                hit = sorted(v for v in grown if v in ids)
                if not hit:
                    continue
                prev = raw[ln - 2] if ln >= 2 else ""
                if MSGLOOP_OK in prev or MSGLOOP_OK in raw[ln - 1]:
                    continue
                rep.warn(path, start + ln - 1, "MSGLOOP",
                         f"[{name}] {mc.group(1)} : texte construit dans une boucle "
                         f"({', '.join(hit)}, ligne {grown[hit[0]]}) - une boite est "
                         f"coupee vers 1 024 caracteres (fichier + notepad, ou pages)")


def collect_archived(files):
    """Modules archives (en-tete "ARCHIVE (MM/AAAA)") : {nom de module}."""
    out = set()
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or not base.endswith(".txt"):
            continue
        try:
            txt, _ = read_source(f)
        except Exception:
            continue
        if is_archived(txt):
            out.add(base[:-4])
    return out


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
        if is_archived(txt):
            continue                  # 2026-11-16 : hors projet installe
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


def check_xref(path, txt, rep, symbols, all_enum_members, archived=frozenset()):
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
            if mod in archived:
                # 2026-11-16 (F9-14) : module archive, absent du projet
                # installe -> "Variable non definie" a la compilation.
                rep.error(path, ln, "XREF",
                          f"{mod}.{member} : {mod} est archive (hors projet "
                          f"installe) - retirer cette reference")
                continue
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


def collect_short_names(folder, files=()):
    """Lit MODULES_DEFAULT dans modMaintenance.txt."""
    p = os.path.join(folder, "modMaintenance.txt")
    # 2026-10-25 : Code range par application, modMaintenance.txt est
    # dans Code\Outlook : on le prend dans la liste des fichiers.
    for f in files:
        if os.path.basename(f) == "modMaintenance.txt" and os.path.exists(f):
            p = f
            break
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

# 2026-11-16 (F2-02) : une cle de config.ini = appel Cfg / CfgL / CfgB (ou le
# helper ReadUrl de modSendToLLM) a DEUX arguments dont le premier est un
# litteral. Exclus : cle construite ("InboxAlertStore" & i, "NewFolderParent"
# & cat), cle passee par variable, et dictionnaire local nomme cfg
# (cfg("Keyword") = ..., un seul argument : faux "Keyword=" dans le modele).
RE_CFG = re.compile(r'\b(Cfg[LB]?|ReadUrl)\(\s*"([A-Za-z0-9_]+)"\s*,', re.I)
# Prefixes de cles numerotees ("InboxAlertStore" & i) et cles construites
# par variable ("NewFolderParent" & category) : documentees a part.
CFG_DYNAMIC_PREFIXES = ("InboxAlertStore", "InboxAlertColor", "InboxAlertSound",
                        "StatMailbox", "NewFolderParent", "OneNoteNotebook")
CFG_DYNAMIC_DOC = (
    "InboxAlertStoreN / InboxAlertColorN / InboxAlertSoundN (N = 1, 2... : une serie par boite)",
    "StatMailboxN (N = 1, 2... : boites des statistiques)",
    "NewFolderParentProject / NewFolderParentContract (obligatoires pour \"Creer le dossier\")",
    "OneNoteNotebook1 a 3 (module OneNote archive)",
)
# Cles internes d'un dictionnaire de modMailAnalytics (jamais dans config.ini),
# filet de securite si un appel a deux arguments les reintroduisait.
CFG_NOT_KEYS = {"addressmode", "colorindex", "colorrgb", "datemin",
                "dateminenabled", "domain", "emailuser", "keyword"}


def classify_default(expr):
    """Defaut d'un appel Cfg : (valeur affichable, nature).
    nature : 'empty' | 'literal' | 'number' | 'bool' | 'constant' | 'built'."""
    e = expr.strip()
    m = re.fullmatch(r'"((?:[^"]|"")*)"', e)
    if m:
        v = m.group(1).replace('""', '"')
        return (v, "literal" if v else "empty")
    if re.fullmatch(r"-?\d+(?:\.\d+)?", e):
        return (e, "number")
    if e.lower() in ("true", "false"):
        return ("1" if e.lower() == "true" else "0", "bool")
    if re.fullmatch(r"[A-Za-z_][\w.]*(?:\(\s*\))?", e):
        return (e, "constant")
    return (e, "built")


def collect_config_keys(folder, files):
    """
    {cle: infos} des Cfg*("Cle", defaut) du depot. infos = dict(module,
    default, nature, kinds (fonctions qui la lisent), nonempty_str (lue par
    Cfg() avec un defaut non vide : une valeur vide dans config.ini
    l'annule)). Premier defaut non vide retenu pour l'affichage.
    """
    keys = {}
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or "config_ini" in base or base.startswith("CHANGELOG"):
            continue
        try:
            txt, _ = read_source(f)
        except Exception:
            continue
        if is_archived(txt):
            continue
        # Constantes du module a valeur simple : defaut affiche en clair
        # (DEF_DELAY_CLAUDE = 4000 -> ";LLM_DelayClaudeMs=4000").
        consts = {}
        for mc in re.finditer(r'^\s*(?:Private\s+|Public\s+)?Const\s+(\w+)(?:\s+As\s+\w+)?\s*=\s*'
                              r'("(?:[^"]|"")*"|-?\d+(?:\.\d+)?)\s*(?:\'.*)?$', txt, re.M | re.I):
            consts[mc.group(1).lower()] = mc.group(2)
        for ln, line in logical_lines(txt):
            if is_comment(line):
                continue
            for m in RE_CFG.finditer(line):
                fn, k = m.group(1), m.group(2)
                if k.lower() in CFG_NOT_KEYS:
                    continue
                args = split_args(call_span(line, line.index("(", m.start()) + 1))
                if len(args) < 2:
                    continue
                d, nature = classify_default(args[1])
                if nature == "constant" and d.lower() in consts:
                    d, nature = classify_default(consts[d.lower()])
                info = keys.setdefault(k, {"module": base[:-4], "default": d,
                                           "nature": nature, "kinds": set(),
                                           "nonempty_str": False})
                info["kinds"].add(fn.lower())
                if info["nature"] == "empty" and nature != "empty":
                    info.update(module=base[:-4], default=d, nature=nature)
                if fn.lower() == "cfg" and nature != "empty":
                    info["nonempty_str"] = True
    return keys


def config_file_entries(path):
    """Cles presentes dans un config.ini (hors commentaires / sections) :
    [(ligne, cle, valeur)]."""
    out = []
    try:
        txt, _ = read_source(path)
    except Exception:
        return out
    for ln, line in enumerate(txt.replace("\r", "\n").split("\n"), 1):
        s = line.strip()
        if not s or s[0] in ";#[":
            continue
        if "=" in s:
            k, v = s.split("=", 1)
            out.append((ln, k.strip(), v.strip()))
    return out


def config_file_keys(path):
    """Compatibilite : [(ligne, cle)]."""
    return [(ln, k) for ln, k, _ in config_file_entries(path)]


def check_config_file(cfg_path, folder, files, rep):
    """
    --config : cles du fichier inconnues du code (typos, cles obsoletes) ;
    2026-11-16 : CFGDUP (meme cle deux fois avec deux valeurs : la derniere
    gagne, cas reel du double bloc des campagnes) et CFGEMPTY (cle presente
    mais vide alors que le code a un defaut non vide : Cfg() renvoie le
    vide, cas reel de ContactSkipFolders= dans le modele). S'applique aussi
    a Docs\\config_ini.example.
    """
    keys = collect_config_keys(folder, files)
    known = {k.lower(): v for k, v in keys.items()}
    seen = {}
    for ln, key, val in config_file_entries(cfg_path):
        kl = key.lower()
        if kl in seen and seen[kl][1] != val:
            rep.warn(cfg_path, ln, "CFGDUP",
                     f"cle '{key}' deja presente ligne {seen[kl][0]} avec une autre "
                     f"valeur - c'est la DERNIERE qui compte")
        seen[kl] = (ln, val)
        if kl in known:
            info = known[kl]
            if val == "" and info["nonempty_str"]:
                rep.warn(cfg_path, ln, "CFGEMPTY",
                         f"cle '{key}' vide : remplace le defaut du code "
                         f"({info['nature']} '{info['default'][:30]}') par du vide - "
                         f"la mettre en commentaire (;{key}=) pour garder le defaut")
            continue
        if any(kl.startswith(p.lower()) for p in CFG_DYNAMIC_PREFIXES):
            continue
        rep.warn(cfg_path, ln, "CFGKEY",
                 f"cle '{key}' inconnue du code (typo ou cle obsolete ?)")


def gen_versions(folder, files):
    """--gen-versions : Select Case de modMaintenance.ModuleVersionOf.
    2026-11-16 : modMailAnalytics inclus (F9-12), modules archives exclus
    (F9-14)."""
    cases = []
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or not (base.startswith("mod") or base == "ThisOutlookSession.txt"):
            continue
        try:
            txt, _ = read_source(f)
        except Exception:
            continue
        if is_archived(txt):
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
    """
    --gen-config : template config_ini.example genere depuis le code.
    2026-11-16 (F2-02) : une cle dont le defaut du code n'est pas vide est
    ecrite EN COMMENTAIRE (";Cle=defaut") : une cle presente mais vide
    remplacerait le defaut par du vide. Defaut construit dans le code
    ("Aper" & Chr(231) & "u") ou constante (DEFAULT_SKIP) : jamais recopie,
    seulement signale. Cles numerotees ou construites : documentees en fin.
    Sortie a relire avant de remplacer le bloc genere du modele.
    """
    keys = collect_config_keys(folder, files)
    by_mod = {}
    for k, info in keys.items():
        by_mod.setdefault(info["module"], []).append((k, info))
    print("; config_ini.example - GENERE par validate.py --gen-config")
    print("; Une cle par ligne : Cle=valeur. Valeurs par defaut = celles du code.")
    print("; Cle en commentaire (;Cle=...) : le code a un defaut non vide ; ne la")
    print("; decommenter que pour la remplacer par une vraie valeur (une cle")
    print("; presente mais vide remplace le defaut par du vide).")
    print("[ClassicOutlookToolkit]")
    for mod in sorted(by_mod):
        print(f"\n; --- {mod} ---")
        for k, info in sorted(by_mod[mod], key=lambda x: x[0].lower()):
            nature, d = info["nature"], info["default"]
            if nature == "empty":
                print(f"{k}=")
            elif nature in ("literal", "number", "bool"):
                print(f";{k}={d}")
            elif nature == "constant":
                print(f";   {k} : defaut = {d} (constante ou fonction du code)")
                print(f";{k}=")
            else:
                print(f";   {k} : defaut construit dans le code ({d[:40]}) - ne pas recopier")
                print(f";{k}=")
    print("\n; --- Cles numerotees ou construites (non generees) ---")
    for line in CFG_DYNAMIC_DOC:
        print(f";   {line}")


# =============================================================================
# 2026-11-16 (F0-02) : livraisons refaites sur une ancienne copie
# =============================================================================

def find_changelog(folder):
    """CHANGELOG.txt du dossier des sources ou d'un dossier parent (2 niveaux)."""
    d = os.path.abspath(folder)
    for _ in range(3):
        p = os.path.join(d, "CHANGELOG.txt")
        if os.path.exists(p):
            return p
        d = os.path.dirname(d)
    return None


def changelog_versions(path):
    """{module: (version, date, ligne)} : plus haut numero cite pour chaque
    module dans les entrees datees de VERSION_SINCE ou apres."""
    out = {}
    try:
        txt, _ = read_source(path)
    except Exception:
        return out
    date = None
    for ln, line in enumerate(txt.split("\n"), 1):
        m = RE_CHANGELOG_DATE.match(line)
        if m:
            date = m.group(1)
        if date is None or date < VERSION_SINCE:
            continue
        for mm in RE_CHANGELOG_CITE.finditer(line):
            v = version_tuple(mm.group(2))
            if v is None:
                continue
            k = mm.group(1)
            if k not in out or v > out[k][0]:
                out[k] = (v, date, ln)
    return out


def check_version_vs_changelog(path, txt, rep, cl_versions):
    """
    VERSION (erreur) : MOD_VERSION du fichier inferieur au plus haut numero
    cite pour ce module dans CHANGELOG.txt (depuis VERSION_SINCE) : le
    fichier a ete refait a partir d'une copie anterieure a une livraison
    (cas reel : modFolderSort "v2" du 22/08 sur la v3.1 du 18/08 ; tri des
    dossiers casse pendant trois mois).
    """
    key = module_key(path)
    if not key or key not in cl_versions:
        return
    m = RE_MOD_VERSION.search(txt)
    if not m:
        return
    fv = version_tuple(m.group(1))
    cv, cdate, cline = cl_versions[key]
    if fv is not None and fv < cv:
        rep.error(path, 0, "VERSION",
                  f"MOD_VERSION v{m.group(1)} inferieur a la v{'.'.join(map(str, cv))} "
                  f"citee dans CHANGELOG.txt (ligne {cline}, {cdate}) - livraison "
                  f"refaite sur une ancienne copie ? Repartir du fichier de Code\\ "
                  f"et donner un numero nouveau")


def history_entries(txt):
    """Entrees d'historique de l'en-tete : {(version, date)} (lignes de
    commentaire "' vN.N (AAAA-MM-JJ)" avant la premiere procedure)."""
    out = set()
    for line in txt.split("\n"):
        if RE_PROC.match(line):
            break
        m = RE_HISTORY_ENTRY.match(line)
        if m:
            out.add((m.group(1), m.group(2)))
    return out


def normalized_code(txt):
    """Texte compare pour RECUL : casse et blancs ignores (le VBE change la
    casse des identifiants - Shell / shell, Item / item - et les espaces
    autour de ":" et apres "," quand on colle un module)."""
    return re.sub(r"\s+", "", txt.lower())


def check_against(files, ref_dir, rep):
    """
    --against <dossier> (alias --installe) : compare chaque fichier a son
    homonyme dans un dossier de reference (dernier Code\\ livre, ou modules
    extraits de la sauvegarde du projet installe ; a plat ou range par
    application). RECUL (erreur) :
      - une entree "vN (date)" de l'en-tete de la reference a disparu alors
        qu'elle n'est pas plus ancienne que la plus ancienne entree gardee
        (un historique ancien deplace dans CHANGELOG n'est pas un recul) ;
      - le contenu differe et MOD_VERSION n'a pas augmente.
    Cas reel : modMaintenance "v4.5 du 02/11" batie sur la v4.4, sans
    l'entree "v4.5 (2026-10-25)" du module installe.
    """
    refs = {}
    for d, _dirs, fs in os.walk(ref_dir):
        for f in fs:
            if f.endswith(".txt"):
                refs.setdefault(f.lower(), os.path.join(d, f))
    compared = 0
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or base.lower() not in refs:
            continue
        try:
            new_txt, _ = read_source(f)
            ref_txt, _ = read_source(refs[base.lower()])
        except Exception:
            continue
        compared += 1
        if normalized_code(new_txt) == normalized_code(ref_txt):
            continue
        new_h, ref_h = history_entries(new_txt), history_entries(ref_txt)
        if new_h:
            oldest_kept = min(dt for _v, dt in new_h)
            lost = sorted(e for e in ref_h - new_h if e[1] >= oldest_kept)
            new_versions = {v for v, _dt in new_h}
            for v, dt in lost:
                note = " (meme numero donne a un autre contenu)" if v in new_versions else ""
                rep.error(f, 0, "RECUL",
                          f"entree d'historique v{v} ({dt}) de la reference absente{note} "
                          f"- fichier refait a partir d'une copie plus ancienne ?")
        mn, mr = RE_MOD_VERSION.search(new_txt), RE_MOD_VERSION.search(ref_txt)
        if mn and mr:
            vn, vr = version_tuple(mn.group(1)), version_tuple(mr.group(1))
            if vn is not None and vr is not None and vn <= vr:
                rep.error(f, 0, "RECUL",
                          f"contenu different de la reference mais MOD_VERSION "
                          f"v{mn.group(1)} <= v{mr.group(1)} - donner un numero nouveau "
                          f"(ou fichier plus ancien que la reference)")
    return compared


# =============================================================================
# Pilote
# =============================================================================

CHECKS_TXT = [
    check_encoding,       # signature (path, txt, raw, rep)
]


def validate_file(path, rep, known, symbols=None, enum_members=None,
                  archived=frozenset(), cl_versions=None):
    try:
        txt, raw = read_source(path)
    except Exception as exc:
        rep.error(path, 0, "READ", f"lecture impossible : {exc}")
        return False
    if is_archived(txt):
        return False      # 2026-11-16 (F9-14) : module archive, non controle

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
    check_msg_loop(path, txt, rep)
    if cl_versions:
        check_version_vs_changelog(path, txt, rep, cl_versions)
    if symbols is not None:
        check_xref(path, txt, rep, symbols, enum_members, archived)
    return True


# Rangement du 2026-10-25 : sources rangees par application sous Code\
# (Outlook\, Excel\<classeur>\, Interface\Rubans\ et
# Interface\Formulaires\, Docs\, Outils\) ; ce script vit dans
# Code\Outils\. Sans fichier en argument, on controle tout Code\,
# sous-dossiers compris, sur deux niveaux au plus (Excel\<classeur>\,
# Interface\Rubans\). Docs\ et Outils\ ne contiennent pas de VBA.
# Un depot a plat (GitHub) se controle comme avant, depuis sa racine.
SKIP_DIRS = {"docs", "outils", "tools", "archives"}   # 2026-11-16 : dossier Archives (F9-14)
MAX_DEPTH = 2


def sources_root(start):
    """Dossier des sources a controler quand aucun fichier n'est donne."""
    here = os.path.abspath(start)
    if os.path.basename(here).lower() in ("outils", "tools"):
        parent = os.path.dirname(here)
        if os.path.isdir(os.path.join(parent, "Outlook")):
            return parent                       # lance depuis Code\Outils
    if os.path.isdir(os.path.join(here, "Code", "Outlook")):
        return os.path.join(here, "Code")       # depuis le dossier de l'outil
    return here                                 # Code\ lui-meme, ou depot a plat


def list_sources(root):
    """Fichiers .txt de root et de ses sous-dossiers (MAX_DEPTH niveaux)."""
    out = []
    for d, dirs, files in os.walk(root):
        depth = 0 if d == root else os.path.relpath(d, root).count(os.sep) + 1
        if depth >= MAX_DEPTH:
            dirs[:] = []
        else:
            dirs[:] = sorted(x for x in dirs
                             if not x.startswith(".") and x.lower() not in SKIP_DIRS)
        for f in sorted(files):
            if f.endswith(".txt") and f not in NOT_VBA:
                out.append(os.path.join(d, f))
    return out


def main(argv):
    quiet = "--quiet" in argv
    strict = "--strict" in argv
    cfg_path = None
    if "--config" in argv:
        i = argv.index("--config")
        if i + 1 < len(argv):
            cfg_path = argv[i + 1]
            argv = argv[:i] + argv[i + 2:]
    # 2026-11-16 (F0-02) : dossier de reference pour le controle RECUL.
    ref_dir = None
    for opt in ("--against", "--installe"):
        if opt in argv:
            i = argv.index(opt)
            if i + 1 < len(argv):
                ref_dir = argv[i + 1]
                argv = argv[:i] + argv[i + 2:]
    args = [a for a in argv if not a.startswith("--")]

    if args:
        files = args
        folder = os.path.dirname(os.path.abspath(args[0])) or "."
    else:
        folder = sources_root(".")
        files = list_sources(folder)

    if "--gen-versions" in argv:
        gen_versions(folder, [os.path.join(folder, f) if not os.path.isabs(f) and not os.path.exists(f) else f for f in files])
        return 0
    if "--gen-config" in argv:
        gen_config_example(folder, [os.path.join(folder, f) if not os.path.isabs(f) and not os.path.exists(f) else f for f in files])
        return 0

    known = collect_short_names(folder, files)
    symbols, enums_by_mod = collect_public_symbols(folder, files)
    archived = collect_archived(files)
    # 2026-11-16 (F0-02) : numeros de version cites dans CHANGELOG.txt.
    cl_path = find_changelog(folder)
    cl_versions = changelog_versions(cl_path) if cl_path else {}
    all_enum_members = set()
    for v in enums_by_mod.values():
        all_enum_members |= v
    rep = Report()
    # 2026-10-25 : un meme fichier a deux endroits (Code range par
    # application) = une livraison posee au mauvais endroit.
    seen = {}
    for f in files:
        seen.setdefault(os.path.basename(f).lower(), []).append(f)
    for same in seen.values():
        for f in same[1:]:
            rep.error(f, 0, "DOUBLON",
                      f"meme fichier a deux endroits : "
                      f"{os.path.relpath(same[0], folder)} et "
                      f"{os.path.relpath(f, folder)} - n'en garder qu'un "
                      f"(livraison posee au mauvais endroit ?)")
    for f in files:
        base = os.path.basename(f)
        if base in NOT_VBA or "config_ini" in base or base.startswith("CHANGELOG"):
            continue
        validate_file(f, rep, known, symbols, all_enum_members, archived, cl_versions)

    if cfg_path:
        check_config_file(cfg_path, folder, [os.path.join(folder, f) if not os.path.exists(f) else f for f in files], rep)

    compared = None
    if ref_dir:
        if os.path.isdir(ref_dir):
            compared = check_against(files, ref_dir, rep)
        else:
            rep.error(ref_dir, 0, "RECUL", "dossier de reference introuvable")

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
    if archived:
        print(f"Archive(s) non controle(s) : {', '.join(sorted(archived))}.")
    if cl_path is None:
        print("CHANGELOG.txt introuvable : controle VERSION non fait.")
    if compared is not None:
        print(f"Compare(s) a la reference {ref_dir} : {compared} fichier(s).")
    print("=" * 78)

    if rep.errors:
        return 1
    if strict and rep.warns:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
