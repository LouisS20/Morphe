"""
Morphe - convertisseur d'images, d'audio, de vidéo et de documents + téléchargeur de vidéos.
Tout est embarqué dans l'appli :
  - Pillow       : images
  - FFmpeg       : audio / vidéo (binaire fourni par imageio-ffmpeg)
  - python-docx, reportlab, pypdf, openpyxl, markdown... : documents
  - yt-dlp + Deno (paquet pip "deno") : téléchargement de vidéos
"""
import csv
import html as _html
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image

try:
    import imageio_ffmpeg
    FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
except Exception:  # pragma: no cover
    FFMPEG = None

try:  # HEIC / HEIF / AVIF (photos d'iPhone, etc.)
    import pillow_heif
    pillow_heif.register_heif_opener()
    try:
        pillow_heif.register_avif_opener()
    except Exception:
        pass
except Exception:  # pragma: no cover
    pass

HAS_YAML = importlib.util.find_spec("yaml") is not None
HAS_TOMLW = importlib.util.find_spec("tomli_w") is not None

try:  # rendu des pages PDF en images (optionnel)
    import pymupdf as _fitz
except Exception:  # pragma: no cover
    try:
        import fitz as _fitz
    except Exception:
        _fitz = None
HAS_FITZ = _fitz is not None

try:  # runtime JavaScript embarqué, nécessaire à YouTube
    import deno as _deno
    DENO = _deno.find_deno_bin()
except Exception:  # pragma: no cover
    DENO = None

try:  # glisser-déposer (optionnel)
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:  # pragma: no cover
    HAS_DND = False

# ---------- Formats ----------
IMAGES = {"png", "jpg", "jpeg", "jpe", "jfif", "webp", "bmp", "tif", "tiff", "ico", "gif", "tga", "ppm",
          "pgm", "pbm", "pcx", "dds", "psd", "jp2", "icns", "avif", "heic", "heif"}
AUDIOS = {"mp3", "wav", "flac", "ogg", "oga", "m4a", "m4b", "aac", "wma", "opus", "aiff", "aif", "mka",
          "amr", "ac3", "mp2", "caf", "au", "3ga", "ape", "wv", "weba"}
VIDEOS = {"mp4", "mkv", "avi", "mov", "webm", "flv", "wmv", "m4v", "mpg", "mpeg", "3gp", "3g2", "ts",
          "mts", "m2ts", "ogv", "vob", "asf", "f4v", "rm", "rmvb", "divx", "mxf", "m2v"}
TABLEAUX = ("csv", "tsv", "xlsx", "json", "yaml", "yml", "toml")
DOCS = {"txt", "md", "markdown", "html", "htm", "docx", "pdf"} | set(TABLEAUX)

SORTIES = {
    "image": ["PNG", "JPG", "WEBP", "ICO", "BMP", "TIFF", "GIF", "PDF"],
    "audio": ["MP3", "WAV", "FLAC", "OGG", "M4A", "AAC", "OPUS"],
    "video": ["MP4", "MKV", "WEBM", "AVI", "MOV", "TS", "GIF", "MP3", "WAV", "M4A", "FLAC"],
}

# ---------- Palette : fond sombre transparent, accent blanc ----------
BG = "#0f0f17"
CARD = "#181824"
CARD2 = "#212133"
HOVER = "#2d2d45"
ACCENT = "#f4f4f5"
ACCENT_HOVER = "#d4d4d8"
ACCENT_TEXT = "#0f0f17"
MUTED = "#8b8ba7"
COULEUR = {"image": "#fb923c", "audio": "#34d399", "video": "#38bdf8", "doc": "#d4d4d8"}


def ext_de(path):
    return Path(path).suffix.lower().lstrip(".")


def type_de(path):
    e = ext_de(path)
    if e in IMAGES:
        return "image"
    if e in AUDIOS:
        return "audio"
    if e in VIDEOS:
        return "video"
    if e in DOCS:
        return "doc"
    return None


def trier_fichiers(chemins):
    """Développe les dossiers, sépare les fichiers reconnus des autres (extensions ignorées)."""
    tous = []
    for c in chemins:
        if os.path.isdir(c):
            for racine, _, noms in os.walk(c):
                tous += [os.path.join(racine, n) for n in sorted(noms)]
        else:
            tous.append(c)
    acceptes, ignores = [], []
    for c in tous:
        if not os.path.isfile(c):
            continue
        if type_de(c):
            acceptes.append(c)
        else:
            ignores.append(ext_de(c))
    return acceptes, ignores


def formats_pour(path):
    t, e = type_de(path), ext_de(path)
    if t == "doc":
        if e in TABLEAUX:
            propre = {"yml": "YAML"}.get(e, e.upper())
            sorties = (["CSV", "TSV", "XLSX", "JSON"] + (["YAML"] if HAS_YAML else [])
                       + (["TOML"] if HAS_TOMLW else []) + ["HTML", "PDF"])
            return [f for f in sorties if f != propre]
        if e == "pdf":
            return ["TXT", "DOCX", "HTML", "MD"] + (["PNG"] if HAS_FITZ else [])
        propre = {"htm": "HTML", "markdown": "MD"}.get(e, e.upper())
        return [f for f in ["PDF", "DOCX", "HTML", "MD", "TXT"] if f != propre]
    return SORTIES[t]


def ressource(nom):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nom)


def sortie_libre(dossier, stem, ext):
    out = Path(dossier) / f"{stem}.{ext}"
    i = 1
    while out.exists():
        out = Path(dossier) / f"{stem}_{i}.{ext}"
        i += 1
    return out


def taille_lisible(octets):
    for u in ("o", "Ko", "Mo", "Go"):
        if octets < 1024 or u == "Go":
            return f"{octets:.0f} {u}" if u == "o" else f"{octets:.1f} {u}"
        octets /= 1024


def ouvrir_dossier(chemin):
    try:
        if os.name == "nt":
            os.startfile(chemin)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", chemin])
        else:
            subprocess.Popen(["xdg-open", chemin])
    except Exception:
        pass


# =====================================================================
#  IMAGES
# =====================================================================
def convertir_image(src, fmt, out, qualite):
    try:
        img = Image.open(src)
        img.load()
    except Exception:
        raise RuntimeError("Image illisible. Pour HEIC / AVIF, il faut le module pillow-heif "
                           "(python -m pip install pillow-heif).")
    try:
        from PIL import ImageOps
        img = ImageOps.exif_transpose(img)  # respecte l'orientation des photos de téléphone
    except Exception:
        pass
    if img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
        img = img.convert("RGBA" if "A" in img.mode else "RGB")
    fmt = fmt.lower()
    if fmt in ("jpg", "bmp", "pdf"):
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGBA")
            fond = Image.new("RGB", img.size, (255, 255, 255))
            fond.paste(img, mask=img.split()[-1])
            img = fond
        else:
            img = img.convert("RGB")
    if fmt == "jpg":
        img.save(out, "JPEG", quality=qualite, optimize=True)
    elif fmt == "webp":
        img.save(out, "WEBP", quality=qualite)
    elif fmt == "pdf":
        img.save(out, "PDF", resolution=100.0)
    elif fmt == "ico":
        img = img.convert("RGBA")
        cote = max(img.size)
        carre = Image.new("RGBA", (cote, cote), (0, 0, 0, 0))
        carre.paste(img, ((cote - img.width) // 2, (cote - img.height) // 2))
        tailles = [(s, s) for s in (16, 24, 32, 48, 64, 128, 256) if s <= max(cote, 16)] or [(256, 256)]
        carre.save(out, "ICO", sizes=tailles)
    elif fmt == "tiff":
        img.save(out, "TIFF")
    elif fmt == "gif":
        img.convert("P", palette=Image.ADAPTIVE).save(out, "GIF")
    else:
        img.save(out, "PNG")


# =====================================================================
#  AUDIO / VIDEO
# =====================================================================
RE_DUREE = re.compile(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)")
RE_TEMPS = re.compile(r"time=(\d+):(\d+):(\d+(?:\.\d+)?)")


def vers_secondes(m):
    return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3))


def _executer_ffmpeg(cmd, rappel_progres):
    """Lance FFmpeg sans fenêtre ; renvoie (code de retour, dernière ligne utile)."""
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.Popen(cmd, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL,
                            stdin=subprocess.DEVNULL, creationflags=flags)
    etat = {"duree": None, "derniere": ""}

    def traiter(ligne):
        if ligne.strip():
            etat["derniere"] = ligne
        if etat["duree"] is None:
            d = RE_DUREE.search(ligne)
            if d:
                etat["duree"] = vers_secondes(d)
        t = RE_TEMPS.search(ligne)
        if t and etat["duree"]:
            rappel_progres(min(vers_secondes(t) / etat["duree"], 1.0))

    buf = b""
    while True:
        bloc = proc.stderr.read(256)
        if not bloc:
            break
        buf += bloc
        morceaux = re.split(rb"[\r\n]", buf)
        buf = morceaux.pop()
        for m in morceaux:
            traiter(m.decode("utf-8", "ignore"))
    if buf:
        traiter(buf.decode("utf-8", "ignore"))
    return proc.wait(), etat["derniere"]


def convertir_ffmpeg(src, fmt, out, rappel_progres):
    exes = list(dict.fromkeys(e for e in (FFMPEG, shutil.which("ffmpeg")) if e))
    if not exes:
        raise RuntimeError("Le moteur audio/vidéo n'a pas pu être chargé.")
    fmt = fmt.lower()
    args = ["-y", "-i", str(src)]
    if type_de(src) == "video" and fmt in AUDIOS:
        args += ["-vn"]
    elif fmt == "mp4":
        args += ["-movflags", "+faststart"]
    elif fmt == "ts":
        args += ["-c:v", "libx264", "-c:a", "aac", "-f", "mpegts"]  # MPEG-TS en H.264 + AAC
    elif fmt == "gif":
        args += ["-vf", "fps=12,scale=480:-1:flags=lanczos"]
    args.append(str(out))

    message = ""
    for exe in exes:
        code, derniere = _executer_ffmpeg([exe] + args, rappel_progres)
        if code == 0:
            return
        plante = code < 0 or code >= 0xC0000000  # le moteur a planté (et pas une vraie erreur de fichier)
        message = ("Le moteur FFmpeg a planté en lisant ce fichier." if plante else derniere) or "Erreur FFmpeg"
        try:
            if os.path.exists(out):
                os.remove(out)
        except OSError:
            pass
        if not plante:
            break  # vraie erreur (fichier invalide...) : inutile de réessayer avec un autre moteur
    raise RuntimeError(message)


# =====================================================================
#  DOCUMENTS  (tout passe par une liste de "blocs" intermédiaire)
#  bloc = {"t": "h"|"p"|"li"|"pre"|"quote"|"table", ...}
#  Le texte "x" contient un balisage minimal : <b> <i> <code> <a href=""> <br/>
# =====================================================================
TOK = re.compile(r'(<b>|</b>|<i>|</i>|<code>|</code>|<br/>|<a href="[^"]*">|</a>)')


def esc(t):
    return _html.escape(t, quote=False)


def lire_texte(path):
    brut = Path(path).read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return brut.decode(enc)
        except UnicodeDecodeError:
            continue
    return brut.decode("utf-8", "replace")


# ----- lecture -----
def txt_vers_blocs(texte):
    blocs = []
    for para in re.split(r"\n\s*\n", texte.replace("\r\n", "\n").replace("\r", "\n")):
        if para.strip():
            blocs.append({"t": "p", "x": "<br/>".join(esc(l.rstrip()) for l in para.strip("\n").split("\n"))})
    return blocs


def html_vers_blocs(code):
    from bs4 import BeautifulSoup, Comment, NavigableString, Tag
    soup = BeautifulSoup(code, "html.parser")
    for t in soup(["script", "style", "head"]):
        t.decompose()
    racine = soup.body or soup
    blocs = []
    TITRES = {"h1", "h2", "h3", "h4", "h5", "h6"}
    CONTENEURS = {"div", "section", "article", "main", "header", "footer", "body", "html",
                  "figure", "nav", "aside", "form", "center", "span_block"}

    def inline(n):
        if isinstance(n, Comment):
            return ""
        if isinstance(n, NavigableString):
            return esc(re.sub(r"\s+", " ", str(n)))
        inter = "".join(inline(c) for c in n.children)
        nom = n.name
        if nom in ("b", "strong"):
            return f"<b>{inter}</b>"
        if nom in ("i", "em"):
            return f"<i>{inter}</i>"
        if nom in ("code", "kbd", "samp", "tt"):
            return f"<code>{inter}</code>"
        if nom == "a" and n.get("href"):
            return f'<a href="{_html.escape(n.get("href"), quote=True)}">{inter}</a>'
        if nom == "br":
            return "<br/>"
        return inter

    def contenu(n):
        return "".join(inline(c) for c in n.children).strip()

    def vide(x):
        return not re.sub(r"<[^>]+>|&\w+;", "", x).strip()

    def lister(liste, ordre, niv):
        for li in liste.find_all("li", recursive=False):
            morceaux = [inline(c) for c in li.children
                        if not (isinstance(c, Tag) and c.name in ("ul", "ol"))]
            x = "".join(morceaux).strip()
            blocs.append({"t": "li", "x": x, "ordre": ordre, "num": 0, "niv": niv})
            for sous in li.find_all(["ul", "ol"], recursive=False):
                lister(sous, sous.name == "ol", niv + 1)

    def parcourir(parent):
        tampon = []

        def vider():
            x = "".join(tampon).strip()
            if x and not vide(x):
                blocs.append({"t": "p", "x": x})
            tampon.clear()

        for c in parent.children:
            if isinstance(c, Comment):
                continue
            if isinstance(c, NavigableString):
                tampon.append(inline(c))
                continue
            n = c.name
            if n in TITRES:
                vider()
                x = contenu(c)
                if not vide(x):
                    blocs.append({"t": "h", "n": int(n[1]), "x": x})
            elif n == "p":
                vider()
                x = contenu(c)
                if not vide(x):
                    blocs.append({"t": "p", "x": x})
            elif n in ("ul", "ol"):
                vider()
                lister(c, n == "ol", 0)
            elif n == "pre":
                vider()
                blocs.append({"t": "pre", "x": c.get_text()})
            elif n == "blockquote":
                vider()
                x = contenu(c)
                if not vide(x):
                    blocs.append({"t": "quote", "x": x})
            elif n == "table":
                vider()
                rows = [[contenu(td) for td in tr.find_all(["td", "th"])] for tr in c.find_all("tr")]
                rows = [r for r in rows if r]
                if rows:
                    blocs.append({"t": "table", "rows": rows})
            elif n in CONTENEURS:
                vider()
                parcourir(c)
            elif n == "br":
                tampon.append("<br/>")
            elif n == "hr":
                vider()
            else:
                tampon.append(inline(c))
        vider()

    parcourir(racine)
    return blocs


def md_vers_blocs(texte):
    import markdown
    code = markdown.markdown(texte, extensions=["tables", "fenced_code", "sane_lists"])
    return html_vers_blocs(code)


def docx_vers_blocs(path):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(str(path))
    blocs = []

    def inline_par(p):
        sortie = []
        for r in p.runs:
            t = esc(r.text).replace("\n", "<br/>").replace("\t", "    ")
            if not t:
                continue
            if r.font.name in ("Consolas", "Courier New", "Courier"):
                t = f"<code>{t}</code>"
            if r.bold:
                t = f"<b>{t}</b>"
            if r.italic:
                t = f"<i>{t}</i>"
            sortie.append(t)
        return "".join(sortie).strip()

    for el in d.element.body.iterchildren():
        if el.tag.endswith("}p"):
            p = Paragraph(el, d)
            x = inline_par(p)
            if not x:
                continue
            try:
                nom = p.style.name or ""
            except Exception:
                nom = ""
            if nom == "Title":
                blocs.append({"t": "h", "n": 1, "x": x})
            elif nom.startswith("Heading"):
                m = re.search(r"(\d+)", nom)
                blocs.append({"t": "h", "n": min(int(m.group(1)) if m else 1, 6), "x": x})
            elif nom.startswith("List"):
                m = re.search(r"(\d+)$", nom)
                blocs.append({"t": "li", "x": x, "ordre": "Number" in nom, "num": 0,
                              "niv": (int(m.group(1)) - 1) if m else 0})
            elif nom in ("Quote", "Intense Quote"):
                blocs.append({"t": "quote", "x": x})
            else:
                blocs.append({"t": "p", "x": x})
        elif el.tag.endswith("}tbl"):
            t = Table(el, d)
            rows = [[esc(c.text.strip()).replace("\n", "<br/>") for c in row.cells] for row in t.rows]
            if rows:
                blocs.append({"t": "table", "rows": rows})
    return blocs


def pdf_vers_blocs(path):
    from pypdf import PdfReader
    lec = PdfReader(str(path))
    if lec.is_encrypted:
        try:
            lec.decrypt("")
        except Exception:
            raise RuntimeError("Ce PDF est protégé par un mot de passe.")
    blocs = []
    for page in lec.pages:
        courant = []
        for ligne in (page.extract_text() or "").replace("\r", "").split("\n"):
            l = ligne.strip()
            if not l:
                if courant:
                    blocs.append({"t": "p", "x": " ".join(courant)})
                    courant = []
                continue
            courant.append(esc(l))
            if re.search(r"[.!?:;»”)]$", l) or len(l) < 45:
                blocs.append({"t": "p", "x": " ".join(courant)})
                courant = []
        if courant:
            blocs.append({"t": "p", "x": " ".join(courant)})
    if not blocs:
        raise RuntimeError("Aucun texte trouvé dans ce PDF (document scanné ?).")
    return blocs


def pdf_vers_png(path, out):
    if not HAS_FITZ:
        raise RuntimeError("PDF -> PNG nécessite le module pymupdf (non installé).")
    doc = _fitz.open(str(path))
    premier = None
    for n, page in enumerate(doc, 1):
        cible = out.with_name(f"{out.stem}_p{n}.png")
        page.get_pixmap(dpi=150).save(str(cible))
        premier = premier or cible
    doc.close()
    return premier


def cellule(v):
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False, default=str)
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def donnees_vers_table(data):
    if data is None:
        data = []
    if isinstance(data, dict):
        listes = [v for v in data.values() if isinstance(v, list) and v and all(isinstance(r, dict) for r in v)]
        if len(data) == 1 and listes:
            data = listes[0]
        else:
            data = [{"clé": k, "valeur": v} for k, v in data.items()]
    if not isinstance(data, list):
        data = [data]
    if data and all(isinstance(r, dict) for r in data):
        cles = []
        for r in data:
            for k in r:
                if k not in cles:
                    cles.append(k)
        return [cles] + [[cellule(r.get(k, "")) for k in cles] for r in data]
    if data and all(isinstance(r, (list, tuple)) for r in data):
        return [[cellule(c) for c in r] for r in data]
    return [[cellule(v)] for v in data]


def rows_vers_objets(rows):
    if len(rows) < 2:
        return rows
    entete = [str(h) for h in rows[0]]
    return [{entete[i]: nombre_si_possible(cellule(v)) for i, v in enumerate(r) if i < len(entete)}
            for r in rows[1:]]


def lire_table(path, e):
    if e == "tsv":
        return [r for r in csv.reader(io.StringIO(lire_texte(path)), delimiter="\t")]
    if e == "csv":
        texte = lire_texte(path)
        try:
            dial = csv.Sniffer().sniff(texte[:4096], delimiters=",;\t|")
        except csv.Error:
            dial = csv.excel
        return [r for r in csv.reader(io.StringIO(texte), dial)]
    if e == "xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(str(path), read_only=True, data_only=True)
        rows = [["" if v is None else v for v in r] for r in wb.active.iter_rows(values_only=True)]
        wb.close()
        return rows
    if e == "json":
        data = json.loads(lire_texte(path))
    elif e in ("yaml", "yml"):
        try:
            import yaml
        except Exception:
            raise RuntimeError("Le module PyYAML n'est pas installé (python -m pip install pyyaml).")
        data = yaml.safe_load(lire_texte(path))
    else:  # toml
        try:
            import tomllib
        except Exception:
            import tomli as tomllib
        data = tomllib.loads(lire_texte(path))
    return donnees_vers_table(data)


def nombre_si_possible(c):
    if isinstance(c, str):
        s = c.strip()
        if re.fullmatch(r"-?[1-9]\d{0,14}", s) or s == "0":
            return int(s)
        if re.fullmatch(r"-?\d{1,15}[.,]\d+", s):
            return float(s.replace(",", "."))
    return c


def ecrire_table(rows, fmt, out):
    if fmt in ("csv", "tsv"):
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter="\t" if fmt == "tsv" else ";").writerows(
                [[cellule(c) for c in r] for r in rows])
    elif fmt == "xlsx":
        from openpyxl import Workbook
        from openpyxl.styles import Font
        wb = Workbook()
        ws = wb.active
        for r in rows:
            ws.append([nombre_si_possible(cellule(c)) for c in r])
        if len(rows) > 1:
            for c in ws[1]:
                c.font = Font(bold=True)
        for col in ws.columns:
            largeur = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(largeur + 2, 60)
        wb.save(str(out))
    elif fmt == "json":
        Path(out).write_text(json.dumps(rows_vers_objets(rows), ensure_ascii=False, indent=2), encoding="utf-8")
    elif fmt == "yaml":
        try:
            import yaml
        except Exception:
            raise RuntimeError("Le module PyYAML n'est pas installé (python -m pip install pyyaml).")
        Path(out).write_text(yaml.safe_dump(rows_vers_objets(rows), allow_unicode=True, sort_keys=False),
                             encoding="utf-8")
    elif fmt == "toml":
        try:
            import tomli_w
        except Exception:
            raise RuntimeError("Le module tomli-w n'est pas installé (python -m pip install tomli-w).")
        objets = rows_vers_objets(rows)
        if not objets or not all(isinstance(o, dict) for o in objets):
            raise RuntimeError("Il faut au moins une ligne d'en-tête et une ligne de données pour le TOML.")
        propres = [{k: ("" if v is None else v) for k, v in o.items()} for o in objets]
        Path(out).write_text(tomli_w.dumps({"items": propres}), encoding="utf-8")


# ----- écriture -----
def numeroter(blocs):
    cpt = {}
    for b in blocs:
        if b["t"] != "li":
            cpt.clear()
            continue
        niv = b["niv"]
        for k in [k for k in cpt if k > niv]:
            del cpt[k]
        if b["ordre"]:
            cpt[niv] = cpt.get(niv, 0) + 1
            b["num"] = cpt[niv]
        else:
            cpt.pop(niv, None)
            b["num"] = 0


def inline_vers_texte(x):
    def lien(m):
        url = _html.unescape(m.group(1))
        return f"{m.group(2)} ({url})" if url and url != m.group(2) else m.group(2)
    x = re.sub(r'<a href="([^"]*)">(.*?)</a>', lien, x)
    x = x.replace("<br/>", "\n")
    x = re.sub(r"</?(b|i|code|a)(\s[^>]*)?>", "", x)
    return _html.unescape(x)


def inline_vers_md(x):
    x = re.sub(r'<a href="([^"]*)">(.*?)</a>', lambda m: f"[{m.group(2)}]({_html.unescape(m.group(1))})", x)
    x = re.sub(r"<b>(.*?)</b>", r"**\1**", x)
    x = re.sub(r"<i>(.*?)</i>", r"*\1*", x)
    x = re.sub(r"<code>(.*?)</code>", r"`\1`", x)
    x = x.replace("<br/>", "  \n")
    return _html.unescape(x)


CSS_HTML = ("body{font-family:system-ui,'Segoe UI',Arial,sans-serif;max-width:820px;margin:40px auto;"
            "padding:0 20px;line-height:1.65;color:#222}table{border-collapse:collapse;margin:1em 0}"
            "td,th{border:1px solid #ccc;padding:6px 10px;vertical-align:top}th{background:#f3f3f3}"
            "pre{background:#f6f6f6;padding:12px;overflow:auto}"
            "blockquote{border-left:4px solid #ddd;margin-left:0;padding-left:16px;color:#555}")


def ecrire_html(blocs, out, titre):
    numeroter(blocs)
    sortie, liste = [], None
    for b in blocs:
        t = b["t"]
        if t != "li" and liste:
            sortie.append(f"</{liste}>")
            liste = None
        if t == "h":
            sortie.append(f"<h{b['n']}>{b['x']}</h{b['n']}>")
        elif t == "p":
            sortie.append(f"<p>{b['x']}</p>")
        elif t == "li":
            tag = "ol" if b["ordre"] else "ul"
            if liste != tag:
                if liste:
                    sortie.append(f"</{liste}>")
                sortie.append(f"<{tag}>")
                liste = tag
            sortie.append(f'<li style="margin-left:{b["niv"] * 24}px">{b["x"]}</li>')
        elif t == "pre":
            sortie.append(f"<pre>{esc(b['x'])}</pre>")
        elif t == "quote":
            sortie.append(f"<blockquote><p>{b['x']}</p></blockquote>")
        elif t == "table":
            lignes = []
            for i, r in enumerate(b["rows"]):
                cell = "th" if (i == 0 and len(b["rows"]) > 1) else "td"
                lignes.append("<tr>" + "".join(f"<{cell}>{c if isinstance(c, str) else esc(str(c))}</{cell}>" for c in r) + "</tr>")
            sortie.append("<table>" + "".join(lignes) + "</table>")
    if liste:
        sortie.append(f"</{liste}>")
    page = (f'<!DOCTYPE html>\n<html lang="fr"><head><meta charset="utf-8"><title>{esc(titre)}</title>'
            f"<style>{CSS_HTML}</style></head><body>\n" + "\n".join(sortie) + "\n</body></html>")
    Path(out).write_text(page, encoding="utf-8")


def ecrire_txt(blocs, out):
    numeroter(blocs)
    sortie = []
    for b in blocs:
        t = b["t"]
        if t == "h":
            sortie.append(inline_vers_texte(b["x"]).upper() if b["n"] == 1 else inline_vers_texte(b["x"]))
        elif t == "p":
            sortie.append(inline_vers_texte(b["x"]))
        elif t == "li":
            marque = f"{b['num']}." if b["ordre"] else "-"
            sortie.append("  " * b["niv"] + f"{marque} " + inline_vers_texte(b["x"]))
        elif t == "pre":
            sortie.append(b["x"].rstrip("\n"))
        elif t == "quote":
            sortie.append("> " + inline_vers_texte(b["x"]))
        elif t == "table":
            sortie.append("\n".join(" | ".join(inline_vers_texte(c) if isinstance(c, str) else str(c) for c in r)
                                    for r in b["rows"]))
    # les éléments de liste consécutifs restent collés, le reste est séparé par une ligne vide
    texte, precedent = "", None
    for b, s in zip(blocs, sortie):
        texte += ("\n" if (precedent == "li" and b["t"] == "li") else ("\n\n" if texte else "")) + s
        precedent = b["t"]
    Path(out).write_text(texte + "\n", encoding="utf-8")


def ecrire_md(blocs, out):
    numeroter(blocs)
    morceaux, precedent = [], None
    for b in blocs:
        t = b["t"]
        if t == "h":
            s = "#" * b["n"] + " " + inline_vers_md(b["x"])
        elif t == "p":
            s = inline_vers_md(b["x"])
        elif t == "li":
            marque = f"{b['num']}." if b["ordre"] else "-"
            s = "  " * b["niv"] + f"{marque} " + inline_vers_md(b["x"])
        elif t == "pre":
            s = "```\n" + b["x"].rstrip("\n") + "\n```"
        elif t == "quote":
            s = "> " + inline_vers_md(b["x"])
        else:
            rows = [[inline_vers_md(c).replace("|", "\\|").replace("\n", " ") if isinstance(c, str) else str(c)
                     for c in r] for r in b["rows"]]
            n = max(len(r) for r in rows)
            rows = [r + [""] * (n - len(r)) for r in rows]
            s = "| " + " | ".join(rows[0]) + " |\n| " + " | ".join(["---"] * n) + " |"
            for r in rows[1:]:
                s += "\n| " + " | ".join(r) + " |"
        sep = "\n" if (precedent == "li" and t == "li") else ("\n\n" if morceaux else "")
        morceaux.append(sep + s)
        precedent = t
    Path(out).write_text("".join(morceaux) + "\n", encoding="utf-8")


def ajouter_runs(par, x):
    gras = ital = code = False
    for m in TOK.split(x):
        if not m:
            continue
        if m == "<b>":
            gras = True
        elif m == "</b>":
            gras = False
        elif m == "<i>":
            ital = True
        elif m == "</i>":
            ital = False
        elif m == "<code>":
            code = True
        elif m == "</code>":
            code = False
        elif m == "<br/>":
            par.add_run().add_break()
        elif m.startswith("<a ") or m == "</a>":
            pass
        else:
            r = par.add_run(_html.unescape(m))
            r.bold = gras or None
            r.italic = ital or None
            if code:
                r.font.name = "Consolas"


def ecrire_docx(blocs, out):
    from docx import Document
    from docx.shared import Pt
    numeroter(blocs)
    d = Document()
    for b in blocs:
        t = b["t"]
        if t == "h":
            ajouter_runs(d.add_heading("", level=min(b["n"], 6)), b["x"])
        elif t == "p":
            ajouter_runs(d.add_paragraph(), b["x"])
        elif t == "li":
            if b["ordre"]:
                p = d.add_paragraph()
                p.paragraph_format.left_indent = Pt(22 + 18 * b["niv"])
                p.paragraph_format.first_line_indent = Pt(-16)
                p.add_run(f"{b['num']}. ")
            else:
                nom = "List Bullet" + (f" {min(b['niv'], 2) + 1}" if b["niv"] else "")
                try:
                    p = d.add_paragraph(style=nom)
                except KeyError:
                    p = d.add_paragraph(style="List Bullet")
            ajouter_runs(p, b["x"])
        elif t == "pre":
            p = d.add_paragraph()
            r = p.add_run(b["x"].rstrip("\n"))
            r.font.name = "Consolas"
            r.font.size = Pt(9)
        elif t == "quote":
            try:
                p = d.add_paragraph(style="Quote")
            except KeyError:
                p = d.add_paragraph()
            ajouter_runs(p, b["x"])
        elif t == "table":
            rows = b["rows"]
            n = max(len(r) for r in rows)
            tbl = d.add_table(rows=len(rows), cols=n)
            try:
                tbl.style = "Table Grid"
            except Exception:
                pass
            for i, r in enumerate(rows):
                for j, c in enumerate(r):
                    cp = tbl.cell(i, j).paragraphs[0]
                    ajouter_runs(cp, c if isinstance(c, str) else esc(str(c)))
                    if i == 0 and len(rows) > 1:
                        for run in cp.runs:
                            run.bold = True
            d.add_paragraph()
    d.save(str(out))


_POLICE = None


def police_pdf():
    global _POLICE
    if _POLICE:
        return _POLICE
    try:
        import reportlab
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        candidats = []
        win = os.environ.get("WINDIR")
        if win:
            dossier = os.path.join(win, "Fonts")
            candidats.append([os.path.join(dossier, n) for n in ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf")])
        rl = os.path.join(os.path.dirname(reportlab.__file__), "fonts")
        candidats.append([os.path.join(rl, n) for n in ("Vera.ttf", "VeraBd.ttf", "VeraIt.ttf", "VeraBI.ttf")])
        for fichiers in candidats:
            if all(os.path.isfile(f) for f in fichiers):
                try:
                    for nom, f in zip(("MorpheN", "MorpheB", "MorpheI", "MorpheBI"), fichiers):
                        pdfmetrics.registerFont(TTFont(nom, f))
                    pdfmetrics.registerFontFamily("MorpheN", normal="MorpheN", bold="MorpheB",
                                                  italic="MorpheI", boldItalic="MorpheBI")
                    _POLICE = ("MorpheN", "MorpheB")
                    return _POLICE
                except Exception:
                    continue
    except Exception:
        pass
    _POLICE = ("Helvetica", "Helvetica-Bold")
    return _POLICE


def ecrire_pdf(blocs, out, titre):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    numeroter(blocs)
    f, fb = police_pdf()
    base = ParagraphStyle("base", parent=getSampleStyleSheet()["Normal"], fontName=f,
                          fontSize=10.5, leading=15, spaceAfter=6)
    tailles = {1: 22, 2: 18, 3: 15, 4: 13, 5: 12, 6: 11}
    titres = {n: ParagraphStyle(f"h{n}", parent=base, fontName=fb, fontSize=s, leading=s + 5,
                                spaceBefore=10, spaceAfter=6) for n, s in tailles.items()}
    code = ParagraphStyle("code", parent=base, fontName="Courier", fontSize=9, leading=12,
                          backColor=colors.HexColor("#f3f3f3"), borderPadding=6, spaceAfter=10)
    citation = ParagraphStyle("quote", parent=base, leftIndent=18, textColor=colors.HexColor("#555555"))
    cell_style = ParagraphStyle("cell", parent=base, fontSize=9, leading=12, spaceAfter=0)
    largeur = A4[0] - 4 * cm

    def par(x, style, **kw):
        x2 = x.replace("<code>", '<font face="Courier">').replace("</code>", "</font>")
        try:
            return Paragraph(x2, style, **kw)
        except Exception:
            return Paragraph(esc(inline_vers_texte(x)), style, **kw)

    flux = []
    for b in blocs:
        t = b["t"]
        if t == "h":
            flux.append(par(b["x"], titres[min(b["n"], 6)]))
        elif t == "p":
            flux.append(par(b["x"], base))
        elif t == "li":
            st = ParagraphStyle("li", parent=base, leftIndent=18 + b["niv"] * 16,
                                bulletIndent=b["niv"] * 16, spaceAfter=3)
            flux.append(par(b["x"], st, bulletText=f"{b['num']}." if b["ordre"] else "•"))
        elif t == "pre":
            flux.append(par(esc(b["x"].rstrip("\n")).replace("\n", "<br/>").replace("  ", "&nbsp;&nbsp;"), code))
        elif t == "quote":
            flux.append(par(b["x"], citation))
        elif t == "table":
            rows = b["rows"]
            n = max(len(r) for r in rows)
            data = [[par(c if isinstance(c, str) else esc(str(c)), cell_style) for c in r + [""] * (n - len(r))]
                    for r in rows]
            tbl = Table(data, colWidths=[largeur / n] * n, repeatRows=1)
            tbl.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bbbbbb")),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
            flux.append(tbl)
            flux.append(Spacer(1, 8))
    if not flux:
        flux.append(Spacer(1, 12))
    SimpleDocTemplate(str(out), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                      topMargin=2 * cm, bottomMargin=2 * cm, title=titre).build(flux)


def ecrire_blocs(blocs, fmt, out, titre):
    if not blocs:
        raise RuntimeError("Le document est vide.")
    if fmt == "html":
        ecrire_html(blocs, out, titre)
    elif fmt == "txt":
        ecrire_txt(blocs, out)
    elif fmt == "md":
        ecrire_md(blocs, out)
    elif fmt == "docx":
        ecrire_docx(blocs, out)
    elif fmt == "pdf":
        ecrire_pdf(blocs, out, titre)
    else:
        raise RuntimeError(f"Format de sortie non pris en charge : .{fmt}")


def convertir_document(src, fmt, out):
    e, f = ext_de(src), fmt.lower()
    titre = Path(src).stem
    if e in TABLEAUX:
        rows = lire_table(src, e)
        if not rows:
            raise RuntimeError("Le fichier ne contient aucune donnée.")
        if f in ("csv", "tsv", "xlsx", "json", "yaml", "toml"):
            ecrire_table(rows, f, out)
        else:
            ecrire_blocs([{"t": "table", "rows": [[esc(str(cellule(c))) for c in r] for r in rows]}], f, out, titre)
        return
    if e == "pdf":
        if f == "png":
            pdf_vers_png(src, out)
        else:
            ecrire_blocs(pdf_vers_blocs(src), f, out, titre)
        return
    if e == "docx":
        blocs = docx_vers_blocs(src)
    elif e in ("html", "htm"):
        blocs = html_vers_blocs(lire_texte(src))
    elif e in ("md", "markdown"):
        blocs = md_vers_blocs(lire_texte(src))
    else:
        blocs = txt_vers_blocs(lire_texte(src))
    ecrire_blocs(blocs, f, out, titre)


def convertir_un(src, fmt, dossier, qualite, rappel):
    t, f = type_de(src), fmt.lower()
    out = sortie_libre(dossier, Path(src).stem, f)
    if t == "image":
        convertir_image(src, f, out, qualite)
    elif t == "doc":
        convertir_document(src, f, out)
    else:
        convertir_ffmpeg(src, f, out, rappel)
    return out


# =====================================================================
#  TÉLÉCHARGEMENT (yt-dlp)
# =====================================================================
def telecharger(url, mode, hauteur, dossier, rappel, annule):
    try:
        import yt_dlp
    except Exception:
        raise RuntimeError("Le module yt-dlp n'est pas installé : python -m pip install \"yt-dlp[default]\" deno")
    Annulation = getattr(yt_dlp.utils, "DownloadCancelled", Exception)

    def hook(d):
        if annule.is_set():
            raise Annulation()
        if d.get("status") == "downloading":
            tot = d.get("total_bytes") or d.get("total_bytes_estimate")
            fait = d.get("downloaded_bytes") or 0
            p = (fait / tot) if tot else None
            morceaux = [f"{int(p * 100)} %"] if p is not None else ["téléchargement…"]
            if d.get("speed"):
                morceaux.append(taille_lisible(d["speed"]) + "/s")
            if d.get("eta") is not None:
                morceaux.append(f"{int(d['eta'])} s restantes")
            rappel(p, " · ".join(morceaux))
        elif d.get("status") == "finished":
            rappel(1.0, "finalisation…")

    opts = {
        "outtmpl": os.path.join(dossier, "%(title).150B.%(ext)s"),
        "noplaylist": True, "quiet": True, "no_warnings": True, "noprogress": True,
        "windowsfilenames": True, "progress_hooks": [hook],
        "retries": 5, "fragment_retries": 5, "concurrent_fragment_downloads": 4,
    }
    if FFMPEG:
        opts["ffmpeg_location"] = FFMPEG
    if DENO:
        opts["js_runtimes"] = {"deno": {"path": DENO}}
    if mode == "audio":
        opts["format"] = "bestaudio/best"
        opts["postprocessors"] = [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3", "preferredquality": "192"}]
    else:
        # on privilégie H.264 + AAC : lisible partout (donc "Max" = meilleur H.264, ~1080p sur YouTube)
        opts["format_sort"] = ["vcodec:h264", "res", "acodec:aac"]
        if hauteur:
            opts["format"] = f"bv*[height<={hauteur}]+ba/b[height<={hauteur}]/bv*+ba/b"
        else:
            opts["format"] = "bv*+ba/b"
        opts["merge_output_format"] = "mp4"
    try:
        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(url, download=True)
    except Annulation:
        raise InterruptedError()
    except Exception as e:
        if annule.is_set():
            raise InterruptedError()
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).strip()
        raise RuntimeError(msg[-400:])
    if info and info.get("entries"):
        return f"Playlist : {info.get('title') or url}"
    return (info or {}).get("title") or url


# =====================================================================
#  INTERFACE
# =====================================================================
_Mixin = TkinterDnD.DnDWrapper if HAS_DND else object


class Choix(ctk.CTkFrame):
    """Rangée de boutons à choix unique : sélection blanche, texte noir."""

    def __init__(self, master, valeurs=None, commande=None, hauteur=34, largeur=66, colonnes=8, taille=13):
        super().__init__(master, fg_color="transparent")
        self.commande, self.valeur, self.boutons = commande, None, {}
        self.hauteur, self.largeur, self.colonnes, self.taille = hauteur, largeur, colonnes, taille
        self.vide = ctk.CTkLabel(self, text="—", text_color="#5b5b78")
        self.definir(valeurs or [])

    def definir(self, valeurs, selection=None):
        for b in self.boutons.values():
            b.destroy()
        self.boutons = {}
        self.vide.grid_forget()
        if not valeurs:
            self.valeur = None
            self.vide.grid(row=0, column=0, sticky="w")
            return
        for i, v in enumerate(valeurs):
            b = ctk.CTkButton(self, text=v, width=self.largeur, height=self.hauteur, corner_radius=10,
                              font=ctk.CTkFont(size=self.taille, weight="bold"),
                              command=lambda v=v: self.choisir(v))
            b.grid(row=i // self.colonnes, column=i % self.colonnes, padx=(0, 6), pady=3, sticky="w")
            self.boutons[v] = b
        if selection in valeurs:
            self.valeur = selection
        elif self.valeur not in valeurs:
            self.valeur = valeurs[0]
        self._colorier()

    def get(self):
        return self.valeur

    def choisir(self, v, notifier=True):
        self.valeur = v
        self._colorier()
        if notifier and self.commande:
            self.commande(v)

    def _colorier(self):
        for v, b in self.boutons.items():
            if v == self.valeur:
                b.configure(fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color=ACCENT_TEXT)
            else:
                b.configure(fg_color=CARD2, hover_color=HOVER, text_color="#e5e5f0")


class Fenetre(ctk.CTk, _Mixin):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        self.title("Morphe")
        self.geometry("820x800")
        self.minsize(740, 700)
        self.configure(fg_color=BG)
        try:
            self.iconbitmap(ressource("icon.ico"))
        except Exception:
            pass

        self.fichiers, self.lignes = [], {}
        self.dossier, self.dernier_dossier, self.occupe = "", "", False
        dl = Path.home() / "Downloads"
        self.dl_dossier = str(dl if dl.exists() else Path.home())
        self.dl_dernier, self.dl_occupe = "", False
        self.dl_annule = threading.Event()

        self._entete()
        self._page_conversion()
        self._page_telechargement()
        self._activer_dnd()
        self._maj_formats()
        self._aller("Convertir")

    # ------------------------------------------------------------ en-tête
    def _entete(self):
        tete = ctk.CTkFrame(self, fg_color="transparent")
        tete.pack(side="top", fill="x", padx=28, pady=(22, 8))
        try:
            logo = ctk.CTkImage(Image.open(ressource("icon.png")), size=(46, 46))
            ctk.CTkLabel(tete, image=logo, text="").pack(side="left")
        except Exception:
            pass
        bloc = ctk.CTkFrame(tete, fg_color="transparent")
        bloc.pack(side="left", padx=12)
        ctk.CTkLabel(bloc, text="Morphe", font=ctk.CTkFont(size=26, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(bloc, text="Convertir · Télécharger", text_color=MUTED).pack(anchor="w")
        self.onglets = Choix(tete, ["Convertir", "Télécharger"], commande=self._aller,
                             hauteur=38, largeur=118, taille=14)
        self.onglets.pack(side="right")

    def _aller(self, nom):
        self.page_conv.pack_forget()
        self.page_dl.pack_forget()
        (self.page_conv if nom == "Convertir" else self.page_dl).pack(side="top", fill="both", expand=True)
        self.onglets.choisir(nom, notifier=False)

    # ------------------------------------------------------------ page Convertir
    def _page_conversion(self):
        p = self.page_conv = ctk.CTkFrame(self, fg_color="transparent")

        pied = ctk.CTkFrame(p, fg_color="transparent")
        pied.pack(side="bottom", fill="x", padx=28, pady=(0, 22))
        self.btn = ctk.CTkButton(pied, text="Convertir", height=54, corner_radius=14,
                                 font=ctk.CTkFont(size=17, weight="bold"), fg_color=ACCENT,
                                 hover_color=ACCENT_HOVER, text_color=ACCENT_TEXT,
                                 text_color_disabled="#6b6b88", command=self.demarrer)
        self.btn.pack(side="bottom", fill="x", pady=(10, 0))
        ligne = ctk.CTkFrame(pied, fg_color="transparent")
        ligne.pack(side="bottom", fill="x")
        self.statut = ctk.CTkLabel(ligne, text="Prêt", text_color=MUTED)
        self.statut.pack(side="left")
        self.btn_dossier = ctk.CTkButton(ligne, text="Ouvrir le dossier", width=130, height=26, corner_radius=8,
                                         fg_color=CARD2, hover_color=HOVER,
                                         command=lambda: ouvrir_dossier(self.dernier_dossier))
        self.barre = ctk.CTkProgressBar(pied, progress_color=ACCENT, fg_color=CARD2, height=10)
        self.barre.set(0)
        self.barre.pack(side="bottom", fill="x", pady=(0, 6))

        opt = ctk.CTkFrame(p, fg_color=CARD, corner_radius=18)
        opt.pack(side="bottom", fill="x", padx=28, pady=10)
        opt.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(opt, text="Convertir en").grid(row=0, column=0, padx=18, pady=(16, 6), sticky="w")
        self.menu = Choix(opt, [], commande=lambda _: self._maj_bouton())
        self.menu.grid(row=0, column=1, padx=18, pady=(16, 6), sticky="ew")
        self.lbl_qualite = ctk.CTkLabel(opt, text="Qualité image : 90")
        self.lbl_qualite.grid(row=1, column=0, padx=18, pady=6, sticky="w")
        self.qualite = ctk.CTkSlider(opt, from_=10, to=100, number_of_steps=90, progress_color=ACCENT,
                                     button_color=ACCENT, button_hover_color=ACCENT_HOVER, fg_color=CARD2,
                                     command=lambda v: self.lbl_qualite.configure(text=f"Qualité image : {int(v)}"))
        self.qualite.set(90)
        self.qualite.grid(row=1, column=1, padx=18, pady=6, sticky="ew")
        ctk.CTkLabel(opt, text="Enregistrer dans").grid(row=2, column=0, padx=18, pady=(6, 16), sticky="w")
        bas = ctk.CTkFrame(opt, fg_color="transparent")
        bas.grid(row=2, column=1, padx=18, pady=(6, 16), sticky="ew")
        self.lbl_dossier = ctk.CTkLabel(bas, text="À côté du fichier d'origine", text_color=MUTED)
        self.lbl_dossier.pack(side="left")
        ctk.CTkButton(bas, text="Choisir", width=80, fg_color=CARD2, hover_color=HOVER,
                      command=self.choisir_dossier).pack(side="right")
        self.btn_reset = ctk.CTkButton(bas, text="Par défaut", width=80, fg_color="transparent",
                                       hover_color=CARD2, text_color=MUTED, command=self.reset_dossier)

        self.zone = ctk.CTkFrame(p, fg_color=CARD, corner_radius=18, border_width=2, border_color=CARD2)
        self.zone.pack(side="top", fill="x", padx=28, pady=10)
        t1 = ctk.CTkLabel(self.zone, text="Glisse tes fichiers ici" if HAS_DND else "Ajoute tes fichiers",
                          font=ctk.CTkFont(size=18, weight="bold"))
        t1.pack(pady=(18, 2))
        t2 = ctk.CTkLabel(self.zone, text="Images, audio, vidéo, documents, tableaux, dossiers… ou clique ici pour parcourir",
                          text_color=MUTED)
        t2.pack(pady=(0, 18))
        for w in (self.zone, t1, t2):
            w.bind("<Button-1>", lambda e: self.parcourir())
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass

        conteneur = ctk.CTkFrame(p, fg_color="transparent")
        conteneur.pack(side="top", fill="both", expand=True, padx=28, pady=(2, 0))
        entete = ctk.CTkFrame(conteneur, fg_color="transparent")
        entete.pack(fill="x", pady=(0, 4))
        self.lbl_nb = ctk.CTkLabel(entete, text="Fichiers", text_color=MUTED)
        self.lbl_nb.pack(side="left", padx=4)
        self.btn_vider = ctk.CTkButton(entete, text="Tout effacer", width=90, height=24, fg_color="transparent",
                                       hover_color=CARD2, text_color=MUTED, command=self.vider)
        self.liste = ctk.CTkScrollableFrame(conteneur, fg_color=CARD, corner_radius=18, height=110)
        self.liste.pack(fill="both", expand=True)
        self.vide = ctk.CTkLabel(self.liste, text="Aucun fichier pour l'instant", text_color="#5b5b78")
        self.vide.pack(pady=30)

    def _activer_dnd(self):
        if not HAS_DND:
            return
        try:
            self.TkdndVersion = TkinterDnD._require(self)
        except Exception:
            return

        def depose(e):
            self.zone.configure(border_color=CARD2)
            self.ajouter(self.tk.splitlist(e.data))
            return e.action

        def entre(e):
            self.zone.configure(border_color=ACCENT)
            return e.action

        def sort(e):
            self.zone.configure(border_color=CARD2)
            return e.action

        for w in (self, self.zone, self.liste):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", depose)
                w.dnd_bind("<<DropEnter>>", entre)
                w.dnd_bind("<<DropLeave>>", sort)
            except Exception:
                pass

    # ----- fichiers -----
    def parcourir(self):
        if not self.occupe:
            self.ajouter(filedialog.askopenfilenames(title="Choisir des fichiers"))

    def ajouter(self, chemins):
        if self.occupe:
            return
        self._aller("Convertir")
        acceptes, ignores = trier_fichiers(chemins)
        for c in acceptes:
            if c not in self.fichiers:
                self.fichiers.append(c)
                self._ligne(c)
        self._maj_formats()
        if ignores:
            exts = sorted({("." + e) if e else "(sans extension)" for e in ignores})
            self.statut.configure(text=f"{len(ignores)} fichier(s) ignoré(s), format non pris en charge : "
                                       + ", ".join(exts[:6]))

    def _ligne(self, chemin):
        t = type_de(chemin)
        ligne = ctk.CTkFrame(self.liste, fg_color=CARD2, corner_radius=12)
        ligne.pack(fill="x", pady=4, padx=4)
        etiquette = {"image": "IMAGE", "audio": "AUDIO", "video": "VIDÉO", "doc": "DOC"}[t]
        ctk.CTkLabel(ligne, text=etiquette, width=60, corner_radius=8, fg_color=COULEUR[t],
                     text_color="#0f0f17", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=10, pady=10)
        nom = Path(chemin).name
        ctk.CTkLabel(ligne, text=nom if len(nom) < 40 else nom[:37] + "…", anchor="w").pack(side="left")
        try:
            taille = taille_lisible(os.path.getsize(chemin))
        except OSError:
            taille = ""
        ctk.CTkLabel(ligne, text=taille, text_color="#6b6b88", font=ctk.CTkFont(size=11)).pack(side="left", padx=8)
        ctk.CTkButton(ligne, text="✕", width=30, height=28, fg_color="transparent", hover_color="#3a2a3f",
                      command=lambda c=chemin: self.retirer(c)).pack(side="right", padx=6)
        etat = ctk.CTkLabel(ligne, text="", text_color=MUTED, width=48)
        etat.pack(side="right", padx=2)
        self.lignes[chemin] = (ligne, etat)

    def retirer(self, chemin):
        if self.occupe or chemin not in self.fichiers:
            return
        self.fichiers.remove(chemin)
        self.lignes.pop(chemin)[0].destroy()
        self._maj_formats()

    def vider(self):
        if self.occupe:
            return
        for ligne, _ in self.lignes.values():
            ligne.destroy()
        self.lignes.clear()
        self.fichiers.clear()
        self.barre.set(0)
        self.btn_dossier.pack_forget()
        self._maj_formats()

    def _maj_formats(self):
        n = len(self.fichiers)
        self.lbl_nb.configure(text=f"{n} fichier{'s' if n > 1 else ''}" if n else "Fichiers")
        if n:
            self.vide.pack_forget()
            self.btn_vider.pack(side="right")
        else:
            self.vide.pack(pady=30)
            self.btn_vider.pack_forget()

        if any(type_de(f) == "image" for f in self.fichiers):
            self.lbl_qualite.grid()
            self.qualite.grid()
        else:
            self.lbl_qualite.grid_remove()
            self.qualite.grid_remove()

        formats = None
        for f in self.fichiers:
            s = formats_pour(f)
            formats = s if formats is None else [x for x in formats if x in s]
        formats = formats or []
        self.menu.definir(formats)
        if n and not formats:
            self.statut.configure(text="Ces fichiers n'ont aucun format de sortie en commun : convertis-les par lots.")
        elif not self.occupe:
            self.statut.configure(text="Prêt")
        self._maj_bouton()

    def _maj_bouton(self):
        fmt, n = self.menu.get(), len(self.fichiers)
        if self.occupe:
            self.btn.configure(state="disabled", fg_color=CARD2, text="Conversion en cours…")
        elif n and fmt:
            self.btn.configure(state="normal", fg_color=ACCENT,
                               text=f"Convertir {n} fichier{'s' if n > 1 else ''} en {fmt}")
        else:
            self.btn.configure(state="disabled", fg_color=CARD2,
                               text="Ajoute des fichiers pour commencer" if not n else "Convertir")

    def choisir_dossier(self):
        d = filedialog.askdirectory()
        if d:
            self.dossier = d
            self.lbl_dossier.configure(text=d if len(d) < 40 else "…" + d[-38:], text_color="#e5e5f0")
            self.btn_reset.pack(side="right", padx=(0, 6))

    def reset_dossier(self):
        self.dossier = ""
        self.lbl_dossier.configure(text="À côté du fichier d'origine", text_color=MUTED)
        self.btn_reset.pack_forget()

    # ----- conversion -----
    def demarrer(self):
        fmt = self.menu.get()
        if self.occupe or not self.fichiers or not fmt:
            return
        self.occupe = True
        self.btn_dossier.pack_forget()
        self.barre.set(0)
        for _, etat in self.lignes.values():
            etat.configure(text="", text_color=MUTED)
        self._maj_bouton()
        threading.Thread(target=self._travail, args=(fmt, int(self.qualite.get())), daemon=True).start()

    def _ui(self, fn):
        self.after(0, fn)

    def _travail(self, fmt, qualite):
        total, erreurs, dernier = len(self.fichiers), [], ""
        for i, src in enumerate(list(self.fichiers)):
            etat = self.lignes[src][1]
            self._ui(lambda e=etat: e.configure(text="…", text_color="#fbbf24"))
            self._ui(lambda s=src: self.statut.configure(text=f"Conversion : {Path(s).name}"))
            try:
                dossier = self.dossier or str(Path(src).parent)
                dernier = dossier

                def prog(p, i=i, e=etat):
                    self._ui(lambda: (self.barre.set((i + p) / total), e.configure(text=f"{int(p * 100)}%")))

                convertir_un(src, fmt, dossier, qualite, prog)
                self._ui(lambda e=etat: e.configure(text="✓", text_color="#34d399"))
            except Exception as ex:
                erreurs.append(f"{Path(src).name} : {ex}")
                self._ui(lambda e=etat: e.configure(text="✗", text_color="#f87171"))
            self._ui(lambda i=i: self.barre.set((i + 1) / total))
        self._ui(lambda: self._fin(erreurs, dernier, total))

    def _fin(self, erreurs, dernier, total):
        self.occupe = False
        self.dernier_dossier = dernier
        self._maj_bouton()
        reussis = total - len(erreurs)
        if dernier and reussis:
            self.btn_dossier.pack(side="right")
        if erreurs:
            self.statut.configure(text=f"{reussis} réussi(s), {len(erreurs)} erreur(s)")
            messagebox.showerror("Morphe", "\n\n".join(erreurs))
        else:
            self.statut.configure(text=f"Terminé ✓  {reussis} fichier{'s' if reussis > 1 else ''} converti{'s' if reussis > 1 else ''}")

    # ------------------------------------------------------------ page Télécharger
    def _page_telechargement(self):
        p = self.page_dl = ctk.CTkFrame(self, fg_color="transparent")

        pied = ctk.CTkFrame(p, fg_color="transparent")
        pied.pack(side="bottom", fill="x", padx=28, pady=(0, 22))
        self.btn_dl = ctk.CTkButton(pied, text="Télécharger", height=54, corner_radius=14,
                                    font=ctk.CTkFont(size=17, weight="bold"), fg_color=ACCENT,
                                    hover_color=ACCENT_HOVER, text_color=ACCENT_TEXT, command=self.dl_action)
        self.btn_dl.pack(side="bottom", fill="x", pady=(10, 0))
        ligne = ctk.CTkFrame(pied, fg_color="transparent")
        ligne.pack(side="bottom", fill="x")
        self.dl_statut = ctk.CTkLabel(ligne, text="Prêt", text_color=MUTED)
        self.dl_statut.pack(side="left")
        self.btn_dl_dossier = ctk.CTkButton(ligne, text="Ouvrir le dossier", width=130, height=26, corner_radius=8,
                                            fg_color=CARD2, hover_color=HOVER,
                                            command=lambda: ouvrir_dossier(self.dl_dernier))
        self.dl_barre = ctk.CTkProgressBar(pied, progress_color=ACCENT, fg_color=CARD2, height=10)
        self.dl_barre.set(0)
        self.dl_barre.pack(side="bottom", fill="x", pady=(0, 6))

        opt = ctk.CTkFrame(p, fg_color=CARD, corner_radius=18)
        opt.pack(side="bottom", fill="x", padx=28, pady=10)
        opt.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(opt, text="Format").grid(row=0, column=0, padx=18, pady=(16, 6), sticky="w")
        self.dl_mode = Choix(opt, ["Vidéo MP4", "Audio MP3"], commande=self._dl_mode_change, largeur=110)
        self.dl_mode.grid(row=0, column=1, padx=18, pady=(16, 6), sticky="w")
        self.dl_lbl_qualite = ctk.CTkLabel(opt, text="Qualité")
        self.dl_lbl_qualite.grid(row=1, column=0, padx=18, pady=6, sticky="w")
        self.dl_qualite = Choix(opt, ["Max", "1080p", "720p", "480p"])
        self.dl_qualite.grid(row=1, column=1, padx=18, pady=6, sticky="w")
        ctk.CTkLabel(opt, text="Enregistrer dans").grid(row=2, column=0, padx=18, pady=(6, 16), sticky="w")
        bas = ctk.CTkFrame(opt, fg_color="transparent")
        bas.grid(row=2, column=1, padx=18, pady=(6, 16), sticky="ew")
        self.dl_lbl_dossier = ctk.CTkLabel(bas, text=self._court(self.dl_dossier), text_color="#e5e5f0")
        self.dl_lbl_dossier.pack(side="left")
        ctk.CTkButton(bas, text="Choisir", width=80, fg_color=CARD2, hover_color=HOVER,
                      command=self._dl_choisir_dossier).pack(side="right")

        liens = ctk.CTkFrame(p, fg_color=CARD, corner_radius=18)
        liens.pack(side="top", fill="x", padx=28, pady=10)
        haut = ctk.CTkFrame(liens, fg_color="transparent")
        haut.pack(fill="x", padx=18, pady=(14, 6))
        ctk.CTkLabel(haut, text="Colle un ou plusieurs liens (un par ligne)",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        ctk.CTkButton(haut, text="Effacer", width=70, height=26, fg_color="transparent", hover_color=CARD2,
                      text_color=MUTED, command=lambda: self.dl_texte.delete("1.0", "end")).pack(side="right")
        ctk.CTkButton(haut, text="Coller", width=70, height=26, fg_color=CARD2, hover_color=HOVER,
                      command=self._dl_coller).pack(side="right", padx=6)
        self.dl_texte = ctk.CTkTextbox(liens, height=92, fg_color=CARD2, corner_radius=12, wrap="none")
        self.dl_texte.pack(fill="x", padx=18, pady=(0, 6))
        ctk.CTkLabel(liens, text="YouTube, TikTok, Instagram, X, Vimeo, Twitch… (à utiliser pour des contenus "
                                 "dont tu as le droit)", text_color="#5b5b78", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=18, pady=(0, 12))

        journal = ctk.CTkFrame(p, fg_color="transparent")
        journal.pack(side="top", fill="both", expand=True, padx=28, pady=(2, 0))
        ctk.CTkLabel(journal, text="Journal", text_color=MUTED).pack(anchor="w", padx=4, pady=(0, 4))
        self.dl_log = ctk.CTkTextbox(journal, fg_color=CARD, corner_radius=18, height=80, state="disabled",
                                     wrap="word")
        self.dl_log.pack(fill="both", expand=True)

    @staticmethod
    def _court(d):
        return d if len(d) < 44 else "…" + d[-42:]

    def _dl_mode_change(self, v):
        if v.startswith("Audio"):
            self.dl_lbl_qualite.grid_remove()
            self.dl_qualite.grid_remove()
        else:
            self.dl_lbl_qualite.grid()
            self.dl_qualite.grid()

    def _dl_coller(self):
        try:
            self.dl_texte.insert("end", self.clipboard_get().strip() + "\n")
        except Exception:
            pass

    def _dl_choisir_dossier(self):
        d = filedialog.askdirectory(initialdir=self.dl_dossier)
        if d:
            self.dl_dossier = d
            self.dl_lbl_dossier.configure(text=self._court(d))

    def _log(self, texte):
        self.dl_log.configure(state="normal")
        self.dl_log.insert("end", texte + "\n")
        self.dl_log.see("end")
        self.dl_log.configure(state="disabled")

    def dl_action(self):
        if self.dl_occupe:
            self.dl_annule.set()
            self.btn_dl.configure(state="disabled", text="Annulation…")
            return
        urls = [l.strip() for l in self.dl_texte.get("1.0", "end").splitlines() if l.strip()]
        urls = [u for u in urls if re.match(r"https?://", u, re.I)]
        if not urls:
            self.dl_statut.configure(text="Colle au moins un lien valide (commençant par http).")
            return
        self.dl_occupe = True
        self.dl_annule.clear()
        self.btn_dl_dossier.pack_forget()
        self.dl_barre.set(0)
        self.btn_dl.configure(text="Annuler", fg_color=CARD2, hover_color=HOVER, text_color="#e5e5f0")
        mode = "audio" if self.dl_mode.get().startswith("Audio") else "video"
        q = self.dl_qualite.get()
        hauteur = None if q == "Max" else int(q[:-1])
        threading.Thread(target=self._dl_travail, args=(urls, mode, hauteur, self.dl_dossier), daemon=True).start()

    def _dl_travail(self, urls, mode, hauteur, dossier):
        total, ok, erreurs, annule = len(urls), 0, 0, False
        for i, url in enumerate(urls):
            self._ui(lambda i=i: self.dl_statut.configure(text=f"Lien {i + 1}/{total} : analyse…"))

            def rappel(p, texte, i=i):
                self._ui(lambda: (self.dl_barre.set((i + (p or 0)) / total),
                                  self.dl_statut.configure(text=f"Lien {i + 1}/{total} : {texte}")))

            try:
                titre = telecharger(url, mode, hauteur, dossier, rappel, self.dl_annule)
                ok += 1
                self._ui(lambda t=titre: self._log(f"✓ {t}"))
            except InterruptedError:
                annule = True
                self._ui(lambda: self._log("■ Téléchargement annulé"))
                break
            except Exception as ex:
                erreurs += 1
                self._ui(lambda u=url, m=str(ex): self._log(f"✗ {u}\n   {m}"))
            self._ui(lambda i=i: self.dl_barre.set((i + 1) / total))
        self._ui(lambda: self._dl_fin(ok, erreurs, annule, dossier))

    def _dl_fin(self, ok, erreurs, annule, dossier):
        self.dl_occupe = False
        self.dl_dernier = dossier
        self.btn_dl.configure(state="normal", text="Télécharger", fg_color=ACCENT, hover_color=ACCENT_HOVER,
                              text_color=ACCENT_TEXT)
        if annule:
            self.dl_statut.configure(text="Annulé")
        elif erreurs:
            self.dl_statut.configure(text=f"{ok} réussi(s), {erreurs} erreur(s) : détails dans le journal")
            self._log("Astuce : YouTube change souvent, relance lancer.bat pour mettre yt-dlp à jour.")
        else:
            self.dl_statut.configure(text=f"Terminé ✓  {ok} téléchargement{'s' if ok > 1 else ''}")
        if ok:
            self.btn_dl_dossier.pack(side="right")


if __name__ == "__main__":
    Fenetre().mainloop()
