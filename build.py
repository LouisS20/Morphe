"""
Installe les dépendances une par une (une qui échoue n'empêche pas les autres)
puis, si on le demande, compile Morphe.exe.

  python build.py install   -> installe / met à jour les modules
  python build.py build     -> installe puis compile dist\\Morphe.exe
"""
import importlib.util
import os
import subprocess
import sys

ICI = os.path.dirname(os.path.abspath(__file__))

# (paquet pip, module à embarquer par PyInstaller, indispensable ?)
PAQUETS = [
    ("customtkinter", "customtkinter", True),
    ("pillow", "PIL", True),
    ("imageio-ffmpeg", "imageio_ffmpeg", True),
    ("tkinterdnd2", "tkinterdnd2", False),
    ("python-docx", "docx", False),
    ("markdown", "markdown", False),
    ("markdownify", "markdownify", False),
    ("beautifulsoup4", "bs4", False),
    ("openpyxl", "openpyxl", False),
    ("pypdf", "pypdf", False),
    ("reportlab", "reportlab", False),
    ("pyyaml", "yaml", False),
    ("tomli-w", "tomli_w", False),
    ("pillow-heif", "pillow_heif", False),
    ("pymupdf", "pymupdf", False),          # optionnel : PDF -> PNG
    ("yt-dlp[default]", "yt_dlp", False),   # téléchargement (+ yt-dlp-ejs pour YouTube)
    ("deno", "deno", False),                # runtime JavaScript embarqué pour YouTube
]


def pip(paquet, maj=False):
    cmd = [sys.executable, "-m", "pip", "install", "-q"] + (["--upgrade"] if maj else []) + [paquet]
    return subprocess.run(cmd).returncode == 0


def installer():
    for paquet, module, _ in PAQUETS:
        deja = importlib.util.find_spec(module) is not None
        # yt-dlp doit toujours être à jour : YouTube casse régulièrement les vieilles versions
        if deja and not paquet.startswith("yt-dlp"):
            continue
        print(f"> {paquet}")
        if not pip(paquet, maj=paquet.startswith("yt-dlp")):
            print(f"  !! échec pour {paquet} : la fonction correspondante sera désactivée")
    if importlib.util.find_spec("PyInstaller") is None:
        pip("pyinstaller")


def compiler():
    os.chdir(ICI)
    subprocess.run([sys.executable, "make_icon.py"])
    cmd = [sys.executable, "-m", "PyInstaller", "--noconsole", "--onefile", "--name", "Morphe",
           "--icon", "icon.ico",
           "--add-data", f"icon.ico{os.pathsep}.", "--add-data", f"icon.png{os.pathsep}."]
    for _, module, _ in PAQUETS:
        if importlib.util.find_spec(module) is not None:
            cmd += ["--collect-all", module]
    cmd.append("morphe.py")
    return subprocess.run(cmd).returncode


if __name__ == "__main__":
    action = sys.argv[1] if len(sys.argv) > 1 else "install"
    installer()
    if action == "build":
        code = compiler()
        print("\nTerminé : dist\\Morphe.exe" if code == 0 else "\nERREUR : la compilation a échoué (voir plus haut).")
        sys.exit(code)
