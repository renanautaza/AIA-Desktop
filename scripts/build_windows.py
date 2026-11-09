"""Build the Windows executable from the repository root."""

from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src" / "Autaza_Inspection_Analytics.py"
ASSETS = ROOT / "assets"

command = [
    sys.executable,
    "-m",
    "PyInstaller",
    "--noconfirm",
    "--windowed",
    "--name",
    "Autaza Inspection Analytics",
    "--icon",
    str(ASSETS / "AIA_Inspection_Analytics.ico"),
    "--add-data",
    f"{ASSETS}{';' if sys.platform.startswith('win') else ':'}assets",
    str(SOURCE),
]

subprocess.run(command, cwd=ROOT, check=True)
