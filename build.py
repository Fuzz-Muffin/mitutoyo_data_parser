import subprocess
import shutil
from pathlib import Path
import sys

def main():
    subprocess.run([sys.executable, "-m", "PyInstaller", "--onefile", "-w", "mitutoyo_parser.py"], check=True)

    # optional cleanup
    for path in ["build", "__pycache__"]:
        shutil.rmtree(path, ignore_errors=True)

if __name__ == "__main__":
    main()