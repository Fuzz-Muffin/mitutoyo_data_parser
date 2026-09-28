import subprocess
import shutil
from pathlib import Path

def main():
    subprocess.run(["pyinstaller", "--onefile", "mitutoyo_parser.py"], check=True)

    # optional cleanup
    for path in ["build", "__pycache__"]:
        shutil.rmtree(path, ignore_errors=True)

if __name__ == "__main__":
    main()