rd /s /q __pycache__
rd /s /q build
rd /s /q dist
pyinstaller GazoNarabe.py --noconsole --icon GN.ico
