rd /s /q __pycache__
rd /s /q build
rd /s /q dist
rem pyinstaller GazoNarabe.py --icon GN.ico
pyinstaller GazoNarabe.py --noconsole --icon GN.ico
