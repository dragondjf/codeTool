rm -rf dist
cxfreeze main.py -c --target-dir dist --target-name=CodeTool.exe --icon=icons/app.ico  --base-name=win32gui
cp -r data  dist
cp -r icons dist
