del -r dist
del -r build
pyinstaller --path=d:\temp\eclipse\LibA5\A5 CdC.py -w -F
copy CdC_Cfg.xlsx dist\.
cd dist\CdC
pause
