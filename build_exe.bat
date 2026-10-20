@echo off
chcp 65001 >nul
echo === Gerando EtiquetasFreedom.exe ===
python --version >nul 2>&1 || (echo Python nao encontrado. Instale em python.org e marque "Add to PATH". & pause & exit /b 1)
python -m pip install -r requirements.txt
python -m PyInstaller --onefile --noconsole --name EtiquetasFreedom ^
  --add-data "fonts;fonts" --add-data "assets;assets" --add-data "logo.png;." --add-data "cadastros_padrao.json;." --add-data "clientes.csv;." ^
  --hidden-import win32timezone app.py
if not exist dist\EtiquetasFreedom.exe (echo Deu erro ao gerar o .exe & pause & exit /b 1)
rem monta a pasta pronta: o .exe e o clientes.csv lado a lado
if not exist EtiquetasFreedom_pronto mkdir EtiquetasFreedom_pronto
copy /Y dist\EtiquetasFreedom.exe EtiquetasFreedom_pronto\ >nul
copy /Y clientes.csv EtiquetasFreedom_pronto\ >nul
echo.
echo Pronto! Use a pasta EtiquetasFreedom_pronto (pode mover para onde quiser, ex.: C:\Etiquetas).
echo O programa cria sozinho a pasta "dados" e o HISTORICO_ETIQUETAS.xlsx ao lado do .exe.
pause
