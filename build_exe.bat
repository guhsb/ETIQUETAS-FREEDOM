@echo off
chcp 65001 >nul
echo === Gerando EtiquetasFreedom (pasta com o .exe) ===
python --version >nul 2>&1 || (echo Python nao encontrado. Instale em python.org e marque "Add to PATH". & pause & exit /b 1)
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --onedir --noconsole --icon assets\icone.ico --name EtiquetasFreedom ^
  --add-data "fonts;fonts" --add-data "assets;assets" --add-data "logo.png;." --add-data "cadastros_padrao.json;." ^
  --add-data "clientes.csv;." --hidden-import win32timezone --collect-all rapidocr_onnxruntime app.py
if not exist dist\EtiquetasFreedom\EtiquetasFreedom.exe (echo Deu erro ao gerar o programa & pause & exit /b 1)
if exist EtiquetasFreedom_pronto rmdir /s /q EtiquetasFreedom_pronto
xcopy /E /I /Y dist\EtiquetasFreedom EtiquetasFreedom_pronto >nul
copy /Y clientes.csv EtiquetasFreedom_pronto\ >nul
echo.
echo Pronto! Use a pasta EtiquetasFreedom_pronto inteira (pode mover para onde quiser, ex.: C:\Etiquetas).
echo O programa cria sozinho a pasta "dados" e o HISTORICO_ETIQUETAS.xlsx ao lado do .exe.
pause
