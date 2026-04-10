@echo off
echo ============================================================
echo  GENERADOR DE INDICE + FOLIADOR -- CRUBC Los Rios v7.1
echo  Instalador de Python y librerias
echo ============================================================
echo.

:: ── PASO 1: Verificar si Python esta instalado ───────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado en este equipo.
    echo.
    echo Para instalarlo:
    echo   1. Abre tu navegador y ve a: https://www.python.org/downloads/
    echo   2. Descarga la version mas reciente (Python 3.12 o superior^)
    echo   3. Ejecuta el instalador
    echo   4. MUY IMPORTANTE: marca la casilla "Add Python to PATH"
    echo      antes de hacer clic en Install Now
    echo   5. Una vez instalado, vuelve a ejecutar este archivo .bat
    echo.
    echo Si no tienes acceso a internet o permisos de instalacion,
    echo contacta a tu soporte informatico.
    echo.
    pause
    exit /b 1
)

:: ── PASO 2: Mostrar version instalada ────────────────────────
echo [OK] Python detectado:
python --version
echo.

:: ── PASO 3: Instalar librerias obligatorias ──────────────────
echo Instalando librerias obligatorias...
echo (pdfplumber, python-docx, openpyxl, pillow^)
echo.
pip install pdfplumber python-docx openpyxl pillow
if %errorlevel% neq 0 (
    echo.
    echo [ADVERTENCIA] Hubo un problema instalando algunas librerias.
    echo Verifica tu conexion a internet e intenta nuevamente.
    echo.
) else (
    echo.
    echo [OK] Librerias obligatorias instaladas correctamente.
    echo.
)

:: ── PASO 4: Instalar librerias opcionales ────────────────────
echo Instalando librerias opcionales...
echo (python-pptx para PowerPoint, rarfile para RAR, py7zr para 7z^)
echo.
pip install python-pptx rarfile py7zr
echo.

:: ── PASO 5: Verificar instalacion ────────────────────────────
echo Verificando instalacion...
python -c "import pdfplumber; print('[OK] pdfplumber')" 2>nul || echo "[FALTA] pdfplumber"
python -c "import docx;       print('[OK] python-docx')" 2>nul || echo "[FALTA] python-docx"
python -c "import openpyxl;   print('[OK] openpyxl')"   2>nul || echo "[FALTA] openpyxl"
python -c "import PIL;        print('[OK] pillow')"      2>nul || echo "[FALTA] pillow"
python -c "import pptx;       print('[OK] python-pptx')" 2>nul || echo "[FALTA - opcional] python-pptx"
python -c "import rarfile;    print('[OK] rarfile')"     2>nul || echo "[FALTA - opcional] rarfile"
python -c "import py7zr;      print('[OK] py7zr')"       2>nul || echo "[FALTA - opcional] py7zr"

echo.
echo ============================================================
echo  Instalacion completada.
echo  Ahora puedes abrir la aplicacion haciendo doble clic en:
echo     GeneradorIndice_CRUBC.vbs
echo ============================================================
echo.
pause
