@echo off
setlocal
title Export Portal Users - ANCPI

rem Always run from the folder where this .bat lives.
cd /d "%~dp0"

echo ============================================================
echo   Export Portal Users - ANCPI
echo ============================================================
echo.

rem --- Locate a Python interpreter (prefer the py launcher) --------------
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY (
    where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo [EROARE] Python nu a fost gasit in PATH.
    echo Instaleaza Python de la https://www.python.org/downloads/ si reincearca.
    echo.
    pause
    exit /b 1
)

echo [1/3] Python gasit: %PY%
%PY% --version
echo.

rem --- Ensure the openpyxl dependency (needed for the .xlsx export) ------
echo [2/3] Verific dependinta openpyxl...
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo      openpyxl lipseste - il instalez...
    %PY% -m pip install --quiet --disable-pip-version-check openpyxl
    if errorlevel 1 (
        echo [EROARE] Instalarea openpyxl a esuat.
        echo.
        pause
        exit /b 1
    )
    echo      openpyxl instalat.
) else (
    echo      openpyxl este deja instalat.
)
echo.

rem --- Launch the application -------------------------------------------
echo [3/3] Pornesc aplicatia...
echo ------------------------------------------------------------
%PY% "%~dp0export_portal_users.py"
set "RC=%ERRORLEVEL%"
echo ------------------------------------------------------------

if "%RC%"=="0" (
    echo.
    echo Gata. Vezi mai sus calea exacta unde au fost salvate fisierele.
) else (
    echo.
    echo Aplicatia s-a incheiat cu eroare ^(cod %RC%^).
)

echo.
pause
endlocal
exit /b %RC%
