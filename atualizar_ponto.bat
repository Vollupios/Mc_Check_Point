@echo off
REM Atualiza o ponto.html com a planilha do Cartao Ponto.
REM Uso: arraste a planilha .xls para cima deste arquivo
REM      ou rode:  atualizar_ponto.bat ponto_ate_m_s_10.xls
chcp 65001 >nul
cd /d "%~dp0"

if "%~1"=="" (
    echo Uso: atualizar_ponto.bat ^<planilha.xls^>
    pause
    exit /b 1
)

set PY=python
where python >nul 2>nul || set PY=py
%PY% extrator.py "%~1" --html ponto.html
pause
