@echo off
REM ===========================================================================
REM AUDIT.bat - Ejecutable rapido para Windows
REM Ejecuta la auditoria de accesos post-baja con Python.
REM ===========================================================================
python AUDIT.py %*
if errorlevel 1 (
    echo.
    echo [ERROR] Hubo un error al ejecutar la auditoria.
)
