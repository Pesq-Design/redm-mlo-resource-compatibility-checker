@echo off
rem Doble clic: abre la ventana para elegir el asset, la carpeta resources y los .cfg.
rem No modifica nada del servidor: solo lee y deja un reporte .txt en la carpeta "reportes".
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0compat_rapido.ps1"
