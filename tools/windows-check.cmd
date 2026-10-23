@echo off
rem Vision Office: a conferencia no Windows (docs/instalacao.md, secao 3.5).
rem Duplo clique neste arquivo, com o windows-check.ps1 na mesma pasta.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0windows-check.ps1" %*
if errorlevel 1 pause
