@echo off
cd /d "%~dp0"
echo [%date% %time%] Iniciando coletor de sinistros...
python coletor_sinistros.py >> log_execucao.txt 2>&1
echo [%date% %time%] Finalizado. Ver log_execucao.txt para detalhes.
pause
