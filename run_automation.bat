@echo off
echo [================================================================]
echo [=            INICIANDO ORQUESTRADOR DE AUTOMACAO DE RELATORIOS            =]
echo [================================================================]
echo.

REM --- ETAPA 1: BAIXAR PLANILHAS DO SHAREPOINT ---
echo [%TIME%] Iniciando download das planilhas mais recentes...
python download_punches.py
echo [%TIME%] Download das planilhas finalizado.
echo.

REM --- ETAPA 2: PROCESSAR DADOS E ENVIAR RELATORIOS (SE HOUVER AGENDAMENTO) ---
echo [%TIME%] Iniciando processamento de relatorios...
python ofensor.py
echo [%TIME%] Processamento de relatorios finalizado.
echo.

echo [================================================================]
echo [=                 EXECUCAO DO ORQUESTRADOR FINALIZADA                  =]
echo [================================================================]

pause
