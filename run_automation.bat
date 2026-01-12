@echo off

REM --- COMO AGENDAR ESTE SCRIPT ---
REM
REM Este arquivo orquestra toda a automacao. Voce deve criar DUAS tarefas no Agendador do Windows:
REM
REM TAREFA 1: ATUALIZACAO DE DADOS (A CADA 15 MINUTOS)
REM   - Objetivo: Apenas baixar as planilhas.
REM   - Acao: Agende a execucao deste arquivo `run_automation.bat`.
REM   - Disparador: Iniciar em uma hora (ex: 07:00), e repetir a cada 15 minutos, indefinidamente.
REM   - Como fazer? Na guia "Disparadores", em "Configuracoes avancadas", marque "Repetir tarefa a cada:" e selecione "15 minutos".
REM
REM TAREFA 2: ENVIO DE RELATORIOS (3 VEZES AO DIA)
REM   - Objetivo: Baixar os dados mais recentes E enviar os emails.
REM   - Acao: Agende a execucao deste arquivo `run_automation.bat`.
REM   - Disparadores: Crie tres disparadores diarios, um para as 08:00, um para as 12:00 e um para as 16:30.
REM
REM O script `ofensor.py` tem a logica para saber qual relatorio enviar em qual horario.

echo [================================================================]
echo [=            INICIANDO ORQUESTRADOR DE AUTOMACAO DE RELATORIOS            =]
echo [================================================================]
echo.

REM --- ETAPA 1: BAIXAR PLANILHAS DO SHAREPOINT ---
echo [%TIME%] Iniciando download das planilhas mais recentes...
python download_punches.py
echo [%TIME%] Download das planilhas finalizado.
echo.

REM --- ETAPA 2: GERAR E ENVIAR RELATORIOS ---
echo [%TIME%] Iniciando geracao e envio dos relatorios...
python ofensor.py
echo [%TIME%] Processamento de relatorios finalizado.
echo.

echo [================================================================]
echo [=                 EXECUCAO DO ORQUESTRADOR FINALIZADA                  =]
echo [================================================================]

pause
