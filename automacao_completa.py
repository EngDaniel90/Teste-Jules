# --- INSTRUCOES DE EXECUCAO ---
# Este é um script autônomo. Para iniciar a automação completa:
# 1. Certifique-se de que todas as bibliotecas em `requirements.txt` estão instaladas (`pip install -r requirements.txt`).
# 2. Configure as variáveis de ambiente `SHAREPOINT_USER` e `SHAREPOINT_PASSWORD` com suas credenciais.
# 3. Execute este script a partir do terminal: `python automacao_completa.py`
#
# Para mantê-lo rodando em segundo plano de forma robusta no Windows, você pode usar o seguinte comando:
# `start "AutomacaoRelatorios" /B python automacao_completa.py`
#
# O script ficará rodando indefinidamente, executando as tarefas nos horários agendados.

import pandas as pd
import traceback
from datetime import datetime
import win32com.client as win32
import matplotlib.pyplot as plt
import seaborn as sns
import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import schedule
import time

# --- CONFIGURACOES GERAIS ---
PATH_PUNCH_TS = r'C:\Users\E797\Downloads\Teste mensagem e print\Punch_DR90_TS.xlsx'
PATH_RDS = r'C:\Users\E797\Downloads\Teste mensagem e print\RDs\RDs.xlsx'
PATH_DASHBOARD_TS = r'C:\Users\E797\Downloads\Teste mensagem e print\dashboard_status.png'
PATH_OP_CHECK = r'C:\Users\E797\Downloads\Teste mensagem e print\Operation to check.xlsx'
PATH_ESUP_CHECK = r'C:\Users\E797\Downloads\Teste mensagem e print\ESUP to check.xlsx'
PATH_JULIUS_CHECK = r'C:\Users\E797\Downloads\Teste mensagem e print\Julius to check.xlsx'
PATH_EHOUSE_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_E-House.xlsx"
PATH_EHOUSE_GRAPH = r"C:\Users\E797\Downloads\Teste mensagem e print\ehouse_status_graph.png"
PATH_VENDORS_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_Vendors.xlsx"
PATH_VENDORS_GRAPH = r"C:\Users\E797\Downloads\Teste mensagem e print\vendors_status_graph.png"
EMAIL_DESTINO = "658b4ef7.petrobras.com.br@br.teams.ms"
EMAIL_JULIUS = "julius.lorzales.prestserv@petrobras.com.br"

SHAREPOINT_USER = os.getenv("SHAREPOINT_USER")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
SHAREPOINT_SITE_URL = "https://seatrium.sharepoint.com/sites/P84P85DesignReview"
PUNCH_LISTS_CONFIG = {
    "Topside": {"list_title": "DR90 Topside Punchlist", "file_path": PATH_PUNCH_TS, "columns_to_keep": None},
    "E-House": {
        "list_title": "DR90 E-House Punchlist", "file_path": PATH_EHOUSE_PUNCH,
        "columns_to_keep": ['Punch No', 'Zone', 'DECK No.', 'Zone-Punch Number', 'Action Description', 'Punched by', 'Punch SnapShot1', 'Punch SnapShot2', 'Closing SnapShot1', 'Hotwork', 'ABB/CIMC Discipline', 'Company', 'Close Out Plan Date', 'Action by', 'Status', 'Action Comment', 'Date Cleared by ABB', 'Days Since Date Cleared by ABB', 'KBR Response', 'KBR Response Date', 'KBR Response by', 'KBR Remarks', 'KBR Category', 'KBR Discipline', 'KBR Screenshot', 'Date Cleared by KBR', 'Days Since Date Cleared By KBR', 'Seatrium Discipline', 'Seatrium Remarks', 'Checked By (Seatrium)', 'Seatrium Comments', 'Date Cleared By Seatrium', 'Days Since Date Cleared by Seatrium', 'Petrobras Response', 'Petrobras Response By', 'Petrobras Screenshot', 'Petrobras Response Date', 'Petrobras Remarks', 'Petrobras Discipline', 'Petrobras Category', 'Date Cleared by Petrobras', 'Days Since Date Cleared By Petrobras', 'Additional Remarks', 'ARC Reference No(HFE Only)', 'Modified', 'Modified By', 'Item Type', 'Path']
    },
    "Vendors": {
        "list_title": "Vendor Package Review Punchlist DR90", "file_path": PATH_VENDORS_PUNCH,
        "columns_to_keep": ['Punch No', 'Zone', 'DECK No.', 'Zone-Punch Number', 'Action Description', 'S3D Item Tags', 'Punched by', 'Punch Snapshot', 'Punch Snapshot 2', 'Punch Snapshot 3', 'Punch Snapshot 4', 'Close-Out Snapshot 1', 'Close-Out Snapshot 2', 'Action Comment', 'Vendor Discipline', 'Company', 'Action by', 'Status', 'Date Cleared by KBR', 'Days Since Date Cleared by KBR', 'Petrobras Response', 'Petrobras Response by', 'Petrobras Response Date', 'Petrobras Screenshot', 'Remarks', 'Petrobras Discipline', 'Petrobras Category', 'Date Cleared by Petrobras', 'Seatrium Remarks', 'Seatrium Discipline', 'Checked By (Seatrium)', 'Seatrium Comments', 'Date Cleared By Seatrium', 'Days Since Date Cleared by Seatrium', 'Modified By', 'Item Type', 'Path']
    }
}

# --- FUNCOES DE DOWNLOAD ---
def format_as_table(writer, df, sheet_name):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    (max_row, max_col) = df.shape
    column_settings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

def download_punch_lists():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] INICIANDO DOWNLOAD DAS PLANILHAS...")
    if not SHAREPOINT_USER or not SHAREPOINT_PASSWORD:
        print("-> ERRO: Variáveis de ambiente SHAREPOINT_USER e SHAREPOINT_PASSWORD não configuradas.")
        return
    try:
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(UserCredential(SHAREPOINT_USER, SHAREPOINT_PASSWORD))
        for name, config in PUNCH_LISTS_CONFIG.items():
            print(f"--> Baixando lista: {name}")
            list_obj = ctx.web.lists.get_by_title(config["list_title"])
            items = list_obj.get_items().execute_query()
            if not items:
                print(f"--> AVISO: Lista '{name}' está vazia ou não foi acessada.")
                continue

            data = [item.properties for item in items]
            df = pd.DataFrame(data)
            df.columns = [col.replace('_x0020_', ' ') for col in df.columns]

            if config["columns_to_keep"]:
                df = df[[col for col in config["columns_to_keep"] if col in df.columns]]

            with pd.ExcelWriter(config["file_path"], engine='xlsxwriter') as writer:
                format_as_table(writer, df, name)
            print(f"--> SUCESSO: Planilha '{os.path.basename(config['file_path'])}' salva.")
    except Exception as e:
        print(f"-> !!! ERRO CRÍTICO DURANTE DOWNLOAD: {e} !!!")

# --- FUNCOES DE PROCESSAMENTO DE DADOS ---
def processar_dados_topside():
    # ... (código da função processar_dados do ofensor.py)
    pass

def processar_dados_ehouse():
    # ... (código da função processar_dados_ehouse do ofensor.py)
    pass

def processar_dados_vendors():
    # ... (código da função processar_dados_vendors do ofensor.py)
    pass

# --- FUNCOES DE GERACAO DE GRAFICOS ---
def gerar_dashboard_topside(dados):
    # ... (código da função gerar_dashboard_imagem do ofensor.py)
    pass

def gerar_grafico_ehouse(dados):
    # ... (código da função gerar_grafico_ehouse do ofensor.py)
    pass

def gerar_dashboard_vendors(dados):
    # ... (código da função gerar_dashboard_vendors do ofensor.py)
    pass

# --- FUNCOES DE ENVIO DE EMAIL ---
def enviar_email_topside(dados, log_processo):
    # ... (código da função enviar_email do ofensor.py)
    pass

def enviar_email_de_falha(log_processo):
    # ... (código da função enviar_email_de_falha do ofensor.py)
    pass

def enviar_mensagem_julius(dados):
    # ... (código da função enviar_mensagem_julius do ofensor.py)
    pass

def enviar_email_ehouse(dados):
    # ... (código da função enviar_email_ehouse do ofensor.py)
    pass

def enviar_email_vendors(dados):
    # ... (código da função enviar_email_vendors do ofensor.py)
    pass

# --- FUNCAO ORQUESTRADORA DE RELATORIOS ---
def gerar_e_enviar_relatorios():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] INICIANDO GERAÇÃO E ENVIO DE RELATÓRIOS...")

    # --- FLUXO 1: Relatório Principal (Topside) ---
    print("\n--- [FLUXO 1/3] Processando Relatório Principal (Topside) ---")
    try:
        dados_topside, log_topside, sucesso_topside = processar_dados_topside()
        if sucesso_topside:
            print("-> Dados Topside processados com sucesso.")
            sucesso_dashboard, log_dashboard = gerar_dashboard_topside(dados_topside)
            log_total_topside = log_topside + log_dashboard
            if sucesso_dashboard: print("-> Dashboard Topside gerado com sucesso.")
            else: print("-> !!! FALHA NA GERAÇÃO DO DASHBOARD TOPSIDE !!!")
            enviar_email_topside(dados_topside, log_total_topside)

            # E-mail para Julius só é enviado se o fluxo principal for bem-sucedido
            print("\n--- Verificando E-mail para Julius ---")
            if 7 <= datetime.now().hour < 9:
                enviar_mensagem_julius(dados_topside)
            else:
                print(f"-> Fora do horário (executado às {datetime.now().hour}h). E-mail para Julius não enviado.")
        else:
            print("\n!!! FALHA CRÍTICA NO PROCESSAMENTO DOS DADOS TOPSIDE !!!")
            enviar_email_de_falha(log_topside)
    except Exception as e:
        enviar_email_de_falha([f"Erro inesperado no fluxo Topside: {e}"])


    # --- FLUXO 2: Relatório E-House ---
    print("\n--- [FLUXO 2/3] Processando Relatório E-House ---")
    try:
        dados_ehouse, log_ehouse, sucesso_ehouse = processar_dados_ehouse()
        if sucesso_ehouse:
            print("-> Dados E-House processados com sucesso.")
            sucesso_grafico, log_grafico = gerar_grafico_ehouse(dados_ehouse)
            if sucesso_grafico:
                print("-> Gráfico E-House gerado com sucesso.")
                enviar_email_ehouse(dados_ehouse)
            else:
                print("-> !!! FALHA NA GERAÇÃO DO GRÁFICO E-HOUSE !!!")
                enviar_email_de_falha(log_ehouse + log_grafico)
    except FileNotFoundError:
        print(f"-> Arquivo E-House não encontrado. O relatório para este fluxo não será gerado.")
    except Exception as e:
        enviar_email_de_falha([f"Erro inesperado no fluxo E-House: {e}"])

    # --- FLUXO 3: Relatório Vendors ---
    print("\n--- [FLUXO 3/3] Processando Relatório Vendors ---")
    try:
        dados_vendors, log_vendors, sucesso_vendors = processar_dados_vendors()
        if sucesso_vendors:
            print("-> Dados de Vendors processados com sucesso.")
            sucesso_grafico, log_grafico = gerar_dashboard_vendors(dados_vendors)
            if sucesso_grafico:
                print("-> Dashboard de Vendors gerado com sucesso.")
                enviar_email_vendors(dados_vendors)
            else:
                print("-> !!! FALHA NA GERAÇÃO DO DASHBOARD DE VENDORS !!!")
                enviar_email_de_falha(log_vendors + log_grafico)
    except FileNotFoundError:
        print(f"-> Arquivo de Vendors não encontrado. O relatório para este fluxo não será gerado.")
    except Exception as e:
        enviar_email_de_falha([f"Erro inesperado no fluxo Vendors: {e}"])

# --- AGENDAMENTO ---
if __name__ == "__main__":
    print(">>> INICIANDO O AGENDADOR DE AUTOMACAO <<<")

    schedule.every(15).minutes.do(download_punch_lists)
    schedule.every().day.at("08:00").do(gerar_e_enviar_relatorios)
    schedule.every().day.at("12:00").do(gerar_e_enviar_relatorios)
    schedule.every().day.at("16:30").do(gerar_e_enviar_relatorios)

    print("-> Executando tarefas pela primeira vez para validação inicial...")
    download_punch_lists()
    gerar_e_enviar_relatorios()

    print("-> Agendamento configurado. O script agora está monitorando os horários...")
    while True:
        schedule.run_pending()
        time.sleep(60) # Verifica a cada 60 segundos
