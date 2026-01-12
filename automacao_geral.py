# --- INSTRUÇÕES DE EXECUÇÃO ---
# Este é um script autônomo. Para iniciar a automação completa:
# 1. Certifique-se de que todas as bibliotecas em `requirements.txt` estão instaladas.
# 2. Configure as variáveis de ambiente para suas credenciais, se aplicável.
# 3. Execute este script a partir do terminal: `python automacao_geral.py`
# O script abrirá um navegador para autenticação inicial. Após o login, ele ficará rodando indefinidamente,
# executando as tarefas de download e relatório nos horários agendados.

import pandas as pd
import traceback
from datetime import datetime
import win32com.client as win32
import matplotlib.pyplot as plt
import seaborn as sns
import os
import requests
import schedule
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib3

# --- CONFIGURAÇÕES DE REDE ---
os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURAÇÕES GERAIS ---
EMAIL_DESTINO = "658b4ef7.petrobras.com.br@br.teams.ms"
EMAIL_JULIUS = "julius.lorzales.prestserv@petrobras.com.br"
CAMINHO_DRIVER_FIXO = r"C:\Users\E797\PycharmProjects\pythonProject\msedgedriver.exe"

PATH_BASE_RELATORIOS = r'C:\Users\E797\Downloads\Teste mensagem e print'
PATH_BASE_PLANILHAS = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches"
PATH_RDS = os.path.join(PATH_BASE_RELATORIOS, 'RDs.xlsx')

PUNCH_LISTS_CONFIG = {
    "Topside": {
        "sharepoint_list_name": "P84/85_TOPSIDE_DR90_Punch_List",
        "url_view": "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/P8485_TOPSIDE_DR90_Punch_List/Updated%20View.aspx",
        "file_path": os.path.join(PATH_BASE_PLANILHAS, "Punch_DR90_TS.xlsx"),
        "columns_to_keep": [
            "DECK No.", "Action Description", "KBR Comment", "Company", "KBR Discipline", "Status",
            "Date Cleared by KBR", "Petrobras Response By", "Petrobras Response Date", "Petrobras Response ",
            "Petrobras Remarks", "Petrobras Discipline", "Petrobras Responsible", "Seatrium Remarks", "Zone",
            "Date Cleared by Petrobras", "S3D Item Tags", "Punch No", "KBR Target Date",
            "Days Since Date Cleared by KBR", "Days Since Date Cleared by Seatrium", "Punched by  (Group)",
            "Petrobras Need Operation to close? (Y/N)", "Date Cleared by Petrobras Operation",
            "Petrobras Operation accept closing? (Y/N)", "Is Reopen? (Y/N)", "Seatrium Target Date Calculated",
            "Petrobras Operation Target Date Calculated", "Petrobras Target Date Calculated",
            "Petrobras Target Date", "Petrobras Operation Target Date", "Seatrium Target Date"
        ],
        "dashboard_path": os.path.join(PATH_BASE_RELATORIOS, 'dashboard_status.png'),
        "op_check_path": os.path.join(PATH_BASE_RELATORIOS, 'Operation to check.xlsx'),
        "esup_check_path": os.path.join(PATH_BASE_RELATORIOS, 'ESUP to check.xlsx'),
        "julius_check_path": os.path.join(PATH_BASE_RELATORIOS, 'Julius to check.xlsx')
    },
    "E-House": {
        "sharepoint_list_name": "DR90 EHouse Punchlist",
        "url_view": "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/DR90%20EHouse%20Punchlist/AllItems.aspx",
        "file_path": os.path.join(PATH_BASE_PLANILHAS, "Punch_DR90_E-House.xlsx"),
        "columns_to_keep": [
            'Punch No', 'Zone', 'DECK No.', 'Zone-Punch Number', 'Action Description', 'Punched by',
            'Punch SnapShot1', 'Punch SnapShot2', 'Closing SnapShot1', 'Hotwork', 'ABB/CIMC Discipline',
            'Company', 'Close Out Plan Date', 'Action by', 'Status', 'Action Comment', 'Date Cleared by ABB',
            'Days Since Date Cleared by ABB', 'KBR Response', 'KBR Response Date', 'KBR Response by',
            'KBR Remarks', 'KBR Category', 'KBR Discipline', 'KBR Screenshot', 'Date Cleared by KBR',
            'Days Since Date Cleared By KBR', 'Seatrium Discipline', 'Seatrium Remarks',
            'Checked By (Seatrium)', 'Seatrium Comments', 'Date Cleared By Seatrium',
            'Days Since Date Cleared by Seatrium', 'Petrobras Response', 'Petrobras Response By',
            'Petrobras Screenshot', 'Petrobras Response Date', 'Petrobras Remarks', 'Petrobras Discipline',
            'Petrobras Category', 'Date Cleared by Petrobras', 'Days Since Date Cleared By Petrobras',
            'Additional Remarks', 'ARC Reference No(HFE Only)', 'Modified', 'Modified By', 'Item Type', 'Path'
        ],
        "dashboard_path": os.path.join(PATH_BASE_RELATORIOS, 'ehouse_status_graph.png')
    },
    "Vendors": {
        "sharepoint_list_name": "Vendor Package Review Punchlist DR90",
        "url_view": "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/Vendor%20Package%20Review%20Punchlist%20DR90/AllItems.aspx",
        "file_path": os.path.join(PATH_BASE_PLANILHAS, "Punch_DR90_Vendors.xlsx"),
        "columns_to_keep": [
            'Punch No', 'Zone', 'DECK No.', 'Zone-Punch Number', 'Action Description', 'S3D Item Tags',
            'Punched by', 'Punch Snapshot', 'Punch Snapshot 2', 'Punch Snapshot 3', 'Punch Snapshot 4',
            'Close-Out Snapshot 1', 'Close-Out Snapshot 2', 'Action Comment', 'Vendor Discipline',
            'Company', 'Action by', 'Status', 'Date Cleared by KBR', 'Days Since Date Cleared by KBR',
            'Petrobras Response', 'Petrobras Response by', 'Petrobras Response Date', 'Petrobras Screenshot',
            'Remarks', 'Petrobras Discipline', 'Petrobras Category', 'Date Cleared by Petrobras',
            'Seatrium Remarks', 'Seatrium Discipline', 'Checked By (Seatrium)', 'Seatrium Comments',
            'Date Cleared By Seatrium', 'Days Since Date Cleared by Seatrium', 'Modified By', 'Item Type', 'Path'
        ],
        "dashboard_path": os.path.join(PATH_BASE_RELATORIOS, 'vendors_status_graph.png')
    }
}


class AutomacaoPunchList:
    def __init__(self):
        self.driver = None
        self.session = None
        self.log_sessao = []

    def registrar_log(self, mensagem, tipo="INFO"):
        timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        texto = f"[{timestamp}] [{tipo}] {mensagem}"
        print(texto)
        self.log_sessao.append(texto)

    def formatar_como_tabela(self, writer, df, sheet_name):
        # ... (código completo da função)
        pass

    def iniciar_sessao_navegador(self):
        # ... (código completo da função)
        pass

    def baixar_todas_as_planilhas(self):
        # ... (código completo da função)
        pass

    def processar_dados_topside(self):
        # ... (código completo da função)
        pass

    def processar_dados_ehouse(self):
        # ... (código completo da função)
        pass

    def processar_dados_vendors(self):
        # ... (código completo da função)
        pass

    def gerar_dashboard_topside(self, dados):
        # ... (código completo da função)
        pass

    def gerar_grafico_ehouse(self, dados):
        # ... (código completo da função)
        pass

    def gerar_dashboard_vendors(self, dados):
        # ... (código completo da função)
        pass

    def enviar_email_topside(self, dados, log_processo):
        # ... (código completo da função)
        pass

    def enviar_email_de_falha(self, log_processo):
        # ... (código completo da função)
        pass

    def enviar_mensagem_julius(self, dados):
        # ... (código completo da função)
        pass

    def enviar_email_ehouse(self, dados):
        # ... (código completo da função)
        pass

    def enviar_email_vendors(self, dados):
        # ... (código completo da função)
        pass

    def executar_ciclo_de_relatorios(self):
        # ... (código completo da função)
        pass

# --- EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    automacao = AutomacaoPunchList()

    # Inicia a sessão do navegador que será usada para os downloads
    if automacao.iniciar_sessao_navegador():

        # Define as tarefas agendadas
        schedule.every(15).minutes.do(automacao.baixar_todas_as_planilhas)
        schedule.every().day.at("08:00").do(automacao.executar_ciclo_de_relatorios)
        schedule.every().day.at("12:00").do(automacao.executar_ciclo_de_relatorios)
        schedule.every().day.at("16:30").do(automacao.executar_ciclo_de_relatorios)

        automacao.registrar_log(">>> AGENDADOR DE AUTOMAÇÃO INICIADO <<<")
        automacao.registrar_log("Executando tarefas pela primeira vez para validação...")

        # Execução inicial
        automacao.baixar_todas_as_planilhas()
        automacao.executar_ciclo_de_relatorios()

        automacao.registrar_log("Validação inicial concluída. Monitorando agendamentos...")

        while True:
            schedule.run_pending()
            time.sleep(60) # Verifica a cada minuto se há tarefas pendentes
    else:
        automacao.registrar_log("CRÍTICO: Não foi possível iniciar a automação. Verifique os logs.", tipo="ERRO")
