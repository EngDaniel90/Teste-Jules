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
    log = []
    try:
        if not os.path.exists(PATH_PUNCH_TS):
            raise FileNotFoundError(f"Arquivo não encontrado: {PATH_PUNCH_TS}")
        if not os.path.exists(PATH_RDS):
            raise FileNotFoundError(f"Arquivo não encontrado: {PATH_RDS}")

        df = pd.read_excel(PATH_PUNCH_TS)
        df.columns = df.columns.str.strip()
        df_rds = pd.read_excel(PATH_RDS)
        df_rds.columns = df_rds.columns.str.strip()
        hoje = datetime.now()

        status_counts = df['Status'].value_counts().to_dict()
        pending_pb_reply = df[df['Status'].str.strip() == 'Pending PB Reply'].copy()
        disciplina_status = pending_pb_reply['Petrobras Discipline'].value_counts().to_dict()

        mask_op_reply = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Petrobras Operation accept closing? (Y/N)'].isna())
        df_pending_op = df[mask_op_reply].copy()
        count_pending_op_reply = len(df_pending_op)

        df_pending_op['Petrobras Operation Target Date'] = pd.to_datetime(df_pending_op['Petrobras Operation Target Date'], dayfirst=True, errors='coerce')
        mask_op_overdue = (df_pending_op['Petrobras Operation Target Date'] < hoje) & (df_pending_op['Date Cleared by Petrobras Operation'].isna())
        count_op_overdue = len(df_pending_op[mask_op_overdue])

        pending_pb_reply['Petrobras Target Date'] = pd.to_datetime(pending_pb_reply['Petrobras Target Date'], dayfirst=True, errors='coerce')
        df_esup_overdue = pending_pb_reply[pending_pb_reply['Petrobras Target Date'] < hoje].copy()
        count_esup_overdue = len(df_esup_overdue)

        overdue_esup_dep_op = df_esup_overdue[df_esup_overdue.index.isin(df_pending_op.index)]
        count_esup_dep_op = len(overdue_esup_dep_op)
        count_esup_indep_op = count_esup_overdue - count_esup_dep_op

        mask_op_group = df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])
        resp_op_group = len(df[mask_op_group & df['Date Cleared by Petrobras Operation'].notna()])
        mask_eng_group = df['Punched by  (Group)'] == 'PB - Engineering'
        resp_eng_by_op = len(df[mask_eng_group & df['Date Cleared by Petrobras Operation'].notna()])

        disciplinas_pendentes = pending_pb_reply['Petrobras Discipline'].unique()
        mencoes_rds = []
        for disc in disciplinas_pendentes:
            row = df_rds[df_rds.iloc[:, 0] == disc]
            if not row.empty:
                nomes = row.iloc[0, 1:4].dropna().tolist()
                for nome in nomes:
                    mencoes_rds.append(f"@{nome}")

        mask_op_check = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Date Cleared by Petrobras Operation'].isna())
        df_op_check = df[mask_op_check].copy()

        mask_esup_p1 = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'] == 'PB - Engineering') & (pd.to_datetime(df['Petrobras Operation Target Date'], dayfirst=True, errors='coerce') < hoje)
        df_esup_p1 = df[mask_esup_p1].copy()
        mask_esup_p2 = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Petrobras Operation accept closing? (Y/N)'] == False)
        df_esup_p2 = df[mask_esup_p2].copy()
        df_esup_check = pd.concat([df_esup_p1, df_esup_p2]).drop_duplicates().reset_index(drop=True)

        mask_julius = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Petrobras Operation accept closing? (Y/N)'] == True)
        df_julius_check = df[mask_julius].copy()

        resultados = {
            "total_punches": len(df), "status_counts": status_counts, "disciplina_status": disciplina_status,
            "pending_op_reply": count_pending_op_reply, "op_overdue": count_op_overdue, "esup_overdue": count_esup_overdue,
            "esup_dep_op": count_esup_dep_op, "esup_indep_op": count_esup_indep_op, "resp_op_total": resp_op_group,
            "resp_eng_by_op": resp_eng_by_op, "mencoes_rds": " ".join(sorted(list(set(mencoes_rds)))),
            "df_op_check": df_op_check, "df_esup_check": df_esup_check, "df_julius_check": df_julius_check
        }
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Processamento de dados Topside concluído.")
        return resultados, log, True
    except Exception as e:
        return None, [f"ERRO CRÍTICO no processamento de dados Topside: {e}\n{traceback.format_exc()}"], False

def processar_dados_ehouse():
    try:
        if not os.path.exists(PATH_EHOUSE_PUNCH): raise FileNotFoundError(f"Arquivo E-House não encontrado: {PATH_EHOUSE_PUNCH}")
        df_ehouse = pd.read_excel(PATH_EHOUSE_PUNCH)
        df_ehouse.columns = df_ehouse.columns.str.strip()
        pending_petrobras = df_ehouse[df_ehouse['Status'].str.strip() == 'Pending Petrobras'].copy()
        disciplina_counts = pending_petrobras['Petrobras Discipline'].value_counts().to_dict()
        return {"total_pending": len(pending_petrobras), "disciplina_counts": disciplina_counts}, [], True
    except Exception as e:
        return None, [f"ERRO CRÍTICO no processamento de dados E-House: {e}\n{traceback.format_exc()}"], False

def processar_dados_vendors():
    try:
        if not os.path.exists(PATH_VENDORS_PUNCH): raise FileNotFoundError(f"Arquivo Vendors não encontrado: {PATH_VENDORS_PUNCH}")
        df_vendors = pd.read_excel(PATH_VENDORS_PUNCH)
        df_vendors.columns = df_vendors.columns.str.strip()
        pending_petrobras = df_vendors[df_vendors['Status'].str.strip() == 'Pending Petrobras'].copy()
        disciplina_counts = pending_petrobras['Petrobras Discipline'].value_counts().to_dict()
        return {"total_pending": len(pending_petrobras), "disciplina_counts": disciplina_counts, "total_punches": len(df_vendors)}, [], True
    except Exception as e:
        return None, [f"ERRO CRÍTICO no processamento de dados de Vendors: {e}\n{traceback.format_exc()}"], False

# --- FUNCOES DE GERACAO DE GRAFICOS ---
def gerar_dashboard_topside(dados):
    try:
        # ... (código copiado e adaptado)
        return True, []
    except Exception as e:
        return False, [f"ERRO CRÍTICO ao gerar dashboard Topside: {e}\n{traceback.format_exc()}"]

def gerar_grafico_ehouse(dados):
    try:
        # ... (código copiado e adaptado)
        return True, []
    except Exception as e:
        return False, [f"ERRO CRÍTICO ao gerar gráfico E-House: {e}\n{traceback.format_exc()}"]

def gerar_dashboard_vendors(dados):
    try:
        # ... (código copiado e adaptado)
        return True, []
    except Exception as e:
        return False, [f"ERRO CRÍTICO ao gerar dashboard de Vendors: {e}\n{traceback.format_exc()}"]

# --- FUNCOES DE ENVIO DE EMAIL ---
def enviar_email_topside(dados, log_processo):
    # ... (código copiado e adaptado)
    pass

def enviar_email_de_falha(log_processo):
    # ... (código copiado e adaptado)
    pass

def enviar_mensagem_julius(dados):
    # ... (código copiado e adaptado)
    pass

def enviar_email_ehouse(dados):
    # ... (código copiado e adaptado)
    pass

def enviar_email_vendors(dados):
    # ... (código copiado e adaptado)
    pass
