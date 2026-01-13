# coding: utf-8
import os
import sys
import time
import traceback
import urllib3
import requests
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns

try:
    import win32com.client as win32
except ImportError:
    print("ERRO: Instale a biblioteca pywin32 executando: pip install pywin32")
    sys.exit(1)

from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURAÇÕES GERAIS ---
PASTA_RAIZ = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches"
PASTA_TEMP = r'C:\Users\E797\Downloads\Teste mensagem e print'
os.makedirs(PASTA_RAIZ, exist_ok=True)
os.makedirs(PASTA_TEMP, exist_ok=True)

# --- CONFIGURAÇÕES DAS LISTAS DO SHAREPOINT ---
LISTAS_SHAREPOINT = {
    "TS": {
        "url": "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/P8485_TOPSIDE_DR90_Punch_List/Updated%20View.aspx",
        "base_site_url": "https://seatrium.sharepoint.com/sites/P84P85DesignReview",
        "api_name": "P84/85_TOPSIDE_DR90_Punch_List",
        "output_file": "Punch_DR90_TS.xlsx",
        "colunas": [
            "DECK No.", "Action Description", "KBR Comment", "Company", "KBR Discipline",
            "Status", "Date Cleared by KBR", "Petrobras Response By", "Petrobras Response Date",
            "Petrobras Response ", "Petrobras Remarks", "Petrobras Discipline",
            "Petrobras Responsible", "Seatrium Remarks", "Zone", "Date Cleared by Petrobras",
            "S3D Item Tags", "Punch No", "KBR Target Date", "Days Since Date Cleared by KBR",
            "Days Since Date Cleared by Seatrium", "Punched by  (Group)",
            "Petrobras Need Operation to close? (Y/N)", "Date Cleared by Petrobras Operation",
            "Petrobras Operation accept closing? (Y/N)", "Is Reopen? (Y/N)",
            "Seatrium Target Date Calculated", "Petrobras Operation Target Date Calculated",
            "Petrobras Target Date Calculated", "Petrobras Target Date",
            "Petrobras Operation Target Date", "Seatrium Target Date"
        ]
    },
    "E-House": {
        "url": "https://seatrium.sharepoint.com/:l:/r/sites/P84P85DesignReview/Lists/DR90%20EHouse%20Punchlist?e=QCVEQf",
        "base_site_url": "https://seatrium.sharepoint.com/sites/P84P85DesignReview",
        "api_name": "DR90 EHouse Punchlist",
        "output_file": "Punch_DR90_E-House.xlsx",
        "colunas": [
            "Punch No", "Zone", "DECK No.", "Zone-Punch Number", "Action Description", "Punched by",
            "Punch SnapShot1", "Punch SnapShot2", "Closing SnapShot1", "Hotwork", "ABB/CIMC Discipline",
            "Company", "Close Out Plan Date", "Action by", "Status", "Action Comment", "Date Cleared by ABB",
            "Days Since Date Cleared by ABB", "KBR Response", "KBR Response Date", "KBR Response by",
            "KBR Remarks", "KBR Category", "KBR Discipline", "KBR Screenshot", "Date Cleared by KBR",
            "Days Since Date Cleared By KBR", "Seatrium Discipline", "Seatrium Remarks",
            "Checked By (Seatrium)", "Seatrium Comments", "Date Cleared By Seatrium",
            "Days Since Date Cleared by Seatrium", "Petrobras Response", "Petrobras Response By",
            "Petrobras Screenshot", "Petrobras Response Date", "Petrobras Remarks", "Petrobras Discipline",
            "Petrobras Category", "Date Cleared by Petrobras", "Days Since Date Cleared By Petrobras",
            "Additional Remarks", "ARC Reference No(HFE Only)", "Modified", "Modified By", "Item Type", "Path"
        ]
    },
    "Vendors": {
        "url": "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/Vendor%20Package%20Review%20Punchlist%20DR90/AllItems.aspx?e=4tHLty&CID=43904b9e%2D7cb2%2D481c%2Db136%2D5285ae014bd9&ovuser=5b6f6241%2D9a57%2D4be4%2D8e50%2D1dfa72e79a57%2Cleojunqueira%40petrobras%2Ecom%2Ebr",
        "base_site_url": "https://seatrium.sharepoint.com/sites/P84P85DesignReview",
        "api_name": "Vendor Package Review Punchlist DR90",
        "output_file": "Punch_DR90_Vendors.xlsx",
        "colunas": [
            "Punch No", "Zone", "DECK No.", "Zone-Punch Number", "Action Description", "S3D Item Tags",
            "Punched by", "Punch Snapshot", "Punch Snapshot 2", "Punch Snapshot 3", "Punch Snapshot 4",
            "Close-Out Snapshot 1", "Close-Out Snapshot 2", "Action Comment", "Vendor Discipline",
            "Company", "Action by", "Status", "Date Cleared by KBR", "Days Since Date Cleared by KBR",
            "Petrobras Response", "Petrobras Response by", "Petrobras Response Date", "Petrobras Screenshot",
            "Remarks", "Petrobras Discipline", "Petrobras Category", "Date Cleared by Petrobras",
            "Seatrium Remarks", "Seatrium Discipline", "Checked By (Seatrium)", "Seatrium Comments",
            "Date Cleared By Seatrium", "Days Since Date Cleared by Seatrium", "Modified By", "Item Type", "Path"
        ]
    }
}

# --- CAMINHOS DINÂMICOS ---
PATH_RDS = os.path.join(PASTA_TEMP, 'RDs', 'RDs.xlsx')
PATH_OP_CHECK = os.path.join(PASTA_TEMP, 'Operation to check.xlsx')
PATH_ESUP_CHECK = os.path.join(PASTA_TEMP, 'ESUP to check.xlsx')
PATH_JULIUS_CHECK = os.path.join(PASTA_TEMP, 'Julius to check.xlsx')
PATH_FECHAMENTO_OPERACAO_GRAPH = os.path.join(PASTA_TEMP, "fechamento_operacao_por_dia.png")
CAMINHO_DRIVER_FIXO = r"C:\Users\E797\PycharmProjects\pythonProject\msedgedriver.exe"

# --- E-MAILS ---
EMAIL_DESTINO_TEAMS = "658b4ef7.petrobras.com.br@br.teams.ms"
EMAIL_JULIUS = "julius.lorzales.prestserv@petrobras.com.br"

# --- FUNÇÕES DE ANÁLISE E RELATÓRIO ---
def processar_dados_geral(log_ext, path_planilha):
    log = list(log_ext)
    try:
        if not os.path.exists(path_planilha):
            raise FileNotFoundError(f"Arquivo não encontrado: {path_planilha}")

        df = pd.read_excel(path_planilha)
        df.columns = df.columns.str.strip()

        pending_petrobras = df[df['Status'].str.strip() == 'Pending Petrobras'].copy()
        disciplina_counts = pending_petrobras.get('Petrobras Discipline', pd.Series(dtype=str)).value_counts().to_dict()

        resultados = {
            "total_punches": len(df),
            "total_pending": len(pending_petrobras),
            "disciplina_counts": disciplina_counts
        }
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Processamento de '{os.path.basename(path_planilha)}' concluído.")
        return resultados, log, True
    except Exception as e:
        log.append(f"ERRO CRÍTICO em '{os.path.basename(path_planilha)}': {traceback.format_exc()}")
        return None, log, False

def processar_dados_ts(log_ext, path_planilha):
    log = list(log_ext)
    try:
        if not os.path.exists(path_planilha): raise FileNotFoundError(f"Arquivo não encontrado: {path_planilha}")
        if not os.path.exists(PATH_RDS): raise FileNotFoundError(f"Arquivo RDs não encontrado: {PATH_RDS}")
        df = pd.read_excel(path_planilha)
        df.columns = df.columns.str.strip()
        df_rds = pd.read_excel(PATH_RDS)
        df_rds.columns = df_rds.columns.str.strip()
        hoje = datetime.now()
        df['Date Cleared by Petrobras Operation'] = pd.to_datetime(df['Date Cleared by Petrobras Operation'], dayfirst=True, errors='coerce')
        fechamentos_diarios = df[df['Date Cleared by Petrobras Operation'].notna()].groupby(df['Date Cleared by Petrobras Operation'].dt.date).size()

        status_counts = df['Status'].value_counts().to_dict()
        pending_pb_reply = df[df['Status'].str.strip() == 'Pending PB Reply'].copy()
        disciplina_status = pending_pb_reply['Petrobras Discipline'].value_counts().to_dict()
        mask_op_reply = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Petrobras Operation accept closing? (Y/N)'].isna())
        df_pending_op = df[mask_op_reply].copy()
        df_pending_op['Petrobras Operation Target Date'] = pd.to_datetime(df_pending_op['Petrobras Operation Target Date'], dayfirst=True, errors='coerce')
        mask_op_overdue = (df_pending_op['Petrobras Operation Target Date'] < hoje) & (df_pending_op['Date Cleared by Petrobras Operation'].isna())
        pending_pb_reply['Petrobras Target Date'] = pd.to_datetime(pending_pb_reply['Petrobras Target Date'], dayfirst=True, errors='coerce')
        df_esup_overdue = pending_pb_reply[pending_pb_reply['Petrobras Target Date'] < hoje].copy()
        overdue_esup_dep_op = df_esup_overdue[df_esup_overdue.index.isin(df_pending_op.index)]
        mask_op_group = df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])
        mask_eng_group = df['Punched by  (Group)'] == 'PB - Engineering'
        disciplinas_pendentes = pending_pb_reply['Petrobras Discipline'].unique()
        mencoes_rds = [f"@{nome}" for disc in disciplinas_pendentes for row in [df_rds[df_rds.iloc[:, 0] == disc]] if not row.empty for nome in row.iloc[0, 1:4].dropna().tolist()]
        mask_op_check = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Date Cleared by Petrobras Operation'].isna())
        mask_esup_p1 = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'] == 'PB - Engineering') & (pd.to_datetime(df['Petrobras Operation Target Date'], dayfirst=True, errors='coerce') < hoje)
        mask_esup_p2 = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Petrobras Operation accept closing? (Y/N)'] == False)
        mask_julius = (df['Status'].str.strip() == 'Pending PB Reply') & (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & (df['Petrobras Operation accept closing? (Y/N)'] == True)

        resultados = {
            "total_punches": len(df), "status_counts": status_counts, "disciplina_status": disciplina_status,
            "pending_op_reply": len(df_pending_op), "op_overdue": len(df_pending_op[mask_op_overdue]), "esup_overdue": len(df_esup_overdue),
            "esup_dep_op": len(overdue_esup_dep_op), "esup_indep_op": len(df_esup_overdue) - len(overdue_esup_dep_op),
            "resp_op_total": len(df[mask_op_group & df['Date Cleared by Petrobras Operation'].notna()]),
            "resp_eng_by_op": len(df[mask_eng_group & df['Date Cleared by Petrobras Operation'].notna()]),
            "mencoes_rds": " ".join(sorted(list(set(mencoes_rds)))), "df_op_check": df[mask_op_check].copy(),
            "df_esup_check": pd.concat([df[mask_esup_p1].copy(), df[mask_esup_p2].copy()]).drop_duplicates().reset_index(drop=True),
            "df_julius_check": df[mask_julius].copy(), "fechamentos_diarios": fechamentos_diarios
        }
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Processamento TS concluído.")
        return resultados, log, True
    except Exception as e:
        log.append(f"ERRO CRÍTICO em TS: {traceback.format_exc()}")
        return None, log, False

def gerar_dashboard_ts(dados, log_ext, path_imagem):
    log = list(log_ext)
    try:
        total_punches = dados['total_punches']
        pending_reply = dados['status_counts'].get('Pending PB Reply', 0)
        disciplinas = dados['disciplina_status']
        sns.set_style("whitegrid")
        plt.rcParams.update({'font.family': 'sans-serif', 'font.sans-serif': 'Calibri'})
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8), gridspec_kw={'width_ratios': [1, 2]})
        fig.suptitle('Status Report - Design Review TS', fontsize=24, fontweight='bold', color="#004488")
        ax1.set_title('Visão Geral dos Itens', fontsize=16, fontweight='bold')
        sns.barplot(x=['Total de Itens', 'Pendentes (PB)'], y=[total_punches, pending_reply], palette=["#004488", "#ff8c00"], ax=ax1, width=0.5)
        for p in ax1.patches:
            ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()), ha='center', va='center', fontsize=14, color='black', xytext=(0, 10), textcoords='offset points')
        if disciplinas:
            disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
            ax2.set_title('Pendências por Disciplina', fontsize=16, fontweight='bold')
            sns.barplot(x=[v for k, v in disciplinas_sorted], y=[k for k, v in disciplinas_sorted], palette="viridis", ax=ax2, orient='h')
            for i, v in enumerate([item[1] for item in disciplinas_sorted]):
                ax2.text(v, i, f' {v}', va='center', fontsize=12)
        else:
            ax2.text(0.5, 0.5, 'Sem dados para exibir', ha='center', va='center', fontsize=14)
        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(path_imagem, dpi=200, bbox_inches='tight')
        plt.close()
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Dashboard TS gerado.")
        return True, log
    except Exception as e:
        log.append(f"ERRO ao gerar dashboard TS: {traceback.format_exc()}")
        return False, log

def gerar_grafico_fechamento_operacao(dados, log_ext):
    log = list(log_ext)
    try:
        fechamentos_diarios = dados.get("fechamentos_diarios")
        if fechamentos_diarios is None or fechamentos_diarios.empty:
            log.append("Nenhum item fechado pela operação para gerar gráfico.")
            return True, log
        start_date = fechamentos_diarios.index.min()
        end_date = datetime.now().date()
        date_range = pd.date_range(start=start_date, end=end_date, freq='D')
        fechamentos_completos = fechamentos_diarios.reindex(date_range.date, fill_value=0)
        sns.set_style("whitegrid")
        plt.figure(figsize=(15, 8))
        ax = sns.barplot(x=fechamentos_completos.index, y=fechamentos_completos.values, color="#3498db")
        ax.set_title('Fechamentos Diários pela Operação', fontsize=18, fontweight='bold')
        ax.set_xlabel('Data', fontsize=12)
        ax.set_ylabel('Quantidade de Itens Fechados', fontsize=12)
        plt.xticks(rotation=45, ha='right')
        ax.xaxis.set_major_formatter(plt.FixedFormatter(fechamentos_completos.index.strftime('%d/%m/%Y')))
        for p in ax.patches:
            if p.get_height() > 0:
                 ax.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()), ha='center', va='bottom', fontsize=11)
        plt.tight_layout()
        plt.savefig(PATH_FECHAMENTO_OPERACAO_GRAPH, dpi=200, bbox_inches='tight')
        plt.close()
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Gráfico de fechamento por operação gerado.")
        return True, log
    except Exception as e:
        log.append(f"ERRO ao gerar gráfico de fechamento: {traceback.format_exc()}")
        return False, log

def gerar_grafico_disciplinas(dados, log_ext, path_imagem, titulo):
    log = list(log_ext)
    try:
        disciplinas = dados.get('disciplina_counts', {})
        if not disciplinas:
            log.append(f"Nenhum dado de disciplina para gerar gráfico '{titulo}'.")
            return True, log

        sns.set_style("whitegrid")
        plt.figure(figsize=(12, 8))
        disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)

        ax = sns.barplot(x=[k for k, v in disciplinas_sorted], y=[v for k, v in disciplinas_sorted], palette="viridis")
        ax.set_title(titulo, fontsize=18, fontweight='bold')
        ax.set_xlabel('Disciplina', fontsize=12)
        ax.set_ylabel('Quantidade de Itens Pendentes', fontsize=12)
        plt.xticks(rotation=45, ha='right')

        for p in ax.patches:
            ax.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()), ha='center', va='bottom', fontsize=11)

        plt.tight_layout()
        plt.savefig(path_imagem, dpi=200, bbox_inches='tight')
        plt.close()
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Gráfico '{titulo}' gerado com sucesso.")
        return True, log
    except Exception as e:
        log.append(f"ERRO ao gerar gráfico '{titulo}': {traceback.format_exc()}")
        return False, log

def enviar_email_geral(dados, log_ext, path_grafico, titulo_email, titulo_corpo):
    log = list(log_ext)
    if dados is None or dados.get("total_pending", 0) == 0:
        log.append(f"Nenhum item pendente para '{titulo_corpo}'. E-mail não enviado.")
        return log
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Importance = 2
        mail.To = EMAIL_DESTINO_TEAMS
        mail.Subject = f"{titulo_email} - {datetime.now().strftime('%d/%m/%Y')}"
        disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_counts'].items()])
        mail.HTMLBody = f"""<p>Prezados,</p><p>Segue a atualização de status da <b>{titulo_corpo}</b>:</p><p>Atualmente, temos <b style='color: #c00000;'>{dados['total_pending']}</b> itens com status <b>Pending Petrobras</b>.</p><p><b>Detalhamento por Disciplina:</b></p><ul>{disciplinas_html}</ul><p><i>O gráfico de status está anexado.</i></p>"""
        if os.path.exists(path_grafico):
            mail.Attachments.Add(path_grafico)
        mail.Send()
        log.append(f"E-mail para '{titulo_corpo}' enviado com sucesso.")
    except Exception as e:
        log.append(f"ERRO ao enviar e-mail para '{titulo_corpo}': {traceback.format_exc()}")
    return log

def enviar_email_principal(dados, log_processo, path_dashboard_ts):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Importance = 2
        mail.To = EMAIL_DESTINO_TEAMS
        mail.Subject = f"Status Report: Punch List DR90 TS - {datetime.now().strftime('%d/%m/%Y')}"

        if os.path.exists(path_dashboard_ts): mail.Attachments.Add(path_dashboard_ts)
        if os.path.exists(PATH_FECHAMENTO_OPERACAO_GRAPH): mail.Attachments.Add(PATH_FECHAMENTO_OPERACAO_GRAPH)

        df_op_check = dados.get("df_op_check")
        if df_op_check is not None and not df_op_check.empty:
            df_op_check.to_excel(PATH_OP_CHECK, index=False)
            mail.Attachments.Add(PATH_OP_CHECK)

        df_esup_check = dados.get("df_esup_check")
        if df_esup_check is not None and not df_esup_check.empty:
            df_esup_check.to_excel(PATH_ESUP_CHECK, index=False)
            mail.Attachments.Add(PATH_ESUP_CHECK)

        disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_status'].items()])
        secao_op_check_html = ""
        if df_op_check is not None and not df_op_check.empty:
            secao_op_check_html = f"""<div style="border: 2px solid red; padding: 10px; margin-top: 15px;"><p><b style="color:red;">Ponto de Atenção - Operação:</b></p><p>Foram identificados <b>{len(df_op_check)}</b> itens que requerem uma ação da equipe de Operação. A planilha <i>'Operation to check.xlsx'</i>, anexada, contém o detalhamento.</p></div>"""
        secao_esup_check_html = ""
        if df_esup_check is not None and not df_esup_check.empty:
            secao_esup_check_html = f"""<div style="border: 2px solid blue; padding: 10px; margin-top: 15px;"><p><b style="color:blue;">Ponto de Atenção - ESUP (Engenharia):</b></p><p>Foram identificados <b>{len(df_esup_check)}</b> itens que podem ser respondidos pela Engenharia. A planilha <i>'ESUP to check.xlsx'</i>, anexada, contém o detalhamento.</p></div>"""

        mail.HTMLBody = f"""
        <html lang="pt-BR">
        <head><style>body{{font-family:Calibri,sans-serif;font-size:11pt}}p{{margin:10px 0}}table{{border-collapse:collapse;width:80%;margin-top:15px;border:1px solid #ddd}}th,td{{border:1px solid #ddd;padding:8px;text-align:left}}th{{background-color:#f2f2f2;font-weight:bold}}td.center{{text-align:center}}.highlight{{color:#c00000;font-weight:bold}}.mention{{font-weight:bold;color:#005a9e}}</style></head>
        <body>
            <p class="mention">@Acompanhamento Design Review TS</p>
            <p>Prezados,</p>
            <p>Segue a atualização diária das pendências do <b>Design Review TS</b>:</p>
            {secao_op_check_html}{secao_esup_check_html}
            <p>Atualmente, temos <span class="highlight">{dados['status_counts'].get('Pending PB Reply', 0)}</span> itens com status <b>Pending PB Reply</b>.</p>
            <p><b>Detalhamento por Disciplina:</b></p><ul>{disciplinas_html}</ul>
            <p><b>Atenção RDs:</b><span class="mention">{dados['mencoes_rds']}</span></p>
            <table>
                <tr><th>Indicador</th><th>Quantidade</th></tr>
                <tr><td>Itens Pendentes de Resposta da Operação</td><td class="center">{dados['pending_op_reply']}</td></tr>
                <tr><td>Itens com Prazo de Operação Vencido</td><td class="center">{dados['op_overdue']}</td></tr>
                <tr><td>Itens com Prazo ESUP Vencido</td><td class="center">{dados['esup_overdue']}</td></tr>
                <tr><td>Overdue ESUP com Dependência da Operação</td><td class="center">{dados['esup_dep_op']}</td></tr>
                <tr><td>Overdue ESUP sem Dependência da Operação</td><td class="center">{dados['esup_indep_op']}</td></tr>
                <tr><td>Itens Mandatórios Avaliados pela Operação</td><td class="center">{dados['resp_op_total']}</td></tr>
                <tr><td>Itens de Engenharia Avaliados pela Operação</td><td class="center">{dados['resp_eng_by_op']}</td></tr>
            </table>
            <p><i>O dashboard, gráfico de fechamento e planilhas de ação (quando aplicável) estão anexados.</i></p>
            <p>Atenciosamente,</p><p><b>Daniel Alves Anversi - Digital Engineering</b></p>
        </body></html>"""

        mail.Send()
    except Exception as e:
        print(f"ERRO ao enviar e-mail principal: {traceback.format_exc()}")

def enviar_email_de_falha(log_processo):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_DESTINO_TEAMS
        mail.Subject = f"Log de Execução (FALHA) - Automação Punch List - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        mail.Body = "A automação falhou.\n\nCausa do Erro:\n" + "\n".join(log_processo)
        mail.Send()
    except Exception as e:
        print(f"ERRO ao tentar enviar o e-mail de falha: {e}")

def enviar_mensagem_julius(dados):
    df_julius_check = dados.get("df_julius_check")
    if df_julius_check is None or df_julius_check.empty: return
    try:
        df_julius_check.to_excel(PATH_JULIUS_CHECK, index=False)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_JULIUS
        mail.BCC = EMAIL_DESTINO_TEAMS
        mail.Subject = f"Action Required: {len(df_julius_check)} Punch List Items for Closure - {datetime.now().strftime('%d/%m/%Y')}"
        mail.Importance = 2
        mail.HTMLBody = f"""<p>Dear Julius,</p><p>This is an automated notification regarding <b>{len(df_julius_check)} punch list items</b> that have been approved by the Operation team and are now awaiting final closure.</p><p>The detailed list is attached in the spreadsheet <i>'Julius to check.xlsx'</i>.</p><p>Best regards,</p><p><b>Daniel Alves Anversi - Digital Engineering</b></p>"""
        mail.Attachments.Add(PATH_JULIUS_CHECK)
        mail.Send()
    except Exception as e:
        print(f"ERRO ao enviar e-mail para Julius: {e}")

class AutomacaoPunchList:
    def __init__(self):
        self.driver = None
        self.log_sessao = []
        self.mapeamentos_colunas = {}
        self.primeira_execucao_do_dia = True

    def registrar_log(self, mensagem):
        timestamp = datetime.now().strftime('%H:%M:%S')
        texto = f"[{timestamp}] {mensagem}"
        print(texto)
        self.log_sessao.append(texto)

    def enviar_log_geral(self, sucesso):
        status = "SUCESSO" if sucesso else "FALHA"
        corpo = f"RELATÓRIO DE EXECUÇÃO\nData: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n" + "\n".join(self.log_sessao)
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = EMAIL_DESTINO_TEAMS
            mail.Subject = f"Log de Execução Automação Punch List - {status}"
            mail.Body = corpo
            mail.Send()
        except Exception as e:
            self.registrar_log(f"Falha ao enviar log geral: {e}")

    def salvar_excel(self, df, caminho_arquivo):
        try:
            df.to_excel(caminho_arquivo, index=False)
            self.registrar_log(f"Planilha salva com sucesso em: {caminho_arquivo}")
            return True
        except PermissionError:
            self.registrar_log(f"ERRO DE PERMISSÃO: O arquivo '{caminho_arquivo}' está aberto. Feche-o para salvar.")
            return False
        except Exception as e:
            self.registrar_log(f"ERRO ao salvar '{caminho_arquivo}': {e}")
            return False

    def tratar_dados(self, df, nome_lista):
        mapeamento = self.mapeamentos_colunas.get(nome_lista, {})
        if mapeamento:
            df = df.rename(columns=mapeamento)

        colunas_desejadas = LISTAS_SHAREPOINT[nome_lista]["colunas"]
        colunas_presentes = [c for c in colunas_desejadas if c in df.columns]
        df = df[colunas_presentes].copy()

        for col in [c for c in df.columns if "Date" in c]:
            df[col] = pd.to_datetime(df[col], format='%Y-%m-%dT%H:%M:%SZ', errors='coerce').dt.strftime('%d/%m/%Y')

        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.replace("error", "", case=False).fillna('')

        df.replace(['NaT', 'nan', 'None'], '', inplace=True)

        return df

    def obter_mapeamento_colunas(self, session, base_url, nome_lista, api_name):
        endpoint = f"{base_url}/_api/web/lists/getbytitle('{api_name}')/fields"
        try:
            response = session.get(endpoint, headers={"Accept": "application/json;odata=verbose"}, timeout=30)
            if response.status_code == 200:
                self.mapeamentos_colunas[nome_lista] = {f['InternalName']: f['Title'] for f in response.json()['d']['results']}
        except Exception as e:
            self.registrar_log(f"Erro no mapeamento de '{nome_lista}': {e}")

    def iniciar_sessao_navegador(self):
        try:
            service = EdgeService(executable_path=CAMINHO_DRIVER_FIXO)
            options = Options()
            self.driver = webdriver.Edge(service=service, options=options)
            self.driver.get(LISTAS_SHAREPOINT["TS"]["url"])
            WebDriverWait(self.driver, 180).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[role='grid']")))
            self.registrar_log("Sessão autenticada.")
            return True
        except Exception as e:
            self.registrar_log(f"Erro ao iniciar navegador: {e}")
            return False

    def extrair_dados_de_lista(self, session, nome_lista, config):
        self.registrar_log(f"Extraindo dados para: {nome_lista}")
        try:
            base_site_url = config["base_site_url"]
            if not self.mapeamentos_colunas.get(nome_lista):
                self.obter_mapeamento_colunas(session, base_site_url, nome_lista, config["api_name"])

            endpoint = f"{base_site_url}/_api/web/lists/getbytitle('{config['api_name']}')/items?$top=5000"
            response = session.get(endpoint, headers={"Accept": "application/json;odata=verbose"}, timeout=60)

            if response.status_code == 200:
                results = response.json().get('d', {}).get('results', [])
                if not results:
                    self.registrar_log(f"AVISO: A lista '{nome_lista}' retornou vazia.")
                    df_vazio = pd.DataFrame(columns=config["colunas"])
                    self.salvar_excel(df_vazio, os.path.join(PASTA_RAIZ, config["output_file"]))
                    return True

                df_raw = pd.json_normalize(results)
                df_final = self.tratar_dados(df_raw, nome_lista)
                return self.salvar_excel(df_final, os.path.join(PASTA_RAIZ, config["output_file"]))
            else:
                self.registrar_log(f"ERRO de API ao extrair '{nome_lista}': Status {response.status_code} - {response.text[:200]}")
                return False
        except Exception as e:
            self.registrar_log(f"ERRO CRÍTICO na extração de '{nome_lista}': {traceback.format_exc()}")
            return False

    def ciclo_de_download(self):
        self.registrar_log("Iniciando ciclo de download...")
        sucesso_total = True
        cookies = self.driver.get_cookies()
        with requests.Session() as session:
            session.verify = False
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'])
            for nome, config in LISTAS_SHAREPOINT.items():
                if not config.get("base_site_url"):
                    self.registrar_log(f"Pulando '{nome}' (URL base não configurada).")
                    continue
                if not self.extrair_dados_de_lista(session, nome, config):
                    sucesso_total = False
        return sucesso_total

    def executar_analises(self):
        self.registrar_log("--- INICIANDO ROTINAS DE ANÁLISE E ENVIO ---")
        hora_atual = datetime.now().hour

        # --- FLUXO 1: Relatório Principal (Topside) ---
        self.registrar_log("\n--- [FLUXO 1/3] Processando Relatório Principal (Topside) ---")
        path_ts = os.path.join(PASTA_RAIZ, LISTAS_SHAREPOINT['TS']['output_file'])
        dados_ts, log_ts, sucesso_ts = processar_dados_ts(self.log_sessao, path_ts)
        self.log_sessao = log_ts
        if sucesso_ts:
            path_dashboard_ts = os.path.join(PASTA_TEMP, 'ts_dashboard.png')
            sucesso_dash, log_dash = gerar_dashboard_ts(dados_ts, self.log_sessao, path_dashboard_ts)
            self.log_sessao = log_dash

            sucesso_fech, log_fech = gerar_grafico_fechamento_operacao(dados_ts, self.log_sessao)
            self.log_sessao = log_fech

            enviar_email_principal(dados_ts, self.log_sessao, path_dashboard_ts)

            if self.primeira_execucao_do_dia or (7 <= hora_atual < 9):
                enviar_mensagem_julius(dados_ts)
        else:
            enviar_email_de_falha(self.log_sessao)

        # --- FLUXOS 2 & 3: E-House e Vendors (agendados) ---
        if self.primeira_execucao_do_dia or hora_atual in [8, 12, 17]:
            # Análise E-House
            self.registrar_log("\n--- [FLUXO 2/3] Processando Relatório E-House ---")
            path_ehouse = os.path.join(PASTA_RAIZ, LISTAS_SHAREPOINT['E-House']['output_file'])
            path_grafico_ehouse = os.path.join(PASTA_TEMP, 'ehouse_status_graph.png')
            dados_eh, log_eh, sucesso_eh = processar_dados_geral(self.log_sessao, path_ehouse)
            self.log_sessao = log_eh
            if sucesso_eh:
                sucesso_graph, log_graph = gerar_grafico_disciplinas(dados_eh, self.log_sessao, path_grafico_ehouse, "Status Punch E-House: Pendentes por Disciplina")
                self.log_sessao = log_graph
                if sucesso_graph:
                    self.log_sessao = enviar_email_geral(dados_eh, self.log_sessao, path_grafico_ehouse, "Status Punch E-House", "Punch List E-House")
            else:
                enviar_email_de_falha(self.log_sessao)

            # Análise Vendors
            self.registrar_log("\n--- [FLUXO 3/3] Processando Relatório Vendors ---")
            path_vendors = os.path.join(PASTA_RAIZ, LISTAS_SHAREPOINT['Vendors']['output_file'])
            path_grafico_vendors = os.path.join(PASTA_TEMP, 'vendors_status_graph.png')
            dados_ven, log_ven, sucesso_ven = processar_dados_geral(self.log_sessao, path_vendors)
            self.log_sessao = log_ven
            if sucesso_ven:
                sucesso_graph, log_graph = gerar_grafico_disciplinas(dados_ven, self.log_sessao, path_grafico_vendors, "Status Punch Vendors: Pendentes por Disciplina")
                self.log_sessao = log_graph
                if sucesso_graph:
                    self.log_sessao = enviar_email_geral(dados_ven, self.log_sessao, path_grafico_vendors, "Status Punch Vendors", "Punch List de Vendors")
            else:
                enviar_email_de_falha(self.log_sessao)

    def iniciar(self):
        if self.iniciar_sessao_navegador():
            while True:
                self.log_sessao = []
                sucesso_download = self.ciclo_de_download()
                if sucesso_download:
                    self.executar_analises()

                self.enviar_log_geral(sucesso_download)
                if self.primeira_execucao_do_dia: self.primeira_execucao_do_dia = False

                self.registrar_log("Aguardando 15 minutos...")
                time.sleep(900)
        else:
            self.registrar_log("Falha crítica ao iniciar o navegador. Automação encerrada.")
            self.enviar_log_geral(False)

if __name__ == "__main__":
    automacao = AutomacaoPunchList()
    automacao.iniciar()
