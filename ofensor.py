"""
Este script automatiza o processo de geração e envio de relatórios de status
para a Punch List do Design Review do DR90.

O script executa os seguintes fluxos principais:
1.  **Relatório Principal (Topside):** Processa a planilha principal, gera um
    dashboard com os status, cria planilhas de ação para as equipes de
    Operação e Engenharia (ESUP), e envia um e-mail consolidado para o canal
    do Teams.
2.  **E-mail para Julius:** Entre 7h e 9h, envia um e-mail de ação específico
    para o Julius com os itens que a Operação já aprovou e que aguardam o
    fechamento final.
3.  **Relatório E-House:** Processa a planilha de pendências da E-House,
    gera um gráfico de status e envia um e-mail para o canal do Teams.
4.  **Relatório Vendors:** Processa a planilha de pendências de Fornecedores
    (Vendors), gera um dashboard de status e envia um e-mail para o canal
    do Teams.

O script utiliza as bibliotecas pandas para manipulação de dados, matplotlib e
seaborn para a geração de gráficos, e win32com para a automação do envio de
e-mails via Outlook.
"""
import os
import traceback
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import win32com.client as win32

# --- CONFIGURAÇÕES GLOBAIS ---

# --- Caminhos para os arquivos de entrada (planilhas) ---
PATH_PUNCH = r'C:\Users\E797\Downloads\Teste mensagem e print\Punch_DR90_TS.xlsx'
PATH_RDS = r'C:\Users\E797\Downloads\Teste mensagem e print\RDs\RDs.xlsx'
PATH_EHOUSE_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_E-House.xlsx"
PATH_VENDORS_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_Vendors.xlsx"

# --- Caminhos para os arquivos de saída (gráficos e planilhas de ação) ---
PATH_DASHBOARD_IMG = r'C:\Users\E797\Downloads\Teste mensagem e print\dashboard_status.png'
PATH_OP_CHECK = r'C:\Users\E797\Downloads\Teste mensagem e print\Operation to check.xlsx'
PATH_ESUP_CHECK = r'C:\Users\E797\Downloads\Teste mensagem e print\ESUP to check.xlsx'
PATH_JULIUS_CHECK = r'C:\Users\E797\Downloads\Teste mensagem e print\Julius to check.xlsx'
PATH_EHOUSE_GRAPH = r"C:\Users\E797\Downloads\Teste mensagem e print\ehouse_status_graph.png"
PATH_VENDORS_GRAPH = r"C:\Users\E797\Downloads\Teste mensagem e print\vendors_status_graph.png"

# --- Endereços de E-mail ---
EMAIL_DESTINO_TEAMS = "658b4ef7.petrobras.com.br@br.teams.ms"
EMAIL_JULIUS = "julius.lorzales.prestserv@petrobras.com.br"


def log_message(message):
    """Gera uma mensagem de log padronizada com timestamp."""
    return f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"


# --- FUNÇÕES DE PROCESSAMENTO DE DADOS ---

def processar_dados_topside():
    """
    Carrega e processa a planilha principal (Topside) para extrair métricas.
    """
    log = []
    try:
        if not os.path.exists(PATH_PUNCH):
            raise FileNotFoundError(f"Arquivo Topside não encontrado: {PATH_PUNCH}")
        if not os.path.exists(PATH_RDS):
            raise FileNotFoundError(f"Arquivo de RDs não encontrado: {PATH_RDS}")

        df = pd.read_excel(PATH_PUNCH)
        df.columns = df.columns.str.strip()
        df_rds = pd.read_excel(PATH_RDS)
        df_rds.columns = df_rds.columns.str.strip()

        log.append(log_message("Planilhas Topside e RDs carregadas com sucesso."))
        hoje = datetime.now()

        # Filtro principal: Itens com status 'Pending PB Reply'
        pending_pb_reply = df[df['Status'].str.strip() == 'Pending PB Reply'].copy()

        # --- Cálculos e Métricas ---
        status_counts = df['Status'].value_counts().to_dict()
        disciplina_status = pending_pb_reply['Petrobras Discipline'].value_counts().to_dict()

        # Itens que dependem da resposta da Operação
        mask_op_reply = (pending_pb_reply['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                        (pending_pb_reply['Petrobras Operation accept closing? (Y/N)'].isna())
        df_pending_op = pending_pb_reply[mask_op_reply].copy()
        count_pending_op_reply = len(df_pending_op)

        # Itens com prazo de Operação vencido
        df_pending_op['Petrobras Operation Target Date'] = pd.to_datetime(
            df_pending_op['Petrobras Operation Target Date'], dayfirst=True, errors='coerce')
        mask_op_overdue = (df_pending_op['Petrobras Operation Target Date'] < hoje) & \
                          (df_pending_op['Date Cleared by Petrobras Operation'].isna())
        count_op_overdue = len(df_pending_op[mask_op_overdue])

        # Itens com prazo de ESUP vencido
        pending_pb_reply['Petrobras Target Date'] = pd.to_datetime(
            pending_pb_reply['Petrobras Target Date'], dayfirst=True, errors='coerce')
        df_esup_overdue = pending_pb_reply[pending_pb_reply['Petrobras Target Date'] < hoje].copy()
        count_esup_overdue = len(df_esup_overdue)

        # Relacionamento entre prazos vencidos de ESUP e Operação
        overdue_esup_dep_op = df_esup_overdue[df_esup_overdue.index.isin(df_pending_op.index)]
        count_esup_dep_op = len(overdue_esup_dep_op)
        count_esup_indep_op = count_esup_overdue - count_esup_dep_op

        # Contagem de itens avaliados pela Operação
        mask_op_group = df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])
        resp_op_group = len(df[mask_op_group & df['Date Cleared by Petrobras Operation'].notna()])

        mask_eng_group = df['Punched by  (Group)'] == 'PB - Engineering'
        resp_eng_by_op = len(df[mask_eng_group & df['Date Cleared by Petrobras Operation'].notna()])

        # Mapeamento de RDs para menção no e-mail
        disciplinas_pendentes = pending_pb_reply['Petrobras Discipline'].unique()
        mencoes_rds = {f"@{nome}" for disc in disciplinas_pendentes
                       for nome in df_rds[df_rds.iloc[:, 0] == disc].iloc[0, 1:4].dropna().tolist()}

        # --- Geração de DataFrames para Ação ---
        # Ação: Operação precisa responder
        df_op_check = df[(df['Status'].str.strip() == 'Pending PB Reply') &
                         (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) &
                         (df['Date Cleared by Petrobras Operation'].isna())].copy()

        # Ação: ESUP pode responder (prazo da operação vencido ou operação recusou)
        mask_esup_p1 = (df['Status'].str.strip() == 'Pending PB Reply') & \
                       (df['Punched by  (Group)'] == 'PB - Engineering') & \
                       (pd.to_datetime(df['Petrobras Operation Target Date'], dayfirst=True, errors='coerce') < hoje)
        mask_esup_p2 = (df['Status'].str.strip() == 'Pending PB Reply') & \
                       (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                       (df['Petrobras Operation accept closing? (Y/N)'] == False)
        df_esup_check = pd.concat([df[mask_esup_p1], df[mask_esup_p2]]).drop_duplicates().reset_index(drop=True)

        # Ação: Julius precisa fechar (operação aceitou)
        mask_julius = (df['Status'].str.strip() == 'Pending PB Reply') & \
                      (df['Punched by  (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                      (df['Petrobras Operation accept closing? (Y/N)'] == True)
        df_julius_check = df[mask_julius].copy()

        resultados = {
            "total_punches": len(df),
            "status_counts": status_counts,
            "disciplina_status": disciplina_status,
            "pending_op_reply": count_pending_op_reply,
            "op_overdue": count_op_overdue,
            "esup_overdue": count_esup_overdue,
            "esup_dep_op": count_esup_dep_op,
            "esup_indep_op": count_esup_indep_op,
            "resp_op_total": resp_op_group,
            "resp_eng_by_op": resp_eng_by_op,
            "mencoes_rds": " ".join(sorted(list(mencoes_rds))),
            "df_op_check": df_op_check,
            "df_esup_check": df_esup_check,
            "df_julius_check": df_julius_check
        }

        log.append(log_message("Processamento de dados Topside concluído."))
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento Topside: {str(e)}\n{erro_detalhado}")
        return None, log, False


def processar_dados_ehouse():
    """
    Carrega e processa a planilha de pendências da E-House.
    """
    log = []
    try:
        if not os.path.exists(PATH_EHOUSE_PUNCH):
            raise FileNotFoundError(f"Arquivo E-House não encontrado: {PATH_EHOUSE_PUNCH}")

        df_ehouse = pd.read_excel(PATH_EHOUSE_PUNCH)
        df_ehouse.columns = df_ehouse.columns.str.strip()

        pending_petrobras = df_ehouse[df_ehouse['Status'].str.strip() == 'Pending Petrobras'].copy()
        disciplina_counts = pending_petrobras['Petrobras Discipline'].value_counts().to_dict()

        resultados = {
            "total_pending": len(pending_petrobras),
            "disciplina_counts": disciplina_counts
        }
        log.append(log_message("Processamento de dados E-House concluído."))
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento E-House: {str(e)}\n{erro_detalhado}")
        return None, log, False


def processar_dados_vendors():
    """
    Carrega e processa a planilha de pendências de Fornecedores (Vendors).
    """
    log = []
    try:
        if not os.path.exists(PATH_VENDORS_PUNCH):
            raise FileNotFoundError(f"Arquivo Vendors não encontrado: {PATH_VENDORS_PUNCH}")

        df_vendors = pd.read_excel(PATH_VENDORS_PUNCH)
        df_vendors.columns = df_vendors.columns.str.strip()

        pending_petrobras = df_vendors[df_vendors['Status'].str.strip() == 'Pending Petrobras'].copy()
        disciplina_counts = pending_petrobras['Petrobras Discipline'].value_counts().to_dict()

        resultados = {
            "total_pending": len(pending_petrobras),
            "disciplina_counts": disciplina_counts,
            "total_punches": len(df_vendors)
        }
        log.append(log_message("Processamento de dados de Vendors concluído."))
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento de Vendors: {str(e)}\n{erro_detalhado}")
        return None, log, False


# --- FUNÇÕES DE GERAÇÃO DE GRÁFICOS ---

def setup_plot_style():
    """Configura o estilo padrão para os gráficos."""
    sns.set_style("whitegrid")
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = 'Calibri'


def annotate_bars(ax, is_horizontal=False):
    """Adiciona anotações de valor às barras de um gráfico."""
    for p in ax.patches:
        if is_horizontal:
            ax.annotate(f' {int(p.get_width())}', (p.get_width(), p.get_y() + p.get_height() / 2.),
                        ha='left', va='center', fontsize=12, color='black')
        else:
            ax.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()),
                        ha='center', va='center', fontsize=14, color='black', xytext=(0, 10),
                        textcoords='offset points')


def gerar_dashboard_topside(dados):
    """
    Gera a imagem do dashboard para o relatório Topside.
    """
    log = []
    try:
        setup_plot_style()
        cor_principal = "#004488"
        cor_destaque = "#ff8c00"

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8), gridspec_kw={'width_ratios': [1, 2]})
        fig.suptitle('Status Report - Design Review TS', fontsize=24, fontweight='bold', color=cor_principal)

        # Gráfico 1: Visão Geral
        ax1.set_title('Visão Geral dos Itens', fontsize=16, fontweight='bold')
        sns.barplot(x=['Total de Itens', 'Pendentes (PB)'],
                    y=[dados['total_punches'], dados['status_counts'].get('Pending PB Reply', 0)],
                    palette=[cor_principal, cor_destaque], ax=ax1, width=0.5,
                    hue=['Total de Itens', 'Pendentes (PB)'], legend=False)
        ax1.set_ylabel('Quantidade', fontsize=12)
        ax1.grid(axis='y', linestyle='--', alpha=0.7)
        annotate_bars(ax1)

        # Gráfico 2: Pendências por Disciplina
        disciplinas = dados['disciplina_status']
        if disciplinas:
            disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
            nomes = [item[0] for item in disciplinas_sorted]
            valores = [item[1] for item in disciplinas_sorted]
            ax2.set_title('Pendências por Disciplina', fontsize=16, fontweight='bold')
            sns.barplot(x=valores, y=nomes, palette="viridis", ax=ax2, orient='h', hue=nomes, legend=False)
            ax2.set_xlabel('Quantidade de Itens Pendentes', fontsize=12)
            ax2.grid(axis='x', linestyle='--', alpha=0.7)
            annotate_bars(ax2, is_horizontal=True)
        else:
            ax2.set_title('Nenhuma Pendência por Disciplina', fontsize=16, fontweight='bold')
            ax2.text(0.5, 0.5, 'Sem dados para exibir', ha='center', va='center', fontsize=14)
            ax2.set_xticks([])
            ax2.set_yticks([])

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(PATH_DASHBOARD_IMG, dpi=200, bbox_inches='tight')
        plt.close()
        log.append(log_message("Dashboard Topside gerado com sucesso."))
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar dashboard Topside: {str(e)}\n{erro_detalhado}")
        return False, log


def gerar_grafico_ehouse(dados):
    """
    Gera o gráfico de barras para o relatório E-House.
    """
    log = []
    try:
        disciplinas = dados.get('disciplina_counts')
        if not disciplinas:
            log.append("Nenhum dado 'Pending Petrobras' em E-House para gerar gráfico.")
            return True, log

        setup_plot_style()
        plt.figure(figsize=(12, 8))

        disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
        nomes = [item[0] for item in disciplinas_sorted]
        valores = [item[1] for item in disciplinas_sorted]

        ax = sns.barplot(x=nomes, y=valores, palette="Blues_r")
        ax.set_title('Status Punch E-House: Pendentes Petrobras por Disciplina', fontsize=18, fontweight='bold')
        ax.set_xlabel('Disciplina', fontsize=12, fontweight='bold')
        ax.set_ylabel('Quantidade de Itens', fontsize=12, fontweight='bold')
        plt.xticks(rotation=45, ha='right')
        annotate_bars(ax)

        plt.tight_layout()
        plt.savefig(PATH_EHOUSE_GRAPH, dpi=200, bbox_inches='tight')
        plt.close()
        log.append(log_message("Gráfico E-House gerado com sucesso."))
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar gráfico E-House: {str(e)}\n{erro_detalhado}")
        return False, log


def gerar_dashboard_vendors(dados):
    """
    Gera a imagem do dashboard para o relatório de Vendors.
    """
    log = []
    try:
        setup_plot_style()
        cor_principal = "#2E8B57"
        cor_destaque = "#FFD700"

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8), gridspec_kw={'width_ratios': [1, 2]})
        fig.suptitle('Status Report - Vendor Packages DR90', fontsize=24, fontweight='bold', color=cor_principal)

        # Gráfico 1: Visão Geral
        ax1.set_title('Visão Geral dos Itens', fontsize=16, fontweight='bold')
        sns.barplot(x=['Total de Itens', 'Pendentes (PB)'],
                    y=[dados['total_punches'], dados.get('total_pending', 0)],
                    palette=[cor_principal, cor_destaque], ax=ax1, width=0.5,
                    hue=['Total de Itens', 'Pendentes (PB)'], legend=False)
        ax1.set_ylabel('Quantidade', fontsize=12)
        ax1.grid(axis='y', linestyle='--', alpha=0.7)
        annotate_bars(ax1)

        # Gráfico 2: Pendências por Disciplina
        disciplinas = dados['disciplina_counts']
        if disciplinas:
            disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
            nomes = [item[0] for item in disciplinas_sorted]
            valores = [item[1] for item in disciplinas_sorted]
            ax2.set_title('Pendências por Disciplina', fontsize=16, fontweight='bold')
            sns.barplot(x=valores, y=nomes, palette="crest", ax=ax2, orient='h', hue=nomes, legend=False)
            ax2.set_xlabel('Quantidade de Itens Pendentes', fontsize=12)
            ax2.grid(axis='x', linestyle='--', alpha=0.7)
            annotate_bars(ax2, is_horizontal=True)
        else:
            ax2.set_title('Nenhuma Pendência por Disciplina', fontsize=16, fontweight='bold')
            ax2.text(0.5, 0.5, 'Sem dados para exibir', ha='center', va='center', fontsize=14)
            ax2.set_xticks([])
            ax2.set_yticks([])

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(PATH_VENDORS_GRAPH, dpi=200, bbox_inches='tight')
        plt.close()
        log.append(log_message("Dashboard de Vendors gerado com sucesso."))
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar dashboard de Vendors: {str(e)}\n{erro_detalhado}")
        return False, log


# --- FUNÇÕES DE ENVIO DE E-MAIL ---

def enviar_email_outlook(to, subject, html_body, attachments=None, bcc=None, importance=2):
    """
    Função genérica para criar e enviar um e-mail via Outlook.
    """
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to
        if bcc:
            mail.BCC = bcc
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Importance = importance  # 2 = Alta, 1 = Normal, 0 = Baixa

        if attachments:
            for attachment in attachments:
                if os.path.exists(attachment):
                    mail.Attachments.Add(attachment)

        mail.Send()
        print(log_message(f"E-mail '{subject}' enviado para {to}."))
        return True
    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao enviar e-mail '{subject}': {str(e)}\n{erro_detalhado}")
        return False


def enviar_email_relatorio_topside(dados):
    """
    Constrói e envia o e-mail do relatório principal (Topside).
    """
    subject = f"Status Report: Punch List DR90 TS - {datetime.now().strftime('%d/%m/%Y')}"
    disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_status'].items()])
    attachments = [PATH_DASHBOARD_IMG]

    # Seção de Ação para Operação
    secao_op_check_html = ""
    df_op_check = dados.get("df_op_check")
    if df_op_check is not None and not df_op_check.empty:
        df_op_check.to_excel(PATH_OP_CHECK, index=False)
        attachments.append(PATH_OP_CHECK)
        secao_op_check_html = f"""
        <div style="border: 2px solid red; padding: 10px; margin-top: 15px;">
            <p><b style="color:red;">Ponto de Atenção - Operação:</b></p>
            <p>Foram identificados <b>{len(df_op_check)}</b> itens que requerem uma ação da equipe de Operação.
            A planilha <i>'Operation to check.xlsx'</i>, anexada, contém o detalhamento.</p>
        </div>
        """

    # Seção de Ação para ESUP
    secao_esup_check_html = ""
    df_esup_check = dados.get("df_esup_check")
    if df_esup_check is not None and not df_esup_check.empty:
        df_esup_check.to_excel(PATH_ESUP_CHECK, index=False)
        attachments.append(PATH_ESUP_CHECK)
        secao_esup_check_html = f"""
        <div style="border: 2px solid blue; padding: 10px; margin-top: 15px;">
            <p><b style="color:blue;">Ponto de Atenção - ESUP (Engenharia):</b></p>
            <p>Foram identificados <b>{len(df_esup_check)}</b> itens que podem ser respondidos pela Engenharia.
            A planilha <i>'ESUP to check.xlsx'</i>, anexada, contém o detalhamento.</p>
        </div>
        """

    html_body = f"""
    <html lang="pt-BR">
    <head>
        <style>
            body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
            p {{ margin: 10px 0; }}
            table {{ border-collapse: collapse; width: 80%; margin-top: 15px; border: 1px solid #ddd; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; font-weight: bold; }}
            td.center {{ text-align: center; }}
            .highlight {{ color: #c00000; font-weight: bold; }}
            .mention {{ font-weight: bold; color: #005a9e; }}
        </style>
    </head>
    <body>
        <p class="mention">@Acompanhamento Design Review TS</p>
        <p>Prezados,</p>
        <p>Segue a atualização diária das pendências do <b>Design Review TS</b>:</p>
        {secao_op_check_html}
        {secao_esup_check_html}
        <p>Atualmente, temos <span class="highlight">{dados['status_counts'].get('Pending PB Reply', 0)}</span> itens com status <b>Pending PB Reply</b>.</p>
        <p><b>Detalhamento por Disciplina:</b></p>
        <ul>{disciplinas_html}</ul>
        <p><b>Atenção RDs:</b> <span class="mention">{dados['mencoes_rds']}</span></p>
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
        <p><i>O dashboard atualizado e as planilhas de ação (quando aplicável) estão anexados.</i></p>
        <p>Atenciosamente,<br><b>Daniel Alves Anversi - Digital Engineering</b></p>
    </body>
    </html>
    """
    enviar_email_outlook(to=EMAIL_DESTINO_TEAMS, subject=subject, html_body=html_body, attachments=attachments)


def enviar_email_acao_julius(dados):
    """
    Constrói e envia o e-mail de ação para o Julius.
    """
    df_julius_check = dados.get("df_julius_check")
    if df_julius_check is None or df_julius_check.empty:
        print(log_message("Nenhum item para ação do Julius. E-mail não enviado."))
        return

    df_julius_check.to_excel(PATH_JULIUS_CHECK, index=False)
    num_items = len(df_julius_check)
    subject = f"Action Required: {num_items} Punch List Items for Closure - {datetime.now().strftime('%d/%m/%Y')}"
    html_body = f"""
    <html lang="en">
    <head>
        <style>
            body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
            p {{ margin: 10px 0; }}
            .highlight {{ font-weight: bold; color: #004488; }}
        </style>
    </head>
    <body>
        <p>Dear Julius,</p>
        <p>This is an automated notification regarding <span class="highlight">{num_items} punch list items</span> that have been approved by the Operation team and are now awaiting final closure.</p>
        <p>Your action is required to proceed with the final verification for these items.</p>
        <p>The detailed list is attached in the spreadsheet <i>'Julius to check.xlsx'</i> for your convenience.</p>
        <p>Thank you for your attention to this matter.</p>
        <p>Best regards,<br><b>Daniel Alves Anversi - Digital Engineering</b></p>
    </body>
    </html>
    """
    enviar_email_outlook(to=EMAIL_JULIUS, subject=subject, html_body=html_body,
                           attachments=[PATH_JULIUS_CHECK], bcc=EMAIL_DESTINO_TEAMS)


def enviar_email_relatorio_ehouse(dados):
    """
    Constrói e envia o e-mail do relatório da E-House.
    """
    if dados is None or dados.get("total_pending", 0) == 0:
        print(log_message("Nenhum item 'Pending Petrobras' no E-House. E-mail não enviado."))
        return

    subject = f"Status Report: Punch List DR90 E-House - {datetime.now().strftime('%d/%m/%Y')}"
    disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_counts'].items()])
    html_body = f"""
    <html lang="pt-BR">
    <head>
        <style>
            body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
            .highlight {{ color: #c00000; font-weight: bold; }}
            .mention {{ font-weight: bold; color: #005a9e; }}
        </style>
    </head>
    <body>
        <p class="mention">@Acompanhamento Design Review TS</p>
        <p>Prezados,</p>
        <p>Segue a atualização de status da <b>Punch List E-House</b>:</p>
        <p>Atualmente, temos <span class="highlight">{dados['total_pending']}</span> itens com status <b>Pending Petrobras</b>.</p>
        <p><b>Detalhamento por Disciplina:</b></p>
        <ul>{disciplinas_html}</ul>
        <p><i>O gráfico de status está anexado.</i></p>
        <p>Atenciosamente,<br><b>Daniel Alves Anversi - Digital Engineering</b></p>
    </body>
    </html>
    """
    enviar_email_outlook(to=EMAIL_DESTINO_TEAMS, subject=subject, html_body=html_body, attachments=[PATH_EHOUSE_GRAPH])


def enviar_email_relatorio_vendors(dados):
    """
    Constrói e envia o e-mail do relatório de Vendors.
    """
    if dados is None or dados.get("total_pending", 0) == 0:
        print(log_message("Nenhum item 'Pending Petrobras' em Vendors. E-mail não enviado."))
        return

    subject = f"Status Report: Punch List DR90 Vendors - {datetime.now().strftime('%d/%m/%Y')}"
    disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_counts'].items()])
    html_body = f"""
    <html lang="pt-BR">
    <head>
        <style>
            body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
            .highlight {{ color: #c00000; font-weight: bold; }}
            .mention {{ font-weight: bold; color: #005a9e; }}
        </style>
    </head>
    <body>
        <p class="mention">@Acompanhamento Design Review TS</p>
        <p>Prezados,</p>
        <p>Segue a atualização da <b>Punch List de Vendors (Fornecedores)</b>:</p>
        <p>Temos <span class="highlight">{dados['total_pending']}</span> itens com status <b>Pending Petrobras</b>.</p>
        <p><b>Detalhamento por Disciplina:</b></p>
        <ul>{disciplinas_html}</ul>
        <p><i>O dashboard de status está anexado.</i></p>
        <p>Atenciosamente,<br><b>Daniel Alves Anversi - Digital Engineering</b></p>
    </body>
    </html>
    """
    enviar_email_outlook(to=EMAIL_DESTINO_TEAMS, subject=subject, html_body=html_body, attachments=[PATH_VENDORS_GRAPH])


def enviar_email_de_log(log_processo, is_falha=False):
    """
    Envia um e-mail com o log de execução (sucesso ou falha).
    """
    status = "FALHA" if is_falha else "Sucesso"
    subject = f"Log de Execução ({status}) - Automação Punch List - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    body = f"A automação {'falhou' if is_falha else 'foi concluída'} em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    body += "Causa do Erro:\n" if is_falha else "Log de Execução:\n"
    body += "\n".join(log_processo)

    # Reutiliza a função genérica, mas envia como texto simples (sem HTML)
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_DESTINO_TEAMS
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        print(log_message(f"E-mail de log de {status.lower()} enviado."))
    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao tentar enviar o e-mail de log de {status.lower()}: {str(e)}\n{erro_detalhado}")


# --- FLUXOS DE EXECUÇÃO ---

def run_fluxo_topside():
    """Executa o fluxo completo para o relatório Topside."""
    print("\n--- [FLUXO 1/4] Processando Relatório Principal (Topside) ---")
    dados, log, sucesso = processar_dados_topside()
    if not sucesso:
        print("\n!!! FALHA CRÍTICA NO PROCESSAMENTO DOS DADOS TOPSIDE !!!")
        enviar_email_de_log(log, is_falha=True)
        return None  # Retorna None para indicar falha

    print("-> Dados Topside processados com sucesso.")
    sucesso_dash, log_dash = gerar_dashboard_topside(dados)
    log.extend(log_dash)

    if not sucesso_dash:
        print("-> !!! FALHA NA GERAÇÃO DO DASHBOARD TOPSIDE !!!")
        enviar_email_de_log(log, is_falha=True)
        return None

    print("-> Dashboard Topside gerado com sucesso.")
    enviar_email_relatorio_topside(dados)
    enviar_email_de_log(log)
    return dados  # Retorna os dados para o fluxo do Julius


def run_fluxo_julius(dados_topside):
    """Executa o fluxo de envio de e-mail para o Julius, se aplicável."""
    print("\n--- [FLUXO 2/4] Verificando E-mail para Julius ---")
    hora_atual = datetime.now().hour
    if not (7 <= hora_atual < 9):
        print(f"-> Fora do horário agendado (executado às {hora_atual}h). E-mail para Julius não será enviado.")
        return

    if dados_topside is None:
        print("-> O processamento de dados do Topside falhou, e-mail para Julius não pôde ser gerado.")
        return

    enviar_email_acao_julius(dados_topside)


def run_fluxo_ehouse():
    """Executa o fluxo completo para o relatório E-House."""
    print("\n--- [FLUXO 3/4] Processando Relatório E-House ---")
    try:
        dados, log, sucesso = processar_dados_ehouse()
        if not sucesso:
            raise RuntimeError("Falha no processamento de dados E-House.")

        print("-> Dados E-House processados com sucesso.")
        sucesso_grafico, log_grafico = gerar_grafico_ehouse(dados)
        log.extend(log_grafico)

        if not sucesso_grafico:
            raise RuntimeError("Falha na geração do gráfico E-House.")

        print("-> Gráfico E-House gerado com sucesso.")
        enviar_email_relatorio_ehouse(dados)

    except FileNotFoundError as e:
        print(f"-> {e}. O relatório para este fluxo não será gerado.")
    except Exception as e:
        print(f"\n!!! FALHA CRÍTICA NO FLUXO E-HOUSE: {e} !!!")
        enviar_email_de_log([str(e)], is_falha=True)


def run_fluxo_vendors():
    """Executa o fluxo completo para o relatório de Vendors."""
    print("\n--- [FLUXO 4/4] Processando Relatório Vendors ---")
    try:
        dados, log, sucesso = processar_dados_vendors()
        if not sucesso:
            raise RuntimeError("Falha no processamento de dados de Vendors.")

        print("-> Dados de Vendors processados com sucesso.")
        sucesso_dash, log_dash = gerar_dashboard_vendors(dados)
        log.extend(log_dash)

        if not sucesso_dash:
            raise RuntimeError("Falha na geração do dashboard de Vendors.")

        print("-> Dashboard de Vendors gerado com sucesso.")
        enviar_email_relatorio_vendors(dados)

    except FileNotFoundError as e:
        print(f"-> {e}. O relatório para este fluxo não será gerado.")
    except Exception as e:
        print(f"\n!!! FALHA CRÍTICA NO FLUXO DE VENDORS: {e} !!!")
        enviar_email_de_log([str(e)], is_falha=True)


# --- EXECUÇÃO PRINCIPAL ---
def main():
    """Função principal que orquestra a execução dos fluxos."""
    print(f"--- INICIANDO PROCESSO DE AUTOMAÇÃO GERAL ({datetime.now().strftime('%d/%m/%Y %H:%M:%S')}) ---")

    # Fluxo 1 é crítico e seus dados são usados no Fluxo 2
    dados_topside = run_fluxo_topside()

    # Outros fluxos são independentes
    run_fluxo_julius(dados_topside)
    run_fluxo_ehouse()
    run_fluxo_vendors()

    print(f"\n--- PROCESSO DE AUTOMAÇÃO GERAL FINALIZADO ({datetime.now().strftime('%d/%m/%Y %H:%M:%S')}) ---")


if __name__ == "__main__":
    main()
