import pandas as pd
import traceback
from datetime import datetime
import win32com.client as win32
import matplotlib.pyplot as plt
import seaborn as sns
import os

# --- CONFIGURAÇÕES DE CAMINHOS E URLs ---
PATH_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_TS.xlsx"
PATH_RDS = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\RDs_ESUP.xlsx"
PATH_DASHBOARD_IMG = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\dashboard_status.png"
PATH_OP_CHECK = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\Operation to check.xlsx"
PATH_ESUP_CHECK = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\ESUP to check.xlsx"
PATH_JULIUS_CHECK = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\Julius to check.xlsx"
PATH_EHOUSE_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_E-House.xlsx"
PATH_EHOUSE_GRAPH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\ehouse_status_graph.png"
PATH_VENDORS_PUNCH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Punch_DR90_Vendors.xlsx"
PATH_VENDORS_GRAPH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\vendors_status_graph.png"
PATH_LAST_RUN = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\last_run.txt"
PATH_FECHAMENTO_GRAPH = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches\Arquivos_de_apoio\fechamento_operacao.png"
EMAIL_DESTINO = "658b4ef7.petrobras.com.br@br.teams.ms"
EMAIL_JULIUS = "julius.lorzales.prestserv@petrobras.com.br"
SCHEDULED_HOURS = [7, 12, 18]


def processar_dados_ehouse():
    """
    Processa os dados da planilha E-House para o relatório específico.
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
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Processamento de dados E-House concluído.")
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento de dados E-House: {str(e)}\n{erro_detalhado}")
        return None, log, False


def processar_dados_vendors():
    """
    Processa os dados da planilha de Vendors para o relatório específico.
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
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Processamento de dados de Vendors concluído.")
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento de dados de Vendors: {str(e)}\n{erro_detalhado}")
        return None, log, False


def processar_dados():
    """
    Carrega os dados das planilhas, processa as métricas e retorna um dicionário com os resultados.
    """
    log = []
    try:
        # 1. Carregamento e Validação dos Arquivos
        if not os.path.exists(PATH_PUNCH):
            raise FileNotFoundError(f"Arquivo não encontrado: {PATH_PUNCH}")
        if not os.path.exists(PATH_RDS):
            raise FileNotFoundError(f"Arquivo não encontrado: {PATH_RDS}")

        df = pd.read_excel(PATH_PUNCH)
        df.columns = df.columns.str.strip()

        df_rds = pd.read_excel(PATH_RDS)
        df_rds.columns = df_rds.columns.str.strip()

        hoje = datetime.now()
        log.append(f"[{hoje.strftime('%Y-%m-%d %H:%M:%S')}] Planilhas carregadas com sucesso.")

        # 2. Contagem de Status Geral
        status_counts = df['Status'].value_counts().to_dict()

        # 3. Disciplina x Status (Pending PB Reply)
        pending_pb_reply = df[df['Status'].str.strip() == 'Pending PB Reply'].copy()
        disciplina_status = pending_pb_reply['Petrobras Discipline'].value_counts().to_dict()

        # 4. Pending Operation Reply
        mask_op_reply = (df['Status'].str.strip() == 'Pending PB Reply') & \
                        (df['Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                        (df['Petrobras Operation accept closing? (Y/N)'].isna())
        df_pending_op = df[mask_op_reply].copy()
        count_pending_op_reply = len(df_pending_op)

        # 5. Petrobras Operation Overdue
        df_pending_op['Petrobras Operation Target Date'] = pd.to_datetime(
            df_pending_op['Petrobras Operation Target Date'], dayfirst=True, errors='coerce')

        mask_op_overdue = (df_pending_op['Petrobras Operation Target Date'] < hoje) & \
                          (df_pending_op['Date Cleared by Petrobras Operation'].isna())
        count_op_overdue = len(df_pending_op[mask_op_overdue])

        # 6. Petrobras ESUP Overdue
        pending_pb_reply['Petrobras Target Date'] = pd.to_datetime(
            pending_pb_reply['Petrobras Target Date'], dayfirst=True, errors='coerce')

        df_esup_overdue = pending_pb_reply[pending_pb_reply['Petrobras Target Date'] < hoje].copy()
        count_esup_overdue = len(df_esup_overdue)

        # 7. Relacionamento ESUP x Operação
        overdue_esup_dep_op = df_esup_overdue[df_esup_overdue.index.isin(df_pending_op.index)]
        count_esup_dep_op = len(overdue_esup_dep_op)
        count_esup_indep_op = count_esup_overdue - count_esup_dep_op

        # 8. Grupos de Avaliação
        mask_op_group = df['Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])
        resp_op_group = len(df[mask_op_group & df['Date Cleared by Petrobras Operation'].notna()])

        mask_eng_group = df['Punched by (Group)'] == 'PB - Engineering'
        resp_eng_by_op = len(df[mask_eng_group & df['Date Cleared by Petrobras Operation'].notna()])

        # 9. Mapeamento de RDs para Menção (@)
        disciplinas_pendentes = pending_pb_reply['Petrobras Discipline'].unique()
        mencoes_rds = []
        for disc in disciplinas_pendentes:
            row = df_rds[df_rds.iloc[:, 0] == disc]
            if not row.empty:
                nomes = row.iloc[0, 1:4].dropna().tolist()
                for nome in nomes:
                    mencoes_rds.append(f"@{nome}")

        # --- 10. Geração de Dataframes para Planilhas ---

        # Itens pendentes de resposta OBRIGATÓRIA da operação
        mask_op_check = (df['Status'].str.strip() == 'Pending PB Reply') & \
                        (df['Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                        (df['Date Cleared by Petrobras Operation'].isna())
        df_op_check = df[mask_op_check].copy()

        # Itens para ESUP checar (Parte 1: Engenharia com prazo de operação vencido)
        mask_esup_p1 = (df['Status'].str.strip() == 'Pending PB Reply') & \
                       (df['Punched by (Group)'] == 'PB - Engineering') & \
                       (pd.to_datetime(df['Petrobras Operation Target Date'], dayfirst=True, errors='coerce') < hoje)
        df_esup_p1 = df[mask_esup_p1].copy()

        # Itens para ESUP checar (Parte 2: Operação respondeu 'False')
        mask_esup_p2 = (df['Status'].str.strip() == 'Pending PB Reply') & \
                       (df['Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                       (df['Petrobras Operation accept closing? (Y/N)'] == False)
        df_esup_p2 = df[mask_esup_p2].copy()

        df_esup_check = pd.concat([df_esup_p1, df_esup_p2]).drop_duplicates().reset_index(drop=True)

        # Itens para Julius checar (Operação respondeu 'True')
        mask_julius = (df['Status'].str.strip() == 'Pending PB Reply') & \
                      (df['Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                      (df['Petrobras Operation accept closing? (Y/N)'] == True)
        df_julius_check = df[mask_julius].copy()

        # --- Consolidação dos Resultados ---
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
            "mencoes_rds": " ".join(sorted(list(set(mencoes_rds)))),
            "df_op_check": df_op_check,
            "df_esup_check": df_esup_check,
            "df_julius_check": df_julius_check,
            "df_full": df
        }

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Processamento de dados concluído.")
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO no processamento de dados: {str(e)}\n{erro_detalhado}")
        log.append(f"ERRO CRÍTICO no processamento de dados: {str(e)}\n{erro_detalhado}")
        return None, log, False


def gerar_dashboard_imagem(dados):
    """
    Gera uma imagem de dashboard com os principais indicadores usando Matplotlib e Seaborn.
    """
    log = []
    try:
        total_punches = dados['total_punches']
        pending_reply = dados['status_counts'].get('Pending PB Reply', 0)
        disciplinas = dados['disciplina_status']

        # --- Configurações Visuais ---
        sns.set_style("whitegrid")
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = 'Calibri'

        cor_principal = "#004488"
        cor_destaque = "#ff8c00"

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8), gridspec_kw={'width_ratios': [1, 2]})
        fig.suptitle('Status Report - Design Review TS', fontsize=24, fontweight='bold', color=cor_principal)

        # --- Gráfico 1: Barras Verticais (Total vs. Pendente) ---
        ax1.set_title('Visão Geral dos Itens', fontsize=16, fontweight='bold')
        sns.barplot(x=['Total de Itens', 'Pendentes (PB)'], y=[total_punches, pending_reply],
                    palette=[cor_principal, cor_destaque], ax=ax1, width=0.5, hue=['Total de Itens', 'Pendentes (PB)'],
                    legend=False)
        ax1.set_ylabel('Quantidade', fontsize=12)
        ax1.grid(axis='y', linestyle='--', alpha=0.7)

        for p in ax1.patches:
            ax1.annotate(f'{int(p.get_height())}',
                         (p.get_x() + p.get_width() / 2., p.get_height()),
                         ha='center', va='center', fontsize=14, color='black', xytext=(0, 10),
                         textcoords='offset points')

        # --- Gráfico 2: Barras Horizontais (Pendências por Disciplina) ---
        if disciplinas:
            disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
            nomes_disciplinas = [item[0] for item in disciplinas_sorted]
            valores_disciplinas = [item[1] for item in disciplinas_sorted]

            ax2.set_title('Pendências por Disciplina', fontsize=16, fontweight='bold')
            sns.barplot(x=valores_disciplinas, y=nomes_disciplinas, palette="viridis", ax=ax2, orient='h',
                        hue=nomes_disciplinas, legend=False)
            ax2.set_xlabel('Quantidade de Itens Pendentes', fontsize=12)
            ax2.grid(axis='x', linestyle='--', alpha=0.7)

            for index, value in enumerate(valores_disciplinas):
                ax2.text(value, index, f' {value}', va='center', fontsize=12, color='black')
        else:
            ax2.set_title('Nenhuma Pendência por Disciplina', fontsize=16, fontweight='bold')
            ax2.text(0.5, 0.5, 'Sem dados para exibir', ha='center', va='center', fontsize=14)
            ax2.set_xticks([])
            ax2.set_yticks([])

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(PATH_DASHBOARD_IMG, dpi=200, bbox_inches='tight')
        plt.close()

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Dashboard gerado com sucesso.")
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar dashboard: {str(e)}\n{erro_detalhado}")
        return False, log


def gerar_grafico_ehouse(dados):
    """
    Gera um gráfico de barras vertical para o status de E-House.
    """
    log = []
    try:
        disciplinas = dados['disciplina_counts']
        if not disciplinas:
            log.append("Nenhum dado de E-House para gerar gráfico.")
            return True, log  # Não é um erro, apenas não há o que fazer

        sns.set_style("whitegrid")
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = 'Calibri'

        cor_principal = "#0072B2"  # Um azul diferente para distinguir

        plt.figure(figsize=(12, 8))

        disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
        nomes_disciplinas = [item[0] for item in disciplinas_sorted]
        valores_disciplinas = [item[1] for item in disciplinas_sorted]

        ax = sns.barplot(x=nomes_disciplinas, y=valores_disciplinas, palette="Blues_r")

        ax.set_title('Status Punch E-House: Pendentes Petrobras por Disciplina', fontsize=18, fontweight='bold')
        ax.set_xlabel('Disciplina', fontsize=12, fontweight='bold')
        ax.set_ylabel('Quantidade de Itens', fontsize=12, fontweight='bold')
        plt.xticks(rotation=45, ha='right')

        for p in ax.patches:
            ax.annotate(f'{int(p.get_height())}',
                        (p.get_x() + p.get_width() / 2., p.get_height()),
                        ha='center', va='center', fontsize=11, color='black', xytext=(0, 5),
                        textcoords='offset points')

        plt.tight_layout()
        plt.savefig(PATH_EHOUSE_GRAPH, dpi=200, bbox_inches='tight')
        plt.close()

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Gráfico E-House gerado com sucesso.")
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar gráfico E-House: {str(e)}\n{erro_detalhado}")
        return False, log


def enviar_email_ehouse(dados):
    """
    Envia um e-mail de status específico para a punch list E-House.
    """
    if dados is None or dados.get("total_pending", 0) == 0:
        print(
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Nenhum item 'Pending Petrobras' no E-House. E-mail não enviado.")
        return

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Importance = 2
        mail.To = EMAIL_DESTINO
        mail.Subject = f"Status Report: Punch List DR90 E-House - {datetime.now().strftime('%d/%m/%Y')}"

        disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_counts'].items()])

        mail.HTMLBody = f"""
        <html lang="pt-BR">
        <head>
            <style>
                body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
                p {{ margin: 10px 0; }}
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

            <p><i>O gráfico de status está anexado a este e-mail.</i></p>
            <p>Atenciosamente,</p>
            <p><b>Daniel Alves Anversi - Digital Engineering</b></p>
        </body>
        </html>
        """

        if os.path.exists(PATH_EHOUSE_GRAPH):
            mail.Attachments.Add(PATH_EHOUSE_GRAPH)

        mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail de status E-House enviado com sucesso.")

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao enviar e-mail de E-House: {str(e)}\n{erro_detalhado}")


def gerar_dashboard_vendors(dados):
    """
    Gera uma imagem de dashboard para o status de Vendors.
    """
    log = []
    try:
        total_punches = dados['total_punches']
        pending_reply = dados.get('total_pending', 0)
        disciplinas = dados['disciplina_counts']

        sns.set_style("whitegrid")
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = 'Calibri'

        cor_principal = "#2E8B57"  # Verde Mar
        cor_destaque = "#FFD700"  # Dourado

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8), gridspec_kw={'width_ratios': [1, 2]})
        fig.suptitle('Status Report - Vendor Packages DR90', fontsize=24, fontweight='bold', color=cor_principal)

        ax1.set_title('Visão Geral dos Itens', fontsize=16, fontweight='bold')
        sns.barplot(x=['Total de Itens', 'Pendentes (PB)'], y=[total_punches, pending_reply],
                    palette=[cor_principal, cor_destaque], ax=ax1, width=0.5, hue=['Total de Itens', 'Pendentes (PB)'],
                    legend=False)
        ax1.set_ylabel('Quantidade', fontsize=12)
        ax1.grid(axis='y', linestyle='--', alpha=0.7)

        for p in ax1.patches:
            ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()),
                         ha='center', va='center', fontsize=14, color='black', xytext=(0, 10),
                         textcoords='offset points')

        if disciplinas:
            disciplinas_sorted = sorted(disciplinas.items(), key=lambda item: item[1], reverse=True)
            nomes_disciplinas = [item[0] for item in disciplinas_sorted]
            valores_disciplinas = [item[1] for item in disciplinas_sorted]
            ax2.set_title('Pendências por Disciplina', fontsize=16, fontweight='bold')
            sns.barplot(x=valores_disciplinas, y=nomes_disciplinas, palette="crest", ax=ax2, orient='h',
                        hue=nomes_disciplinas, legend=False)
            ax2.set_xlabel('Quantidade de Itens Pendentes', fontsize=12)
            ax2.grid(axis='x', linestyle='--', alpha=0.7)
            for index, value in enumerate(valores_disciplinas):
                ax2.text(value, index, f' {value}', va='center', fontsize=12, color='black')
        else:
            ax2.set_title('Nenhuma Pendência por Disciplina', fontsize=16, fontweight='bold')
            ax2.text(0.5, 0.5, 'Sem dados para exibir', ha='center', va='center', fontsize=14)
            ax2.set_xticks([]);
            ax2.set_yticks([])

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(PATH_VENDORS_GRAPH, dpi=200, bbox_inches='tight')
        plt.close()

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Dashboard de Vendors gerado com sucesso.")
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar dashboard de Vendors: {str(e)}\n{erro_detalhado}")
        return False, log


def gerar_grafico_fechamento_operacao(df):
    """
    Gera um gráfico de barras mostrando a quantidade de itens que a operação fechou por dia.
    """
    log = []
    try:
        df_cleaned = df.dropna(subset=['Date Cleared by Petrobras Operation']).copy()
        df_cleaned['Date Cleared'] = pd.to_datetime(df_cleaned['Date Cleared by Petrobras Operation']).dt.date

        # Contagem de fechamentos por dia
        fechamentos_por_dia = df_cleaned['Date Cleared'].value_counts().sort_index()

        # Garantir que todos os dias no intervalo de datas estejam presentes
        if not fechamentos_por_dia.empty:
            date_range = pd.date_range(start=fechamentos_por_dia.index.min(), end=fechamentos_por_dia.index.max(),
                                       freq='D')
            fechamentos_por_dia = fechamentos_por_dia.reindex(date_range.date, fill_value=0)

        # Geração do Gráfico
        plt.figure(figsize=(15, 8))
        ax = sns.barplot(x=fechamentos_por_dia.index, y=fechamentos_por_dia.values, color="#005a9e")

        ax.set_title('Desempenho de Fechamento de Itens pela Operação', fontsize=18, fontweight='bold')
        ax.set_xlabel('Data', fontsize=12, fontweight='bold')
        ax.set_ylabel('Quantidade de Itens Fechados', fontsize=12, fontweight='bold')
        plt.xticks(rotation=45, ha='right')

        # Formatar o eixo x para mostrar as datas de forma mais limpa
        ax.xaxis.set_major_formatter(plt.FixedFormatter(fechamentos_por_dia.index.strftime('%d/%m/%Y')))
        ax.figure.autofmt_xdate()

        for p in ax.patches:
            if p.get_height() > 0:
                ax.annotate(f'{int(p.get_height())}',
                            (p.get_x() + p.get_width() / 2., p.get_height()),
                            ha='center', va='center', fontsize=11, color='black', xytext=(0, 5),
                            textcoords='offset points')

        plt.tight_layout()
        plt.savefig(PATH_FECHAMENTO_GRAPH, dpi=200, bbox_inches='tight')
        plt.close()

        log.append(
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Gráfico de fechamento pela operação gerado com sucesso.")
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO ao gerar gráfico de fechamento: {str(e)}\n{erro_detalhado}")
        return False, log


def enviar_email_vendors(dados):
    """
    Envia um e-mail de status específico para a punch list de Vendors.
    """
    if dados is None or dados.get("total_pending", 0) == 0:
        print(
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Nenhum item 'Pending Petrobras' em Vendors. E-mail não enviado.")
        return

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Importance = 2
        mail.To = EMAIL_DESTINO
        mail.Subject = f"Status Report: Punch List DR90 Vendors - {datetime.now().strftime('%d/%m/%Y')}"

        disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_counts'].items()])

        mail.HTMLBody = f"""
        <html lang="pt-BR">
        <head>
            <style>
                body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
                p {{ margin: 10px 0; }}
                .highlight {{ color: #c00000; font-weight: bold; }}
                .mention {{ font-weight: bold; color: #005a9e; }}
            </style>
        </head>
        <body>
            <p class="mention">@Acompanhamento Design Review TS</p>
            <p>Prezados,</p>
            <p>Segue a atualização de status da <b>Punch List de Vendors (Fornecedores)</b>:</p>

            <p>Atualmente, temos <span class="highlight">{dados['total_pending']}</span> itens com status <b>Pending Petrobras</b>.</p>

            <p><b>Detalhamento por Disciplina:</b></p>
            <ul>{disciplinas_html}</ul>

            <p><i>O dashboard de status está anexado a este e-mail.</i></p>
            <p>Atenciosamente,</p>
            <p><b>Daniel Alves Anversi - Digital Engineering</b></p>
        </body>
        </html>
        """

        if os.path.exists(PATH_VENDORS_GRAPH):
            mail.Attachments.Add(PATH_VENDORS_GRAPH)

        mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail de status Vendors enviado com sucesso.")

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao enviar e-mail de Vendors: {str(e)}\n{erro_detalhado}")


def enviar_email(dados, log_processo):
    """
    Cria e envia um e-mail formatado com os dados do relatório e o log de execução.
    """
    try:
        outlook = win32.Dispatch('outlook.application')

        mail = outlook.CreateItem(0)
        mail.Importance = 2
        mail.To = EMAIL_DESTINO
        mail.Subject = f"Status Report: Punch List DR90 TS - {datetime.now().strftime('%d/%m/%Y')}"

        disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_status'].items()])

        secao_op_check_html = ""
        df_op_check = dados.get("df_op_check")
        if df_op_check is not None and not df_op_check.empty:
            df_op_check.to_excel(PATH_OP_CHECK, index=False)
            mail.Attachments.Add(PATH_OP_CHECK)
            secao_op_check_html = f"""
            <div style="border: 2px solid red; padding: 10px; margin-top: 15px;">
                <p><b style="color:red;">Ponto de Atenção - Operação:</b></p>
                <p>Foram identificados <b>{len(df_op_check)}</b> itens que requerem uma ação necessária da equipe de Operação.
                A planilha <i>'Operation to check.xlsx'</i>, anexada a este e-mail, contém o detalhamento completo.</p>
            </div>
            """

        secao_esup_check_html = ""
        df_esup_check = dados.get("df_esup_check")
        if df_esup_check is not None and not df_esup_check.empty:
            df_esup_check.to_excel(PATH_ESUP_CHECK, index=False)
            mail.Attachments.Add(PATH_ESUP_CHECK)
            secao_esup_check_html = f"""
            <div style="border: 2px solid blue; padding: 10px; margin-top: 15px;">
                <p><b style="color:blue;">Ponto de Atenção - ESUP (Engenharia):</b></p>
                <p>Foram identificados <b>{len(df_esup_check)}</b> itens que agora podem ser respondidos pela equipe de Engenharia.
                A planilha <i>'ESUP to check.xlsx'</i>, anexada a este e-mail, contém o detalhamento completo.</p>
            </div>
            """

        mail.HTMLBody = f"""
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

            <p><b>Atenção RDs:</b>
<span class="mention">{dados['mencoes_rds']}</span></p>

            <table>
                <tr>
                    <th>Indicador</th>
                    <th>Quantidade</th>
                </tr>
                <tr><td>Itens Pendentes de Resposta da Operação (Pending Operation Reply)</td><td class="center">{dados['pending_op_reply']}</td></tr>
                <tr><td>Itens com Prazo de Operação Vencido (Petrobras Operation Overdue)</td><td class="center">{dados['op_overdue']}</td></tr>
                <tr><td>Itens com Prazo ESUP Vencido (Petrobras ESUP Overdue)</td><td class="center">{dados['esup_overdue']}</td></tr>
                <tr><td>Overdue ESUP com Dependência da Operação</td><td class="center">{dados['esup_dep_op']}</td></tr>
                <tr><td>Overdue ESUP sem Dependência da Operação</td><td class="center">{dados['esup_indep_op']}</td></tr>
                <tr><td>Itens Mandatórios Avaliados pela Operação</td><td class="center">{dados['resp_op_total']}</td></tr>
                <tr><td>Itens de Engenharia Avaliados pela Operação</td><td class="center">{dados['resp_eng_by_op']}</td></tr>
            </table>

            <p><i>O dashboard atualizado e as planilhas de ação (quando aplicável) estão anexados a este e-mail.</i></p>
            <p>Atenciosamente,</p>
            <p><b>Daniel Alves Anversi - Digital Engineering</b></p>
        </body>
        </html>
        """

        if os.path.exists(PATH_DASHBOARD_IMG):
            mail.Attachments.Add(PATH_DASHBOARD_IMG)
        if os.path.exists(PATH_FECHAMENTO_GRAPH):
            mail.Attachments.Add(PATH_FECHAMENTO_GRAPH)

        mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail principal enviado para {EMAIL_DESTINO}.")

        log_mail = outlook.CreateItem(0)
        log_mail.To = EMAIL_DESTINO
        log_mail.Subject = f"Log de Execução (Sucesso) - Automação Punch List - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        log_mail.Body = f"Execução concluída com sucesso em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n" + "\n".join(
            log_processo)
        log_mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail de log de sucesso enviado.")

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao enviar e-mail: {str(e)}\n{erro_detalhado}")


def enviar_email_de_falha(log_processo):
    """
    Envia um e-mail de notificação de falha com o log do erro.
    """
    try:
        outlook = win32.Dispatch('outlook.application')
        log_mail = outlook.CreateItem(0)
        log_mail.To = EMAIL_DESTINO
        log_mail.Subject = f"Log de Execução (FALHA) - Automação Punch List - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        log_mail.Body = (f"A automação falhou em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                         "Causa do Erro:\n" + "\n".join(log_processo))
        log_mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail de log de falha enviado.")
    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao tentar enviar o e-mail de falha: {str(e)}\n{erro_detalhado}")


def enviar_mensagem_julius(dados):
    """
    Cria e envia um e-mail de ação específico para o Julius.
    """
    df_julius_check = dados.get("df_julius_check")
    if df_julius_check is None or df_julius_check.empty:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Nenhum item para Julius. E-mail não enviado.")
        return

    try:
        df_julius_check.to_excel(PATH_JULIUS_CHECK, index=False)

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_JULIUS
        mail.CC = EMAIL_DESTINO
        mail.Subject = f"Action Required: {len(df_julius_check)} Punch List Items for Closure - {datetime.now().strftime('%d/%m/%Y')}"
        mail.Importance = 2

        mail.HTMLBody = f"""
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
            <p>This is an automated notification regarding <span class="highlight">{len(df_julius_check)} punch list items</span> that have been approved by the Operation team and are now awaiting final closure.</p>
            <p>Your action is required to proceed with the final verification for these items.</p>
            <p>The detailed list is attached in the spreadsheet <i>'Julius to check.xlsx'</i> for your convenience.</p>
            <p>Thank you for your attention to this matter.</p>
            <p>Best regards,</p>
            <p><b>Daniel Alves Anversi - Digital Engineering</b></p>
        </body>
        </html>
        """
        mail.Attachments.Add(PATH_JULIUS_CHECK)
        mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail para Julius enviado com sucesso.")

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO CRÍTICO ao enviar e-mail para Julius: {str(e)}\n{erro_detalhado}")


# --- EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    print(f"--- INICIANDO PROCESSO DE AUTOMAÇÃO GERAL ({datetime.now().strftime('%d/%m/%Y %H:%M:%S')}) ---")
    hora_atual = datetime.now().hour

    # --- FLUXO 1: Relatório Principal (Topside) ---
    print("\n--- [FLUXO 1/4] Processando Relatório Principal (Topside) ---")
    dados_topside, log_topside, sucesso_topside = processar_dados()
    if sucesso_topside:
        print("-> Dados Topside processados com sucesso.")

        # Geração do novo gráfico de fechamento
        sucesso_fechamento, log_fechamento = gerar_grafico_fechamento_operacao(dados_topside['df_full'])
        if not sucesso_fechamento:
            print("-> !!! FALHA NA GERAÇÃO DO GRÁFICO DE FECHAMENTO !!!")
            # A falha aqui não impede o envio do e-mail principal, mas o erro será logado.
            log_total_topside = log_topside + log_fechamento
        else:
            log_total_topside = log_topside

        sucesso_dashboard, log_dashboard = gerar_dashboard_imagem(dados_topside)
        log_total_topside += log_dashboard
        if sucesso_dashboard:
            print("-> Dashboard Topside gerado com sucesso.")
        else:
            print("-> !!! FALHA NA GERAÇÃO DO DASHBOARD TOPSIDE !!!")
        enviar_email(dados_topside, log_total_topside)
    else:
        print("\n!!! FALHA CRÍTICA NO PROCESSAMENTO DOS DADOS TOPSIDE !!!")
        enviar_email_de_falha(log_topside)

    # --- FLUXO 2: E-mail para Julius ---
    print("\n--- [FLUXO 2/4] Verificando E-mail para Julius ---")
    if 7 <= hora_atual < 9:
        if sucesso_topside:
            enviar_mensagem_julius(dados_topside)
        else:
            print("-> O processamento de dados do Topside falhou, e-mail para Julius não pôde ser gerado.")
    else:
        print(f"-> Fora do horário agendado (executado às {hora_atual}h). E-mail para Julius não enviado.")

    # --- FLUXO 3: Relatório E-House ---
    print("\n--- [FLUXO 3/4] Processando Relatório E-House ---")
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
    except FileNotFoundError as e:
        print(f"-> Arquivo E-House não encontrado. O relatório para este fluxo não será gerado. Erro: {e}")
    except Exception as e:
        print(f"\n!!! FALHA CRÍTICA NO PROCESSAMENTO DOS DADOS E-HOUSE: {e} !!!")
        enviar_email_de_falha([str(e)])

    # --- FLUXO 4: Relatório Vendors ---
    print("\n--- [FLUXO 4/4] Processando Relatório Vendors ---")
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
    except FileNotFoundError as e:
        print(f"-> Arquivo de Vendors não encontrado. O relatório para este fluxo não será gerado. Erro: {e}")
    except Exception as e:
        print(f"\n!!! FALHA CRÍTICA NO PROCESSAMENTO DOS DADOS DE VENDORS: {e} !!!")
        enviar_email_de_falha([str(e)])

    print(f"\n--- PROCESSO DE AUTOMAÇÃO GERAL FINALIZADO ({datetime.now().strftime('%d/%m/%Y %H:%M:%S')}) ---")
