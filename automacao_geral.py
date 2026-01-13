import pandas as pd
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import win32com.client as win32
from datetime import datetime, time as datetime_time
import traceback
import matplotlib.pyplot as plt
import seaborn as sns
import schedule

# --- CONFIGURAÇÕES GLOBAIS ---
# Usando os.path.join para criar caminhos de forma segura
DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'Downloads', 'automacao_punch_list')
PATH_PUNCH_TS = os.path.join(DOWNLOAD_DIR, 'Punch_DR90_TS.xlsx')
PATH_PUNCH_HULL = os.path.join(DOWNLOAD_DIR, 'Punch_DR90_HULL.xlsx')
PATH_PUNCH_TEC = os.path.join(DOWNLOAD_DIR, 'Punch_TEC.xlsx')
PATH_RDS = os.path.join(DOWNLOAD_DIR, 'RDs.xlsx')

# Diretórios para salvar saídas
OUTPUT_DIR_TS = os.path.join(DOWNLOAD_DIR, 'output_ts')
OUTPUT_DIR_HULL = os.path.join(DOWNLOAD_DIR, 'output_hull')
OUTPUT_DIR_TEC = os.path.join(DOWNLOAD_DIR, 'output_tec')

# Criar diretórios se não existirem
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR_TS, exist_ok=True)
os.makedirs(OUTPUT_DIR_HULL, exist_ok=True)
os.makedirs(OUTPUT_DIR_TEC, exist_ok=True)

# --- URLs e Credenciais (Mantenha seguro) ---
SHAREPOINT_URLS = {
    "ts": "https://petrobras.sharepoint.com/sites/FPSOAlexandredeGusmo-DR90/Lists/Punch%20List%20DR90%20TS/AllItems.aspx",
    "hull": "https://petrobras.sharepoint.com/sites/FPSOAlexandredeGusmo-DR90/Lists/Punch%20List%20DR90%20Hull/AllItems.aspx",
    "tec": "https://petrobras.sharepoint.com/sites/FPSOAlexandredeGusmo-DR90/Lists/Punch%20List%20TEC/AllItems.aspx",
    "rds": "https://petrobras.sharepoint.com/sites/FPSOAlexandredeGusmo-DR90/Lists/RDs/AllItems.aspx"
}
EMAIL_DESTINO = "658b4ef7.petrobras.com.br@br.teams.ms"
EMAIL_LOG = "seu_email_para_logs@exemplo.com" # Altere para seu e-mail de log

# --- FUNÇÕES DE DOWNLOAD ---
def configurar_driver_chrome(download_path):
    """Configura o WebDriver do Chrome com um diretório de download específico."""
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": download_path}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    # chrome_options.add_argument("--headless") # Descomente para rodar em segundo plano
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def baixar_lista_sharepoint(driver, url, nome_arquivo):
    """Navega até a URL do SharePoint e baixa a lista como XLSX."""
    download_path = os.path.join(DOWNLOAD_DIR, f"{nome_arquivo}.xlsx")
    # Limpa arquivo antigo se existir para garantir um novo download
    if os.path.exists(download_path):
        os.remove(download_path)

    print(f"Acessando URL: {url}")
    driver.get(url)
    wait = WebDriverWait(driver, 180) # Timeout de 3 minutos para login manual

    try:
        print("Aguardando o botão 'Exportar'...")
        export_button = wait.until(EC.element_to_be_clickable((By.ID, "id__228-menu-item")))
        export_button.click()

        print("Aguardando o botão 'Excel'...")
        excel_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='Excel']")))
        excel_button.click()

        print("Aguardando o download do arquivo...")
        # Lógica para esperar o download ser concluído
        timeout = time.time() + 120 # Timeout de 2 minutos para o download
        while not os.path.exists(download_path):
            time.sleep(1)
            if time.time() > timeout:
                raise TimeoutError(f"Download do arquivo {nome_arquivo}.xlsx demorou demais.")
        print(f"Arquivo '{nome_arquivo}.xlsx' baixado com sucesso em {DOWNLOAD_DIR}.")
        return True
    except Exception as e:
        print(f"Erro ao tentar baixar '{nome_arquivo}': {e}")
        print(traceback.format_exc())
        return False

def executar_downloads():
    """Função principal para orquestrar o download de todas as planilhas."""
    driver = configurar_driver_chrome(DOWNLOAD_DIR)
    sucesso_geral = True
    try:
        # O primeiro acesso requer login, os demais usarão a sessão
        print("Faça o login no SharePoint na janela do navegador que foi aberta.")
        if not baixar_lista_sharepoint(driver, SHAREPOINT_URLS["ts"], "Punch_DR90_TS"):
            sucesso_geral = False
        if not baixar_lista_sharepoint(driver, SHAREPOINT_URLS["hull"], "Punch_DR90_HULL"):
            sucesso_geral = False
        if not baixar_lista_sharepoint(driver, SHAREPOINT_URLS["tec"], "Punch_TEC"):
            sucesso_geral = False
        if not baixar_lista_sharepoint(driver, SHAREPOINT_URLS["rds"], "RDs"):
            sucesso_geral = False

    finally:
        driver.quit()
        print("Downloads concluídos. Navegador fechado.")
    return sucesso_geral

# --- FUNÇÕES DE PROCESSAMENTO E ANÁLISE (O CÓDIGO DO 'OFENSOR') ---
def processar_dados(path_punch, path_rds, nome_relatorio):
    """Função principal de processamento de dados de uma planilha de Punch List."""
    log = [f"--- Log de Processamento para: {nome_relatorio} ---"]
    try:
        if not os.path.exists(path_punch):
            raise FileNotFoundError(f"Arquivo de punch list não encontrado: {path_punch}")
        if not os.path.exists(path_rds):
            raise FileNotFoundError(f"Arquivo de RDs não encontrado: {path_rds}")

        df = pd.read_excel(path_punch)
        df.columns = df.columns.str.strip()

        df_rds = pd.read_excel(path_rds)
        df_rds.columns = df_rds.columns.str.strip()

        hoje = datetime.now()
        log.append(f"[{hoje.strftime('%Y-%m-%d %H:%M:%S')}] Planilhas carregadas.")

        # --- Cálculos ---
        status_counts = df['Status'].value_counts().to_dict()
        pending_pb = df[df['Status'].str.strip() == 'Pending PB Reply'].copy()
        disciplina_status = pending_pb['Petrobras Discipline'].value_counts().to_dict()

        mask_op_reply = (df['Petrobras Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
                        (df['Petrobras Operation accept closing? (Y/N)'].isna())
        df_pending_op = df[mask_op_reply].copy()
        count_pending_op_reply = len(df_pending_op)

        df_pending_op['Petrobras Operation Target Date'] = pd.to_datetime(df_pending_op['Petrobras Operation Target Date'], dayfirst=True, errors='coerce')
        df_pending_op['Date Cleared by Petrobras Operation'] = pd.to_datetime(df_pending_op['Date Cleared by Petrobras Operation'], dayfirst=True, errors='coerce')
        mask_op_overdue = (df_pending_op['Petrobras Operation Target Date'] < hoje) & \
                          (df_pending_op['Date Cleared by Petrobras Operation'].isna())
        count_op_overdue = len(df_pending_op[mask_op_overdue])
        df_op_overdue_export = df_pending_op[mask_op_overdue].copy() # Para exportar

        pending_pb['Petrobras Target Date'] = pd.to_datetime(pending_pb['Petrobras Target Date'], dayfirst=True, errors='coerce')
        df_esup_overdue = pending_pb[pending_pb['Petrobras Target Date'] < hoje].copy()
        count_esup_overdue = len(df_esup_overdue)
        df_esup_overdue_export = df_esup_overdue.copy() # Para exportar

        overdue_esup_dep_op = df_esup_overdue[df_esup_overdue.index.isin(df_pending_op.index)]
        count_esup_dep_op = len(overdue_esup_dep_op)
        count_esup_indep_op = count_esup_overdue - count_esup_dep_op

        mask_op_group = df['Petrobras Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])
        resp_op_group = len(df[mask_op_group & df['Date Cleared by Petrobras Operation'].notna()])

        mask_eng_group = df['Petrobras Punched by (Group)'] == 'PB - Engineering'
        resp_eng_by_op = len(df[mask_eng_group & df['Date Cleared by Petrobras Operation'].notna()])

        disciplinas_pendentes = pending_pb['Petrobras Discipline'].unique()
        mencoes_rds = []
        for disc in disciplinas_pendentes:
            if pd.isna(disc): continue
            row = df_rds[df_rds.iloc[:, 0].str.strip() == disc.strip()]
            if not row.empty:
                nomes = row.iloc[0, 1:4].dropna().tolist()
                for nome in nomes:
                    mencoes_rds.append(f"@{nome}")

        resultados = {
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
            "nome_relatorio": nome_relatorio
        }

        log.append("Processamento de dados concluído com sucesso.")
        return resultados, log, True, df_op_overdue_export, df_esup_overdue_export

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento de {nome_relatorio}: {str(e)}\n{erro_detalhado}")
        return None, log, False, None, None

def criar_grafico_disciplinas(disciplina_status, output_path, nome_relatorio):
    """Cria e salva um gráfico de barras para as disciplinas."""
    if not disciplina_status:
        print(f"Não há dados de disciplina para gerar gráfico de {nome_relatorio}.")
        return None

    plt.style.use('seaborn-v0_8-ggrid')
    fig, ax = plt.subplots(figsize=(12, 8))

    disciplinas = list(disciplina_status.keys())
    valores = list(disciplina_status.values())

    sns.barplot(x=valores, y=disciplinas, ax=ax, palette='viridis', orient='h')

    ax.set_title(f'Pendências por Disciplina - {nome_relatorio}', fontsize=16, weight='bold')
    ax.set_xlabel('Quantidade de Itens Pendentes', fontsize=12)
    ax.set_ylabel('Disciplina', fontsize=12)

    for i, v in enumerate(valores):
        ax.text(v + 0.1, i, str(v), color='black', va='center', fontweight='bold')

    plt.tight_layout()

    filepath = os.path.join(output_path, f'grafico_disciplinas_{nome_relatorio}.png')
    plt.savefig(filepath)
    plt.close(fig)
    print(f"Gráfico de disciplinas para {nome_relatorio} salvo em: {filepath}")
    return filepath

def criar_grafico_indicadores(dados, output_path, nome_relatorio):
    """Cria e salva um gráfico de barras para os indicadores chave."""
    indicadores = {
        'Pending Operation Reply': dados['pending_op_reply'],
        'Petrobras Operation Overdue': dados['op_overdue'],
        'Petrobras ESUP Overdue': dados['esup_overdue'],
        'Overdue ESUP (Dep. Operação)': dados['esup_dep_op'],
        'Overdue ESUP (Não Dep. Operação)': dados['esup_indep_op']
    }

    plt.style.use('seaborn-v0_8-ggrid')
    fig, ax = plt.subplots(figsize=(12, 8))

    nomes = list(indicadores.keys())
    valores = list(indicadores.values())

    sns.barplot(x=valores, y=nomes, ax=ax, palette='plasma', orient='h')

    ax.set_title(f'Indicadores Chave de Pendências - {nome_relatorio}', fontsize=16, weight='bold')
    ax.set_xlabel('Quantidade', fontsize=12)
    ax.set_ylabel('')

    for i, v in enumerate(valores):
        ax.text(v + 0.1, i, str(v), color='black', va='center', fontweight='bold')

    plt.tight_layout()

    filepath = os.path.join(output_path, f'grafico_indicadores_{nome_relatorio}.png')
    plt.savefig(filepath)
    plt.close(fig)
    print(f"Gráfico de indicadores para {nome_relatorio} salvo em: {filepath}")
    return filepath

def criar_grafico_status_geral(status_counts, output_path, nome_relatorio):
    """Cria e salva um gráfico de pizza para o status geral."""
    if not status_counts:
        print(f"Não há dados de status para gerar gráfico de {nome_relatorio}.")
        return None

    plt.style.use('seaborn-v0_8-ggrid')
    fig, ax = plt.subplots(figsize=(10, 10))

    labels = status_counts.keys()
    sizes = status_counts.values()

    # Explode a fatia com maior valor para destaque
    explode = tuple([0.05 if size == max(sizes) else 0 for size in sizes])

    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, pctdistance=0.85, explode=explode,
           wedgeprops=dict(width=0.3))

    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig.gca().add_artist(centre_circle)

    ax.set_title(f'Distribuição Geral de Status - {nome_relatorio}', fontsize=16, weight='bold')
    ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

    plt.tight_layout()

    filepath = os.path.join(output_path, f'grafico_status_{nome_relatorio}.png')
    plt.savefig(filepath)
    plt.close(fig)
    print(f"Gráfico de status para {nome_relatorio} salvo em: {filepath}")
    return filepath

def enviar_email(dados, log_processo, anexos=None, email_especial=False):
    """Envia um e-mail formatado via Outlook."""
    if anexos is None:
        anexos = []

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Importance = 2  # Alta importância

        if email_especial: # Email para RDs com pendências Overdue ESUP
            mail.To = EMAIL_DESTINO # Pode ser alterado para destinatários específicos
            mail.Subject = f"ALERTA: Pendências ESUP Atrasadas - {dados['nome_relatorio']} - {datetime.now().strftime('%d/%m/%Y')}"

            mail.HTMLBody = f"""
            <div style="font-family: Calibri, sans-serif; font-size: 11pt;">
                <p style="color: red; font-weight: bold;">[ALERTA DE PENDÊNCIAS]</p>
                <p><b>Atenção RDs:</b><br>{dados['mencoes_rds']}</p>
                <p>Foram identificados <b>{dados['esup_overdue']}</b> itens com status "Pending PB Reply" cuja data alvo (Petrobras Target Date) está vencida.</p>
                <p>Destes, <b>{dados['esup_dep_op']}</b> dependem de uma ação da Operação e <b>{dados['esup_indep_op']}</b> não dependem.</p>
                <p>Solicitamos verificação e ação para os itens sob sua responsabilidade.</p>
                <p>Uma lista detalhada dos itens atrasados está anexa a este e-mail.</p>
                <p><i>Este é um e-mail automático.</i></p>
            </div>
            """
        elif dados['op_overdue'] > 0: # E-mail de alerta para Operação
            mail.To = EMAIL_DESTINO # Pode ser alterado para um grupo da Operação
            mail.Subject = f"ALERTA: Pendências da Operação Atrasadas - {dados['nome_relatorio']} - {datetime.now().strftime('%d/%m/%Y')}"

            mail.HTMLBody = f"""
            <div style="font-family: Calibri, sans-serif; font-size: 11pt;">
                <p style="color: red; font-weight: bold;">[ALERTA DE PENDÊNCIAS DA OPERAÇÃO]</p>
                <p>Prezados,</p>
                <p>Foram identificados <b>{dados['op_overdue']}</b> itens sob responsabilidade da Operação que estão com a data alvo (Petrobras Operation Target Date) vencida e sem data de liberação.</p>
                <p>Solicitamos especial atenção a estes itens para evitar maiores impactos no cronograma.</p>
                <p>Uma lista detalhada dos itens atrasados está anexa a este e-mail.</p>
                <p><i>Este é um e-mail automático.</i></p>
            </div>
            """
        else: # E-mail de relatório geral
            mail.To = EMAIL_DESTINO
            mail.Subject = f"Relatório de Pendências - {dados['nome_relatorio']} - {datetime.now().strftime('%d/%m/%Y')}"

            disciplinas_texto = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_status'].items()]) if dados['disciplina_status'] else "<li>Nenhuma pendência por disciplina.</li>"

            mail.HTMLBody = f"""
            <div style="font-family: Calibri, sans-serif; font-size: 11pt;">
                <p style="color: #0078D4; font-weight: bold;">[RELATÓRIO DIÁRIO DE PENDÊNCIAS]</p>
                <p>@Acompanhamento Design Review {dados['nome_relatorio']}</p>
                <p>Prezados,</p>
                <p>Segue a atualização sobre as pendências do <b>Design Review {dados['nome_relatorio']}</b>:</p>
                <p>Atualmente, temos <b>{dados['status_counts'].get('Pending PB Reply', 0)}</b> itens com status <b>"Pending PB Reply"</b>.</p>

                <p><b>Atenção RDs com pendências:</b><br>{dados['mencoes_rds'] if dados['mencoes_rds'] else "Nenhum RD com pendências."}</p>

                <p>Abaixo seguem os gráficos com o resumo da situação atual. Os arquivos com os detalhes dos itens atrasados (se houver) estão anexados.</p>

                <p><i>Este é um e-mail automático. Em caso de dúvidas, consulte a lista no SharePoint.</i></p>
            </div>
            """

        # Anexar arquivos
        for anexo_path in anexos:
            if anexo_path and os.path.exists(anexo_path):
                mail.Attachments.Add(anexo_path)

        mail.Send()
        print(f"E-mail '{mail.Subject}' enviado com sucesso.")

    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        print(traceback.format_exc())

def enviar_log(log_geral):
    """Envia um e-mail de log com o resultado da execução."""
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_LOG
        mail.Subject = f"Log de Execução da Automação Punch List - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        mail.Body = f"Execução em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n" + "\n".join(log_geral)
        mail.Send()
        print("E-mail de log enviado.")
    except Exception as e:
        print(f"Falha ao enviar e-mail de log: {e}")

def tarefa_principal():
    """Função que encapsula toda a lógica de execução da automação."""
    print(f"\n--- INICIANDO EXECUÇÃO DA AUTOMAÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
    log_geral_execucao = [f"INÍCIO DA EXECUÇÃO: {datetime.now()}"]

    # 1. Download dos arquivos
    log_geral_execucao.append("\n--- FASE 1: DOWNLOAD DOS ARQUIVOS ---")
    sucesso_download = executar_downloads()
    if not sucesso_download:
        log_geral_execucao.append("ERRO: Falha no download de um ou mais arquivos. A automação será interrompida.")
        print("!!! FALHA CRÍTICA NOS DOWNLOADS. ABORTANDO EXECUÇÃO. !!!")
        enviar_log(log_geral_execucao)
        return # Interrompe a execução se os downloads falharem
    log_geral_execucao.append("SUCESSO: Todos os arquivos foram baixados.")

    # 2. Processamento e Envio para cada tipo de Punch List
    punch_lists_config = [
        {"nome": "TS", "path": PATH_PUNCH_TS, "output_dir": OUTPUT_DIR_TS},
        {"nome": "HULL", "path": PATH_PUNCH_HULL, "output_dir": OUTPUT_DIR_HULL},
        {"nome": "TEC", "path": PATH_PUNCH_TEC, "output_dir": OUTPUT_DIR_TEC},
    ]

    for config in punch_lists_config:
        nome_relatorio = config["nome"]
        path_punch = config["path"]
        output_dir = config["output_dir"]

        print(f"\n--- FASE 2: PROCESSANDO RELATÓRIO '{nome_relatorio}' ---")
        log_geral_execucao.append(f"\n--- PROCESSANDO: {nome_relatorio} ---")

        dados, log_proc, sucesso, df_op_overdue, df_esup_overdue = processar_dados(path_punch, PATH_RDS, nome_relatorio)
        log_geral_execucao.extend(log_proc)

        if sucesso:
            print(f"Dados de '{nome_relatorio}' processados com sucesso. Gerando artefatos...")
            anexos_email_geral = []

            # Gerar gráficos
            grafico_disciplinas_path = criar_grafico_disciplinas(dados['disciplina_status'], output_dir, nome_relatorio)
            grafico_indicadores_path = criar_grafico_indicadores(dados, output_dir, nome_relatorio)
            grafico_status_path = criar_grafico_status_geral(dados['status_counts'], output_dir, nome_relatorio)

            if grafico_disciplinas_path: anexos_email_geral.append(grafico_disciplinas_path)
            if grafico_indicadores_path: anexos_email_geral.append(grafico_indicadores_path)
            if grafico_status_path: anexos_email_geral.append(grafico_status_path)

            # Exportar itens atrasados para Excel
            path_op_overdue_excel = None
            if df_op_overdue is not None and not df_op_overdue.empty:
                path_op_overdue_excel = os.path.join(output_dir, f'operacao_overdue_{nome_relatorio}.xlsx')
                df_op_overdue.to_excel(path_op_overdue_excel, index=False)
                print(f"Arquivo de operação overdue de '{nome_relatorio}' salvo.")

            path_esup_overdue_excel = None
            if df_esup_overdue is not None and not df_esup_overdue.empty:
                path_esup_overdue_excel = os.path.join(output_dir, f'esup_overdue_{nome_relatorio}.xlsx')
                df_esup_overdue.to_excel(path_esup_overdue_excel, index=False)
                print(f"Arquivo de ESUP overdue de '{nome_relatorio}' salvo.")

            # Enviar e-mails condicionais
            # 1. E-mail geral
            print(f"Enviando e-mail de relatório geral para '{nome_relatorio}'...")
            enviar_email(dados, log_proc, anexos=anexos_email_geral)

            # 2. E-mail para Operação se houver atrasos
            if dados['op_overdue'] > 0 and path_op_overdue_excel:
                print(f"Enviando e-mail de ALERTA de operação para '{nome_relatorio}'...")
                enviar_email(dados, log_proc, anexos=[path_op_overdue_excel], email_especial=False) # A lógica dentro de enviar_email decide o corpo

            # 3. E-mail para RDs se houver ESUP overdue
            if dados['esup_overdue'] > 0 and path_esup_overdue_excel:
                print(f"Enviando e-mail de ALERTA ESUP para '{nome_relatorio}'...")
                enviar_email(dados, log_proc, anexos=[path_esup_overdue_excel], email_especial=True)

            log_geral_execucao.append("SUCESSO: E-mails enviados.")
        else:
            print(f"\n!!! FALHA NO PROCESSAMENTO DE '{nome_relatorio}' !!!")
            log_geral_execucao.append(f"FALHA: O processamento de '{nome_relatorio}' falhou. Verifique os logs detalhados.")

    log_geral_execucao.append(f"\nFIM DA EXECUÇÃO: {datetime.now()}")
    enviar_log(log_geral_execucao)
    print(f"\n--- EXECUÇÃO DA AUTOMAÇÃO FINALIZADA: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")


# --- AGENDAMENTO ---
if __name__ == "__main__":
    print("======================================================")
    print("     AUTOMATIZAÇÃO DE RELATÓRIO DE PUNCH LISTS      ")
    print("======================================================")
    print("A automação está configurada para rodar nos seguintes horários:")
    print(" - 08:00")
    print(" - 12:00")
    print(" - 16:30")
    print("\nO script está em execução e aguardando o próximo horário agendado.")
    print("Para interromper, pressione Ctrl+C.")
    print("------------------------------------------------------")

    # Agendamento das tarefas
    schedule.every().day.at("08:00").do(tarefa_principal)
    schedule.every().day.at("12:00").do(tarefa_principal)
    schedule.every().day.at("16:30").do(tarefa_principal)

    # Loop para manter o script rodando e checando os agendamentos
    while True:
        schedule.run_pending()
        time.sleep(1)
