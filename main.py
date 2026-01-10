import pandas as pd
import os
import time
import traceback
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import win32com.client as win32

# --- CONFIGURAÇÕES DE CAMINHOS E URLs ---
PATH_PUNCH = r'C:\Users\E797\Downloads\Teste mensagem e print\Punch_DR90_TS.xlsx'
PATH_RDS = r'C:\Users\E797\Downloads\Teste mensagem e print\RDs.xlsx'
PATH_SCREENSHOT = r'C:\Users\E797\Downloads\Teste mensagem e print\screenshot_pbi.png'
URL_PBI = "https://app.powerbi.com/groups/me/apps/19dbdcff-e619-4505-a1be-b99673c7d0c0/reports/e668a8c0-fc8a-48d2-96e9-4fcc014221f4/1fe6836411ece2219e39?ctid=5b6f6241-9a57-4be4-8e50-1dfa72e79a57&experience=power-bi"
EMAIL_DESTINO = "658b4ef7.petrobras.com.br@br.teams.ms"

# --- CONFIGURAÇÕES DE LOGIN (IMPORTANTE) ---
# Para evitar riscos de segurança, as credenciais do Power BI não devem ser escritas
# diretamente neste código. Elas devem ser configuradas como variáveis de ambiente no
# sistema onde o script será executado.
#
# Como configurar as variáveis de ambiente:
# 1. Abra o menu Iniciar e pesquise por "Editar as variáveis de ambiente do sistema".
# 2. Na janela de "Propriedades do Sistema", clique em "Variáveis de Ambiente...".
# 3. Em "Variáveis de usuário", clique em "Novo..." e crie duas variáveis:
#    - Nome da variável: POWERBI_USER
#      Valor da variável: seu_email@petrobras.com.br
#    - Nome da variável: POWERBI_PASSWORD
#      Valor da variável: sua_senha
# 4. Reinicie o PyCharm ou o terminal para que as novas variáveis sejam carregadas.

POWERBI_USER = os.getenv("POWERBI_USER")
POWERBI_PASSWORD = os.getenv("POWERBI_PASSWORD")


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
        mask_op_reply = (df['Petrobras Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])) & \
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
        mask_op_group = df['Petrobras Punched by (Group)'].isin(['PB - Operation', 'SEA/KBR'])
        resp_op_group = len(df[mask_op_group & df['Date Cleared by Petrobras Operation'].notna()])

        mask_eng_group = df['Petrobras Punched by (Group)'] == 'PB - Engineering'
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
            "mencoes_rds": " ".join(sorted(list(set(mencoes_rds))))
        }

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Processamento de dados concluído.")
        return resultados, log, True

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO no processamento de dados: {str(e)}\n{erro_detalhado}")
        return None, log, False


def capturar_power_bi_web():
    """
    Abre o Power BI no navegador em modo headless, faz o login e tira um screenshot.
    """
    log = []
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")  # Resolução da captura
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")

        # O webdriver-manager baixa e gerencia o chromedriver automaticamente
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Acessando a URL do Power BI...")
        driver.get(URL_PBI)

        # --- LÓGICA DE LOGIN ---
        # AVISO: Esta automação pode falhar se a Autenticação de Múltiplos Fatores (MFA)
        # estiver habilitada na conta, pois exigirá uma interação manual não prevista no script.

        wait = WebDriverWait(driver, 45) # Tempo de espera aumentado

        # Insere o e-mail (usando um seletor mais estável)
        email_field = wait.until(EC.presence_of_element_located((By.NAME, "loginfmt")))
        email_field.send_keys(POWERBI_USER)
        driver.find_element(By.ID, "idSIButton9").click()
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail inserido.")

        # Insere a senha (usando um seletor mais estável)
        password_field = wait.until(EC.presence_of_element_located((By.NAME, "passwd")))
        password_field.send_keys(POWERBI_PASSWORD)
        driver.find_element(By.ID, "idSIButton9").click()
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Senha inserida.")

        # Lida com a tela "Permanecer conectado"
        try:
            stay_signed_in_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            stay_signed_in_button.click()
            log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Clicado em 'Permanecer conectado'.")
        except:
            log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Tela 'Permanecer conectado' não apareceu, seguindo.")
            pass

        # Aguarda o carregamento do relatório de forma robusta
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Aguardando o relatório carregar...")
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "report-canvas")))
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "visual-container-component")))

        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Relatório carregado, aguardando renderização final...")
        time.sleep(5) # Pausa curta para garantir que os elementos visuais terminem de renderizar

        # Captura de tela
        driver.save_screenshot(PATH_SCREENSHOT)
        log.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Screenshot capturado com sucesso.")

        driver.quit()
        return True, log

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        log.append(f"ERRO CRÍTICO na captura do Power BI: {str(e)}\n{erro_detalhado}")
        if 'driver' in locals():
            driver.quit()
        return False, log


def enviar_email(dados, log_processo):
    """
    Cria e envia um e-mail formatado com os dados do relatório e o log de execução.
    """
    try:
        outlook = win32.Dispatch('outlook.application')

        # --- E-mail Principal ---
        mail = outlook.CreateItem(0)
        mail.Importance = 2  # Marca como Importante
        mail.To = EMAIL_DESTINO
        mail.Subject = f"Status Report: Punch List DR90 TS - {datetime.now().strftime('%d/%m/%Y')}"

        disciplinas_html = "".join([f"<li><b>{k}:</b> {v}</li>" for k, v in dados['disciplina_status'].items()])

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
            <p class="highlight">[MENSAGEM AUTOMÁTICA IMPORTANTE]</p>
            <p class="mention">@Acompanhamento Design Review TS</p>
            <p>Prezados,</p>
            <p>Segue a atualização diária das pendências do <b>Design Review TS</b>:</p>

            <p>Atualmente, temos <span class="highlight">{dados['status_counts'].get('Pending PB Reply', 0)}</span> itens com status <b>Pending PB Reply</b>.</p>

            <p><b>Detalhamento por Disciplina:</b></p>
            <ul>{disciplinas_html}</ul>

            <p><b>Atenção RDs:</b><br><span class="mention">{dados['mencoes_rds']}</span></p>

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

            <p><i>O dashboard atualizado está anexado a este e-mail.</i></p>
            <p>Atenciosamente,</p>
            <p><b>Automação de Relatórios DR90 TS</b></p>
        </body>
        </html>
        """

        if os.path.exists(PATH_SCREENSHOT):
            mail.Attachments.Add(PATH_SCREENSHOT)

        mail.Send()
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] E-mail principal enviado para {EMAIL_DESTINO}.")

        # --- E-mail de Log de Sucesso ---
        log_mail = outlook.CreateItem(0)
        log_mail.To = EMAIL_DESTINO
        log_mail.Subject = f"Log de Execução (Sucesso) - Automação Punch List - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        log_mail.Body = f"Execução concluída com sucesso em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n" + "\n".join(log_processo)
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


# --- EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    print("--- INICIANDO PROCESSO DE AUTOMAÇÃO ---")

    # Etapa 1: Processamento de Dados
    dados_finais, log_proc, sucesso_proc = processar_dados()

    if sucesso_proc:
        print("Dados da planilha processados com sucesso.")

        # Etapa 2: Captura de Tela do Power BI
        sucesso_screenshot, log_screenshot = capturar_power_bi_web()
        log_total = log_proc + log_screenshot

        if sucesso_screenshot:
            print("Captura de tela do Power BI realizada com sucesso.")
        else:
            print("!!! FALHA NA CAPTURA DE TELA DO POWER BI !!!")

        # Etapa 3: Envio de E-mails
        print("Enviando e-mails...")
        enviar_email(dados_finais, log_total)

    else:
        print("\n!!! FALHA CRÍTICA NO PROCESSAMENTO DOS DADOS !!!")
        enviar_email_de_falha(log_proc)

    print("--- PROCESSO FINALIZADO ---")
