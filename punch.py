import os
import sys
import time
import shutil
import urllib3
import urllib.parse
import requests
import pandas as pd
from datetime import datetime
import re

# Importação para comunicação com Outlook Local
try:
    import win32com.client as win32
except ImportError:
    print("ERRO: Instale a biblioteca pywin32 executando: pip install pywin32")

# Selenium imports
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURAÇÕES DE REDE CORPORATIVA ---
os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURAÇÕES DO AMBIENTE ---
URL_LOGIN_SEATRIUM = "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/DR90%20EHouse%20Punchlist/AllItems.aspx"
URL_BASE_SHAREPOINT = "https://seatrium.sharepoint.com/sites/P84P85DesignReview"

# --- CONFIGURAÇÃO DAS PASTAS DE DESTINO (LISTA) ---
PASTAS_DESTINO = [
    r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches",
    r"C:\Users\E797\Downloads\Teste mensagem e print"
]

CAMINHO_DRIVER_FIXO = r"C:\Users\E797\PycharmProjects\pythonProject\msedgedriver.exe"
EMAIL_DESTINO = '658b4ef7.petrobras.com.br@br.teams.ms'

# --- CONFIGURAÇÃO DAS LISTAS SHAREPOINT ---
LISTAS_SHAREPOINT = {
    "Topside": {
        "nome_api": "P84/85_TOPSIDE_DR90_Punch_List",
        "arquivo_saida": "Punch_DR90_TS.xlsx",
        "colunas": [
            "DECK No.", "Action Description", "KBR Comment", "Company", "KBR Discipline", "Status",
            "Date Cleared by KBR", "Petrobras Response By", "Petrobras Response Date", "Petrobras Response ",
            "Petrobras Remarks", "Petrobras Discipline", "Petrobras Responsible", "Seatrium Remarks", "Zone",
            "Date Cleared by Petrobras", "S3D Item Tags", "Punch No", "KBR Target Date",
            "Days Since Date Cleared by KBR",
            "Days Since Date Cleared by Seatrium", "Punched by (Group)", "Petrobras Need Operation to close? (Y/N)",
            "Date Cleared by Petrobras Operation", "Petrobras Operation accept closing? (Y/N)", "Is Reopen? (Y/N)",
            "Seatrium Target Date Calculated", "Petrobras Operation Target Date Calculated",
            "Petrobras Target Date Calculated", "Petrobras Target Date", "Petrobras Operation Target Date",
            "Seatrium Target Date"
        ]
    },
    "E-House": {
        "nome_api": "DR90 E-House Punchlist",
        "arquivo_saida": "Punch_DR90_E-House.xlsx",
        "colunas": [
            "Punch No", "Zone", "DECK No.", "Zone-Punch Number", "Action Description", "Punched by", "Punch SnapShot1",
            "Punch SnapShot2", "Closing SnapShot1", "Hotwork", "ABB/CIMC Discipline", "Company", "Close Out Plan Date",
            "Action by", "Status", "Action Comment", "Date Cleared by ABB", "Days Since Date Cleared by ABB",
            "KBR Response", "KBR Response Date", "KBR Response by", "KBR Remarks", "KBR Category", "KBR Discipline",
            "KBR Screenshot", "Date Cleared by KBR", "Days Since Date Cleared By KBR", "Seatrium Discipline",
            "Seatrium Remarks", "Checked By (Seatrium)", "Seatrium Comments", "Date Cleared By Seatrium",
            "Days Since Date Cleared by Seatrium", "Petrobras Response", "Petrobras Response By",
            "Petrobras Screenshot",
            "Petrobras Response Date", "Petrobras Remarks", "Petrobras Discipline", "Petrobras Category",
            "Date Cleared by Petrobras", "Days Since Date Cleared By Petrobras", "Additional Remarks",
            "ARC Reference No(HFE Only)", "Modified", "Modified By", "Item Type", "Path"
        ]
    },
    "Vendors": {
        "nome_api": "Vendor Package Review Punchlist DR90",
        "arquivo_saida": "Punch_DR90_Vendors.xlsx",
        "colunas": [
            "Punch No", "Zone", "DECK No.", "Zone-Punch Number", "Action Description", "S3D Item Tags", "Punched by",
            "Punch Snapshot", "Punch Snapshot 2", "Punch Snapshot 3", "Punch Snapshot 4", "Close-Out Snapshot 1",
            "Close-Out Snapshot 2", "Action Comment", "Vendor Discipline", "Company", "Action by", "Status",
            "Date Cleared by KBR", "Days Since Date Cleared by KBR", "Petrobras Response", "Petrobras Response by",
            "Petrobras Response Date", "Petrobras Screenshot", "Remarks", "Petrobras Discipline", "Petrobras Category",
            "Date Cleared by Petrobras", "Seatrium Remarks", "Seatrium Discipline", "Checked By (Seatrium)",
            "Seatrium Comments", "Date Cleared By Seatrium", "Days Since Date Cleared by Seatrium", "Modified By",
            "Item Type", "Path"
        ]
    }
}


class AutomacaoPunchList:
    def __init__(self):
        self.driver = None
        self.log_sessao = []

    def registrar_log(self, mensagem):
        timestamp = datetime.now().strftime('%H:%M:%S')
        texto = f"[{timestamp}] {mensagem}"
        print(texto)
        self.log_sessao.append(texto)

    def enviar_via_outlook_app(self, sucesso):
        status_geral = "SUCESSO" if sucesso else "FALHA"
        cor_status_geral = "#28a745" if sucesso else "#dc3545"

        log_html_lines = []
        for linha in self.log_sessao:
            classe_css = ""
            linha_sem_ts = linha.split("] ", 1)[-1] if "] " in linha else linha

            if "SUCESSO" in linha_sem_ts:
                classe_css = "success"
            elif "ERRO" in linha_sem_ts or "FALHA" in linha_sem_ts or "Falha" in linha_sem_ts:
                classe_css = "error"
            elif "AVISO" in linha_sem_ts:
                classe_css = "warning"
            elif "---" in linha_sem_ts:
                classe_css = "info"

            log_html_lines.append(f'<div class="log-line {classe_css}">{linha}</div>')

        log_html = "".join(log_html_lines)

        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; color: #333; }}
                .container {{ padding: 20px; border: 1px solid #dee2e6; border-radius: 5px; max-width: 900px; margin: auto; }}
                .header {{ font-size: 24px; font-weight: bold; color: #004085; border-bottom: 2px solid #004085; padding-bottom: 10px; margin-bottom: 20px;}}
                .status {{ font-size: 20px; font-weight: bold; padding: 12px; color: white; background-color: {cor_status_geral}; border-radius: 4px; text-align: center; }}
                .log-container {{ margin-top: 25px; padding: 20px; background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 5px; font-family: 'Courier New', Courier, monospace; font-size: 14px; white-space: pre-wrap; line-height: 1.6; }}
                .log-line {{ margin-bottom: 5px; }}
                .success {{ color: #155724; background-color: #d4edda; border-left: 5px solid #28a745; padding: 5px 10px; }}
                .error {{ color: #721c24; background-color: #f8d7da; border-left: 5px solid #dc3545; padding: 5px 10px; font-weight: bold; }}
                .warning {{ color: #856404; background-color: #fff3cd; border-left: 5px solid #ffc107; padding: 5px 10px; }}
                .info {{ color: #004085; background-color: #cce5ff; border-left: 5px solid #007bff; padding: 5px 10px; font-weight: bold; margin-top: 15px; margin-bottom: 15px;}}
            </style>
        </head>
        <body>
            <div class="container">
                <p class="header">Relatório de Execução Corporativa</p>
                <p><strong>Data de Execução:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
                <div class="status">Status Geral: {status_geral}</div>
                <div class="log-container">
                    {log_html}
                </div>
            </div>
        </body>
        </html>
        """

        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = EMAIL_DESTINO
            mail.Subject = f"Relatório Automático Punch DR90 - {status_geral}"
            mail.HTMLBody = html_body
            mail.Send()
            self.registrar_log("Log enviado via aplicativo Outlook (Formato HTML).")
        except Exception as e:
            self.registrar_log(f"Falha ao enviar e-mail via Outlook Desktop: {e}")

    def tratar_dados(self, raw_data, colunas_desejadas):
        self.registrar_log("Processando e estruturando dados recebidos...")

        if not raw_data:
            self.registrar_log("AVISO: Não há dados para processar. Criando planilha com cabeçalhos.")
            return pd.DataFrame(columns=colunas_desejadas)

        # Normaliza os dados brutos usando json_normalize
        df = pd.json_normalize(raw_data)
        self.registrar_log(f"Dados brutos normalizados. DataFrame inicial com {df.shape[0]} linhas e {df.shape[1]} colunas.")

        # Mapeamento inteligente de colunas
        col_map = {}
        mapped_original_cols = set()

        # Função helper para normalização
        def normalize(text):
            return re.sub(r'[^a-zA-Z0-9]', '', text).lower()

        # Estratégia 1: Mapear campos complexos (ex: 'Autor' para 'Autor.Title')
        for col_desejada in colunas_desejadas:
            normalized_desejada = normalize(col_desejada)
            for original_col in df.columns:
                if original_col in mapped_original_cols:
                    continue

                # Prioriza a correspondência com '.Title' para campos de lookup/pessoa
                if '.Title' in original_col:
                    base_name = original_col.split('.Title')[0]
                    if normalize(base_name) == normalized_desejada:
                        col_map[original_col] = col_desejada
                        mapped_original_cols.add(original_col)
                        break # Próxima coluna desejada

        # Estratégia 2: Mapear campos simples por correspondência direta
        for col_desejada in colunas_desejadas:
            if col_desejada in col_map.values():
                continue # Já foi mapeado

            normalized_desejada = normalize(col_desejada)
            for original_col in df.columns:
                if original_col in mapped_original_cols:
                    continue

                if normalize(original_col) == normalized_desejada:
                    col_map[original_col] = col_desejada
                    mapped_original_cols.add(original_col)
                    break # Próxima coluna desejada

        # Renomeia as colunas do DataFrame de acordo com o mapeamento
        df.rename(columns=col_map, inplace=True)

        # Garante que todas as colunas desejadas existam, adicionando as que estiverem faltando
        for col in colunas_desejadas:
            if col not in df.columns:
                df[col] = ''  # Adiciona a coluna vazia

        # Reordena e filtra o DataFrame para garantir que ele tenha exatamente as colunas desejadas na ordem correta
        df = df[colunas_desejadas]

        self.registrar_log("Limpando e formatando o DataFrame final...")
        for col in df.columns:
            # Converte tudo para string para manipulação segura
            df[col] = df[col].astype(str).fillna('')
            df.loc[df[col].str.contains("error|#error", case=False, na=False), col] = ""

            # Verifica se a coluna parece ser de data
            is_date_col = "date" in col.lower() or "target" in col.lower()
            contains_iso_date = df[col].str.contains(r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z', na=False).any()

            if is_date_col or contains_iso_date:
                dt_series = pd.to_datetime(df[col], errors='coerce', utc=True)
                valid_dates = dt_series.notna()
                if valid_dates.any():
                    df.loc[valid_dates, col] = dt_series[valid_dates].dt.strftime('%d/%m/%Y')

            # Limpa valores nulos restantes
            df[col] = df[col].replace(['NaT', 'nan', 'None', ''], pd.NA)

        self.registrar_log(f"Limpeza concluída. DataFrame final com {df.shape[0]} linhas e {df.shape[1]} colunas.")
        return df

    def iniciar_sessao_navegador(self):
        if not os.path.exists(CAMINHO_DRIVER_FIXO):
            self.registrar_log(f"ERRO CRÍTICO: Driver não encontrado em {CAMINHO_DRIVER_FIXO}")
            return

        edge_options = Options()
        edge_options.add_argument("--ignore-certificate-errors")

        try:
            service = EdgeService(executable_path=CAMINHO_DRIVER_FIXO)
            self.driver = webdriver.Edge(service=service, options=edge_options)
            self.driver.get(URL_LOGIN_SEATRIUM)

            self.registrar_log("Aguardando login na Seatrium...")
            wait = WebDriverWait(self.driver, 120)

            wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "[role='grid'], #O365_MainLink_Me, #O365_HeaderLeftRegion, #spCommandBar"
            )))

            self.registrar_log("Sessão autenticada detectada.")
        except Exception as e:
            self.registrar_log(f"Erro no navegador: {e}")

    def extrair_dados(self):
        self.log_sessao = []
        self.registrar_log("Iniciando ciclo de extração...")
        ciclo_sucesso = True

        try:
            if not self.driver:
                self.registrar_log("Driver não inicializado.")
                return

            cookies = self.driver.get_cookies()
            session = requests.Session()
            session.verify = False
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'])

            for pasta in PASTAS_DESTINO:
                if not os.path.exists(pasta):
                    try:
                        os.makedirs(pasta)
                        self.registrar_log(f"Pasta criada: {pasta}")
                    except Exception as e:
                        self.registrar_log(f"ERRO ao criar pasta {pasta}: {e}")

            for nome_lista, config in LISTAS_SHAREPOINT.items():
                self.registrar_log(f"--- Iniciando processamento da lista: {nome_lista} ---")
                nome_api = config["nome_api"]
                arquivo_saida = config["arquivo_saida"]
                colunas_desejadas = config["colunas"]

                try:
                    safe_nome_api = urllib.parse.quote(nome_api)
                    # Query simplificada para obter todos os itens e campos
                    endpoint = f"{URL_BASE_SHAREPOINT}/_api/web/lists/getbytitle('{safe_nome_api}')/items?$top=5000"
                    headers = {"Accept": "application/json;odata=verbose"}

                    self.registrar_log(f"Baixando dados para '{nome_api}'...")
                    response = session.get(endpoint, headers=headers)

                    if response.status_code == 200:
                        results = response.json().get('d', {}).get('results', [])
                        if results:
                            df_final = self.tratar_dados(results, colunas_desejadas)

                            for pasta_destino in PASTAS_DESTINO:
                                caminho_final = os.path.join(pasta_destino, arquivo_saida)
                                try:
                                    if not os.path.exists(pasta_destino):
                                        os.makedirs(pasta_destino)

                                    # --- Início da Lógica para Salvar com Tabela ---
                                    with pd.ExcelWriter(caminho_final, engine='xlsxwriter') as writer:
                                        df_final.to_excel(writer, sheet_name='Sheet1', index=False)

                                        # Obter os objetos workbook e worksheet do xlsxwriter
                                        worksheet = writer.sheets['Sheet1']

                                        # Obter as dimensões do dataframe
                                        (num_rows, num_cols) = df_final.shape

                                        if num_rows > 0:
                                            # Criar a tabela
                                            worksheet.add_table(0, 0, num_rows, num_cols - 1, {
                                                'name': 'Tabela_query',
                                                'columns': [{'header': col} for col in df_final.columns]
                                            })

                                    self.registrar_log(f"SUCESSO: Planilha '{nome_lista}' salva com tabela em: {caminho_final}")
                                    # --- Fim da Lógica ---

                                except PermissionError:
                                    self.registrar_log(
                                        f"ERRO DE PERMISSÃO: O arquivo '{arquivo_saida}' está aberto. Feche-o e tente novamente.")
                                    ciclo_sucesso = False
                                except Exception as e_save:
                                    self.registrar_log(f"ERRO ao salvar arquivo em {pasta_destino}: {e_save}")
                                    ciclo_sucesso = False
                        else:
                            self.registrar_log(f"AVISO: A lista '{nome_lista}' está vazia.")
                    else:
                        self.registrar_log(f"ERRO: Falha ao baixar dados. Status API: {response.status_code}, {response.text}")
                        ciclo_sucesso = False

                except Exception as e_lista:
                    self.registrar_log(f"ERRO CRÍTICO na lista '{nome_lista}': {str(e_lista)}")
                    ciclo_sucesso = False
                finally:
                    self.registrar_log(f"--- Fim lista: {nome_lista} ---\n")

        except Exception as e_ciclo:
            self.registrar_log(f"Falha crítica no ciclo: {e_ciclo}")
            ciclo_sucesso = False
        finally:
            self.enviar_via_outlook_app(ciclo_sucesso)

    def executar(self):
        self.iniciar_sessao_navegador()
        if self.driver:
            self.extrair_dados()
            while True:
                print("Próximo ciclo em 10 minutos...")
                time.sleep(600)
                self.extrair_dados()


if __name__ == "__main__":
    AutomacaoPunchList().executar()
