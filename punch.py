import os
import sys
import time
import shutil
import urllib3
import urllib.request
import requests
import pandas as pd
from datetime import datetime

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
URL_LOGIN_SEATRIUM = "https://seatrium.sharepoint.com/sites/P84P85DesignReview/Lists/DR90%20EHouse%20Punchlist/AllItems.aspx?e=bTNUys&CID=98bfe3a1%2D508a%2D6000%2D2c51%2D857abb203708&cidOR=SPO&ovuser=5b6f6241%2D57be4%2D8e50%2D1dfa72e79a57%2Cdaniel%2Eanversi%40petrobras%2Ecom%2Ebr&OR=Teams%2DHL&CT=1766373503454&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNTExMzAwMTMxMiIsIkhhc0ZlZGVyYXRlZFVzZXIiOnRydWV9"
URL_BASE_SHAREPOINT = "https://seatrium.sharepoint.com/sites/P84P85DesignReview"

PASTA_DESTINO = r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches"
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
        self.mapeamento_colunas = {}

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
            # Remove o timestamp para a verificação de palavras-chave
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

    def tratar_dados(self, df, colunas_desejadas):
        self.registrar_log("Limpando e formatando dados...")

        df = df.rename(columns=self.mapeamento_colunas)
        colunas_existentes = [c for c in colunas_desejadas if c in df.columns]
        df = df[colunas_existentes].copy()

        for col in df.columns:
            df.loc[:, col] = df[col].astype(str)
            df.loc[df[col].str.contains("error", case=False, na=False), col] = ""

            if "Date" in col or df[col].str.contains(r'\d{4}-\d{2}-\d{2}T', na=False).any():
                try:
                    # Especifica o formato para evitar UserWarning e garantir a conversão correta
                    df_dt = pd.to_datetime(df[col], format='%Y-%m-%dT%H:%M:%SZ', errors='coerce', utc=True)
                    mask = df_dt.notna()
                    df.loc[mask, col] = df_dt[mask].dt.strftime('%d/%m/%Y')
                    df.loc[:, col] = df[col].replace(['NaT', 'nan', 'None', 'nan/nan/nan'], "")
                except Exception:  # Captura exceções caso o formato não seja consistente
                    continue

        return df

    def obter_mapeamento_colunas(self, session, base_url, nome_api):
        self.registrar_log(f"Obtendo mapeamento de colunas para a lista '{nome_api}'...")
        endpoint = f"{base_url}/_api/web/lists/getbytitle('{nome_api}')/fields"
        try:
            headers = {"Accept": "application/json;odata=verbose"}
            response = session.get(endpoint, headers=headers)
            if response.status_code == 200:
                fields = response.json()['d']['results']
                self.mapeamento_colunas = {f['InternalName']: f['Title'] for f in fields}
                self.registrar_log("Dicionário de colunas sincronizado.")
                return True
            else:
                self.registrar_log(
                    f"Falha ao mapear schema do SharePoint para '{nome_api}'. Status: {response.status_code}")
                return False
        except Exception as e:
            self.registrar_log(f"Erro no mapeamento para '{nome_api}': {e}")
            return False

    def iniciar_sessao_navegador(self):
        if not os.path.exists(CAMINHO_DRIVER_FIXO):
            self.registrar_log("Driver não encontrado!")
            return

        edge_options = Options()
        edge_options.add_argument("--ignore-certificate-errors")

        try:
            service = EdgeService(executable_path=CAMINHO_DRIVER_FIXO)
            self.driver = webdriver.Edge(service=service, options=edge_options)
            self.driver.get(URL_LOGIN_SEATRIUM)

            self.registrar_log("Aguardando login na Seatrium...")
            wait = WebDriverWait(self.driver, 120)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[role='grid']")))
            self.registrar_log("Sessão autenticada detectada.")
        except Exception as e:
            self.registrar_log(f"Erro no navegador: {e}")

    def extrair_dados(self):
        self.log_sessao = []
        self.registrar_log("Iniciando ciclo de extração...")
        ciclo_sucesso = True

        try:
            cookies = self.driver.get_cookies()
            session = requests.Session()
            session.verify = False
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'])

            if not os.path.exists(PASTA_DESTINO):
                os.makedirs(PASTA_DESTINO)
                self.registrar_log(f"Pasta de destino criada em: {PASTA_DESTINO}")

            for nome_lista, config in LISTAS_SHAREPOINT.items():
                self.registrar_log(f"--- Iniciando processamento da lista: {nome_lista} ---")
                nome_api = config["nome_api"]
                arquivo_saida = config["arquivo_saida"]
                colunas_desejadas = config["colunas"]

                try:
                    if not self.obter_mapeamento_colunas(session, URL_BASE_SHAREPOINT, nome_api):
                        ciclo_sucesso = False
                        continue

                    self.registrar_log(f"Baixando dados da lista '{nome_api}'...")
                    endpoint = f"{URL_BASE_SHAREPOINT}/_api/web/lists/getbytitle('{nome_api}')/items?$top=5000"
                    headers = {"Accept": "application/json;odata=verbose"}
                    response = session.get(endpoint, headers=headers)

                    if response.status_code == 200:
                        results = response.json().get('d', {}).get('results', [])
                        if results:
                            df_raw = pd.json_normalize(results)
                            df_final = self.tratar_dados(df_raw, colunas_desejadas)

                            caminho_final = os.path.join(PASTA_DESTINO, arquivo_saida)
                            try:
                                df_final.to_excel(caminho_final, index=False)
                                self.registrar_log(f"SUCESSO: Planilha '{nome_lista}' salva em: {caminho_final}")
                            except PermissionError:
                                self.registrar_log(
                                    f"ERRO: O arquivo '{arquivo_saida}' está aberto. Feche-o para salvar.")
                                ciclo_sucesso = False
                            except Exception as e_save:
                                self.registrar_log(f"ERRO ao salvar o arquivo '{arquivo_saida}': {e_save}")
                                ciclo_sucesso = False
                        else:
                            self.registrar_log(f"AVISO: A lista '{nome_lista}' retornou vazia.")
                    else:
                        self.registrar_log(
                            f"ERRO: Falha ao baixar dados da lista '{nome_lista}'. Status API: {response.status_code}")
                        ciclo_sucesso = False

                except Exception as e_lista:
                    self.registrar_log(f"ERRO: Falha inesperada no processamento da lista '{nome_lista}': {e_lista}")
                    ciclo_sucesso = False
                finally:
                    self.registrar_log(f"--- Finalizado processamento da lista: {nome_lista} ---\n")

        except Exception as e_ciclo:
            self.registrar_log(f"Falha crítica no ciclo de extração: {e_ciclo}")
            ciclo_sucesso = False
        finally:
            self.enviar_via_outlook_app(ciclo_sucesso)

    def executar(self):
        self.iniciar_sessao_navegador()
        if self.driver:
            # Executa imediatamente na primeira vez
            self.extrair_dados()

            while True:
                print("Próximo ciclo em 10 minutos...")
                time.sleep(600)
                self.extrair_dados()


if __name__ == "__main__":
    AutomacaoPunchList().executar()
