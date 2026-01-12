import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os
import io

# --- CONFIGURAÇÕES DE CREDENCIAIS E SHAREPOINT ---
# As credenciais devem ser configuradas como variáveis de ambiente
SHAREPOINT_USER = os.getenv("SHAREPOINT_USER")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
SHAREPOINT_SITE_URL = "https://seatrium.sharepoint.com/sites/P84P85DesignReview"

# --- CONFIGURAÇÕES DAS PLANILHAS ---
PUNCH_LISTS = {
    "Topside": {
        "url": "/sites/P84P85DesignReview/Lists/DR90%20Topside%20Punchlist/AllItems.aspx",
        "file_name": "Punch_DR90_TS.xlsx",
        "save_path": r"C:\Users\E797\Downloads\Teste mensagem e print",
        "columns_to_keep": None  # Manter todas as colunas
    },
    "E-House": {
        "url": "/sites/P84P85DesignReview/Lists/DR90%20EHouse%20Punchlist/AllItems.aspx?e=bTNUys&CID=98bfe3a1%2D508a%2D6000%2D2c51%2D857abb203708&cidOR=SPO&ovuser=5b6f6241%2D9a57%2D4be4%2D8e50%2D1dfa72e79a57%2Cdaniel%2Eanversi%40petrobras%2Ecom%2Ebr&OR=Teams%2DHL&CT=1766373503454&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNTExMzAwMTMxMiIsIkhhc0ZlZGVyYXRlZFVzZXIiOnRydWV9",
        "file_name": "Punch_DR90_E-House.xlsx",
        "save_path": r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches",
        "columns_to_keep": [
            'Punch No', 'Zone', 'DECK No.', 'Zone-Punch Number', 'Action Description',
            'Punched by', 'Punch SnapShot1', 'Punch SnapShot2', 'Closing SnapShot1', 'Hotwork',
            'ABB/CIMC Discipline', 'Company', 'Close Out Plan Date', 'Action by', 'Status',
            'Action Comment', 'Date Cleared by ABB', 'Days Since Date Cleared by ABB', 'KBR Response',
            'KBR Response Date', 'KBR Response by', 'KBR Remarks', 'KBR Category', 'KBR Discipline',
            'KBR Screenshot', 'Date Cleared by KBR', 'Days Since Date Cleared By KBR', 'Seatrium Discipline',
            'Seatrium Remarks', 'Checked By (Seatrium)', 'Seatrium Comments', 'Date Cleared By Seatrium',
            'Days Since Date Cleared by Seatrium', 'Petrobras Response', 'Petrobras Response By',
            'Petrobras Screenshot', 'Petrobras Response Date', 'Petrobras Remarks', 'Petrobras Discipline',
            'Petrobras Category', 'Date Cleared by Petrobras', 'Days Since Date Cleared By Petrobras',
            'Additional Remarks', 'ARC Reference No(HFE Only)', 'Modified', 'Modified By', 'Item Type', 'Path'
        ]
    },
    "Vendors": {
        "url": "/sites/P84P85DesignReview/Lists/Vendor%20Package%20Review%20Punchlist%20DR90/AllItems.aspx?e=4tHLty&CID=43904b9e%2D7cb2%2D481c%2Db136%2D5285ae014bd9&ovuser=5b6f6241%2D9a57%2D4be4%2D8e50%2D1dfa72e79a57%2Cleojunqueira%40petrobras%2Ecom%2Ebr",
        "file_name": "Punch_DR90_Vendors.xlsx",
        "save_path": r"C:\Users\E797\PETROBRAS\SRGE SI-II SCP85 ES - Planilha_BI_Punches",
        "columns_to_keep": [
            'Punch No', 'Zone', 'DECK No.', 'Zone-Punch Number', 'Action Description', 'S3D Item Tags',
            'Punched by', 'Punch Snapshot', 'Punch Snapshot 2', 'Punch Snapshot 3', 'Punch Snapshot 4',
            'Close-Out Snapshot 1', 'Close-Out Snapshot 2', 'Action Comment', 'Vendor Discipline',
            'Company', 'Action by', 'Status', 'Date Cleared by KBR', 'Days Since Date Cleared by KBR',
            'Petrobras Response', 'Petrobras Response by', 'Petrobras Response Date', 'Petrobras Screenshot',
            'Remarks', 'Petrobras Discipline', 'Petrobras Category', 'Date Cleared by Petrobras',
            'Seatrium Remarks', 'Seatrium Discipline', 'Checked By (Seatrium)', 'Seatrium Comments',
            'Date Cleared By Seatrium', 'Days Since Date Cleared by Seatrium', 'Modified By', 'Item Type', 'Path'
        ]
    }
}

def format_as_table(writer, df, sheet_name):
    """Formata os dados em uma tabela do Excel."""
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    (max_row, max_col) = df.shape
    column_settings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

def download_and_process_list(ctx, list_config):
    """Baixa uma lista do SharePoint, processa e salva como Excel."""
    list_name = list_config["url"].split("/")[-2]
    list_obj = ctx.web.lists.get_by_title(list_name)

    items = list_obj.get_items().execute_query()

    data = []
    for item in items:
        data.append(item.properties)

    if not data:
        print(f"A lista '{list_name}' está vazia ou não foi acessada corretamente.")
        return

    df = pd.DataFrame(data)

    # Garante que os nomes das colunas correspondem ao SharePoint
    df.columns = [col.replace('_x0020_', ' ') for col in df.columns]

    if list_config["columns_to_keep"]:
        # Filtra apenas as colunas desejadas, mantendo a ordem
        df = df[[col for col in list_config["columns_to_keep"] if col in df.columns]]

    file_path = os.path.join(list_config["save_path"], list_config["file_name"])

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        format_as_table(writer, df, "Data")

    print(f"Planilha '{list_config['file_name']}' salva em '{file_path}' com sucesso.")

if __name__ == "__main__":
    if not SHAREPOINT_USER or not SHAREPOINT_PASSWORD:
        print("Erro: As variáveis de ambiente SHAREPOINT_USER e SHAREPOINT_PASSWORD não estão configuradas.")
    else:
        try:
            ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(UserCredential(SHAREPOINT_USER, SHAREPOINT_PASSWORD))

            for list_name, config in PUNCH_LISTS.items():
                print(f"--- Baixando a lista: {list_name} ---")
                download_and_process_list(ctx, config)

        except Exception as e:
            print(f"Ocorreu um erro durante a execução: {e}")
