# Guia de Compilação - AtaMaster Pro

Este documento descreve como configurar o ambiente e gerar o executável (.exe) para Windows.

## 1. Pré-requisitos
- Python 3.10 ou superior
- Pip (gerenciador de pacotes)

## 2. Instalação de Dependências
Execute o comando abaixo para instalar todas as bibliotecas necessárias:

```bash
pip install flet==0.80.4 sqlalchemy aiosqlite reportlab openpyxl pypdf pyinstaller
```

## 3. Estrutura do Projeto
- `atamaster.py`: Código fonte principal (UI e Lógica).
- `atamaster.db`: Banco de dados SQLite (gerado automaticamente na primeira execução).

## 4. Geração do Executável (.exe)
Para criar um arquivo único para Windows, utilize o PyInstaller via Flet:

```bash
flet pack atamaster.py --name "AtaMasterPro" --icon "icon.ico"
```

*Nota: Se você não tiver um ícone, pode omitir a flag `--icon`.*

## 5. Execução
Após a compilação, o executável estará disponível na pasta `dist/`.

---
Desenvolvido por Daniel Alves Anversi
