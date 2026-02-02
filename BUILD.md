# BUILD - AtaMaster Pro v3

Este documento explica como gerar o executável (.exe) para Windows.

## Dependências
Instale as bibliotecas necessárias:
```bash
pip install flet sqlalchemy reportlab pypdf openpyxl pyinstaller
```

## Geração do Executável
Utilize o comando abaixo:
```bash
flet pack atamaster.py --name "AtaMasterPro" --icon "app_icon.ico"
```

## Arquitetura
- **Persistência**: SQLite (atamaster_pro.db)
- **Interface**: Flet (Material 3)
- **Relatórios**: ReportLab + PyPDF (Fusion Engine)
- **Lógica**: Ata Viva (Task persistence between meetings)

## Desenvolvedor
Daniel Alves Anversi
