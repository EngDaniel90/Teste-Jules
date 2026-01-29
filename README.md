# AtaMaster Pro - Sistema de Atas de Reunião "Ata Viva"

Este programa foi desenvolvido para automatizar a criação de atas de reunião, permitindo o acompanhamento contínuo de tarefas ("Ata Viva"), onde itens abertos de uma reunião são automaticamente transportados para a próxima.

## Arquitetura do Sistema

O sistema utiliza uma stack tecnológica moderna baseada em Python:

- **Interface Gráfica (GUI):** [Flet](https://flet.dev) - Framework baseado em Flutter para criar interfaces modernas.
- **Banco de Dados:** SQLite com SQLAlchemy - Banco de dados local em arquivo único (`atamaster.db`).
- **Geração de PDF:** ReportLab - Biblioteca profissional para criação de documentos PDF.
- **ORM:** SQLAlchemy - Para mapeamento objeto-relacional e gerenciamento de banco de dados.

### Estrutura de Arquivos Necessária

Para que o programa funcione e possa ser compilado, a seguinte estrutura deve ser mantida:

- `atamaster.py`: Arquivo principal contendo toda a lógica do programa e interface.
- `generate_manual.py`: Script auxiliar para gerar o manual do usuário em PDF.
- `BUILD.md`: Instruções detalhadas para gerar o executável (.exe).
- `atamaster.db`: (Gerado automaticamente) Arquivo do banco de dados local.

## Como Executar

1. Certifique-se de ter o Python 3.10 ou superior instalado.
2. Instale as dependências:
   ```bash
   pip install flet sqlalchemy reportlab openpyxl
   ```
3. Execute o programa:
   ```bash
   python atamaster.py
   ```

## Funcionalidades Principais

- **Grupos de Reunião:** Organize suas atas por contexto (ex: "Engenharia", "Diretoria").
- **Ata Viva:** Tarefas com status "OPEN" são importadas automaticamente para novas reuniões do mesmo grupo.
- **Alertas de Prazo:** Sistema visual que destaca tarefas atrasadas, especialmente quando atingem o 3º prazo.
- **Participantes e Empresas:** Cadastro simplificado e importação em massa via Excel (Col A: Nome, Col B: Email).
- **Exportação PDF:** Gera atas profissionais com tabelas, status e espaços para assinatura.
- **Backup/Restore:** Sistema simplificado de segurança dos dados.

---
Desenvolvido por Daniel Alves Anversi
