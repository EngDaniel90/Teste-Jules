# Guia para Criação do Executável (.exe) - AtaMaster Pro

Para transformar o script `atamaster.py` em um arquivo executável para Windows, siga os passos abaixo:

## 1. Instalar o PyInstaller e Flet

Abra o terminal ou prompt de comando e instale as ferramentas necessárias:

```bash
pip install pyinstaller flet sqlalchemy reportlab openpyxl
```

## 2. Gerar o Executável

O Flet possui uma ferramenta integrada que simplifica o uso do PyInstaller. Execute o comando abaixo na pasta do projeto:

```bash
flet pack atamaster.py --name "AtaMasterPro" --icon "icon.ico"
```

*Nota: Se você não tiver um arquivo `icon.ico`, remova o parâmetro `--icon "icon.ico"`.*

## 3. Localizar o Programa

Após a conclusão, uma pasta chamada `dist` será criada. Dentro dela, você encontrará o arquivo `AtaMasterPro.exe`.

## 4. Distribuição

Para que o programa funcione em outros computadores:
- Você só precisa enviar o arquivo `.exe` da pasta `dist`.
- O banco de dados (`atamaster.db`) será criado automaticamente na primeira execução se não existir.
- Não é necessário que a outra pessoa tenha Python instalado.

## Dicas de Arquitetura

- **Portabilidade:** O programa salva os dados no mesmo local onde o executável está rodando.
- **Auto-Contido:** Todas as dependências (SQLAlchemy, ReportLab) são embutidas no executável pelo comando `flet pack`.
- **Compatibilidade:** O código foi ajustado para funcionar com a versão 0.80.4 do Flet, garantindo estabilidade visual.
