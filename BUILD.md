# Guia de Criação do Executável (.exe) - AtaMaster Pro

Este guia descreve como transformar o script Python `atamaster.py` em um arquivo executável standalone para Windows.

## Pré-requisitos

Antes de começar, certifique-se de ter o Python instalado e as bibliotecas necessárias. Abra o terminal (Prompt de Comando ou PowerShell) e execute:

```bash
pip install flet reportlab sqlalchemy
```

Além disso, instale o pacote que lidará com a criação do executável:

```bash
pip install pyinstaller
```

## Passo a Passo para Gerar o Executável

O Flet possui um comando simplificado para empacotar aplicações. Siga estes passos:

1. **Abra o Terminal** na pasta onde está o arquivo `atamaster.py`.
2. **Execute o comando abaixo**:

```bash
flet pack atamaster.py --name "AtaMasterPro" --icon "icon.ico"
```

*Nota: Se você não tiver um arquivo `icon.ico`, você pode omitir o parâmetro `--icon` ou usar um arquivo .png que o Flet converterá automaticamente.*

### O que o comando faz:
- Ele utiliza o `PyInstaller` internamente para agrupar o script e todas as dependências (como SQLAlchemy e ReportLab) em um único arquivo.
- O executável final será criado dentro de uma pasta chamada `dist`.

## Dicas Importantes

- **Arquivo Único:** Se você preferir um único arquivo executável (em vez de uma pasta), o comando acima já tenta fazer isso por padrão no Flet.
- **Portabilidade:** O banco de dados `atamaster.db` será criado na mesma pasta onde o executável for rodado pela primeira vez. Para levar seus dados para outro computador, lembre-se de levar o arquivo `.db` junto ou utilizar a função de **Backup/Exportar** dentro do programa.
- **Windows Only:** Para gerar um `.exe` para Windows, você deve executar o comando `flet pack` em um computador com Windows.

## Solução de Problemas

Se o executável fechar imediatamente ao abrir:
1. Tente rodar o comando `flet pack` sem o parâmetro `--noconsole` (se estiver usando PyInstaller puro) para ver se há erros de importação.
2. Certifique-se de que todas as bibliotecas (sqlalchemy, reportlab, flet) estão instaladas no ambiente onde você está compilando.
