# Automação NFS-e

Sistema desktop para organização automática de Notas Fiscais de Serviço Eletrônicas (NFS-e) baixadas do portal nacional, classificando e movendo os arquivos (XML + PDF) diretamente para a estrutura de pastas do servidor fiscal da empresa.

---

## Funcionalidades

- **Classificação automática de NFS-e tomadas** em Retidas, Não Retidas e Canceladas, com base na leitura dos campos de retenção do XML (ISSQN, PIS/COFINS, IRRF, CSLL, CP)
- **Organização de NFS-e emitidas** na estrutura correta do servidor
- **Suporte a ZIP e pastas** — processa empresas extraindo ZIPs automaticamente
- **Mapeamento de nomes** via planilha Excel: correlaciona o nome do portal NFSe com o nome real da pasta no servidor
- **Interface gráfica moderna** (dark mode) com barra de progresso e log de execução em tempo real
- **Configuração persistente** — salva os caminhos informados para a próxima execução

---

## Estrutura de destino gerada

```
Servidor Z:/
└── <Nome da Empresa>/
    └── Fiscal/
        └── Fechamento Fiscal/
            └── <Ano>/
                └── <Mês>/
                    ├── Tomadas/
                    │   ├── Retidas/
                    │   ├── Não Retidas/
                    │   └── Canceladas/
                    └── Emitidas/
```

---

## Requisitos

- Python 3.10+
- Dependências listadas em `requirements.txt`

```
customtkinter>=5.2.0
openpyxl>=3.1.0
```

Instale com:

```bash
pip install -r requirements.txt
```

---

## Como usar

### 1. Gerar o arquivo de configuração (primeira vez)

Execute `criar_config.py` para gerar o arquivo `config.xlsx` com a estrutura de mapeamento de empresas:

```bash
python criar_config.py
```

Preencha a planilha gerada com duas colunas:

| Nome no Portal NFSe | Nome na Pasta do Servidor |
|---|---|
| EMPRESA EXEMPLO LTDA | Empresa Exemplo |
| EMPRESA 123 LTDA | EMPRESA123 E AMIGOS |

### 2. Executar a interface

```bash
python interface.py
```

### 3. Configurar e executar

Na interface, informe:

- **Pasta de Origem** — pasta onde estão as NFS-e baixadas do portal (pastas por empresa ou ZIPs)
- **Pasta de Destino** — pasta raiz do servidor fiscal (ex: `Z:\`)
- **Mapeamento de Empresas** — arquivo `config.xlsx` com a correlação de nomes

Clique em **Executar Automação**. O log exibe em tempo real o que está sendo processado.

---

## Regras de classificação

A nota tomada é classificada como **Retida** se qualquer um dos campos abaixo estiver presente no XML:

| Campo XML | Condição |
|---|---|
| `tpRetISSQN` | `= 1` |
| `tpRetPisCofins` | `= 1` |
| `vRetIRRF` | `> 0` |
| `vRetCP` | `> 0` |
| `vRetCSLL` | `> 0` |

Caso contrário, é classificada como **Não Retida**. Notas em subpastas com "Cancelad" no nome vão para **Canceladas**.

---

## Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| `processador.py` | Lógica de leitura dos XMLs, classificação e movimentação dos arquivos |
| `interface.py` | Interface gráfica (CustomTkinter) |
| `criar_config.py` | Script utilitário para gerar o `config.xlsx` modelo |
| `requirements.txt` | Dependências Python |
| `icone.ico` | Ícone do executável |

---

## Gerar executável (.exe)

Para distribuir sem Python instalado, use o PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icone.ico --name AutomacaoNFSe interface.py
```

O executável será gerado em `dist\AutomacaoNFSe.exe`.

---

## Observações técnicas

- A leitura dos XMLs usa o namespace oficial do SPED: `http://www.sped.fazenda.gov.br/nfse`
- Pastas de período são detectadas pelo padrão de data `DD-MM-AAAA` no nome da pasta
- Nomes de empresas são comparados em maiúsculas para evitar erros de case
- O arquivo `config.xlsx` e as pastas `build/`, `dist/` e `_teste_automacao/` são ignorados pelo git
