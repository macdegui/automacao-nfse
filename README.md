# Automação NFS-e

Sistema de automação para classificação e organização de Notas Fiscais de Serviço Eletrônica (NFS-e) do Portal Nacional.

---

## Instalação

### 1. Instalar dependências

Abra o Prompt de Comando na pasta do projeto e execute:

```
pip install -r requirements.txt
```

### 2. Criar o arquivo de configuração

Execute uma vez para gerar o `config.xlsx`:

```
python criar_config.py
```

Abra o `config.xlsx` e preencha o mapeamento de empresas:

| Nome no Portal NFSe | Nome na Pasta do Servidor |
|---|---|
| EL COMMERCE UTILIDADES E SERVICOS LTDA | El Commerce |
| EMPRESA EXEMPLO LTDA | Empresa Exemplo |

### 3. Executar o sistema

```
python interface.py
```

---

## Como usar

1. **Pasta de Origem** → Selecione a pasta onde o portal baixou as notas (ex: `C:\Users\...\Downloads\NFSe`)
2. **Pasta de Destino** → Selecione a raiz dos clientes no servidor (ex: `Z:\Clientes`)
3. **Mapeamento de Empresas** → Selecione o arquivo `config.xlsx`
4. Clique em **Executar Automação**

As configurações são salvas automaticamente para a próxima execução.

---

## Estrutura gerada no servidor

```
Z:\Clientes\
    Nome da Empresa\
        Fiscal\
            Fechamento Fiscal\
                2026\
                    Março\
                        Tomadas\
                            Retidas\
                            Não Retidas\
                        Emitidas\
```

---

## Regras de classificação

| Imposto | Condição |
|---|---|
| ISS | `tpRetISSQN = 1` |
| PIS/COFINS | `tpRetPisCofins = 1` |
| IRRF | `vRetIRRF > 0` |
| INSS | `vRetCP > 0` |
| CSLL/PIS/COFINS | `vRetCSLL > 0` |

Qualquer retenção identificada classifica a nota como **Retida**.

---

## Gerar executável (.exe)

```
pip install pyinstaller
pyinstaller --onefile --windowed --name "AutomacaoNFSe" interface.py
```

O arquivo `AutomacaoNFSe.exe` será gerado na pasta `dist\`.
