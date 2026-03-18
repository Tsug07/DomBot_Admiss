<div align="center">

<img src="assets/DomBot_New.png" alt="DomBot Logo" width="120"/>

# DomBot Admissional

**Automação RPA para emissão de Contratos Admissionais no sistema Domínio Folha**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Framework](https://img.shields.io/badge/GUI-CustomTkinter-1ABC9C?style=for-the-badge)](https://github.com/TomSchimansky/CustomTkinter)
[![Automation](https://img.shields.io/badge/RPA-pywinauto-9B59B6?style=for-the-badge)](https://pywinauto.readthedocs.io/)
[![Status](https://img.shields.io/badge/Status-Estável-2ECC71?style=for-the-badge)]()
[![Author](https://img.shields.io/badge/Autor-Hugo%20L.%20Almeida-E74C3C?style=for-the-badge)]()

</div>

---

## Visão Geral

O **DomBot Admissional** é uma ferramenta de **Automação Robótica de Processos (RPA)** desenvolvida para automatizar a emissão de relatórios de contratos admissionais no sistema **Domínio Folha**. A partir de uma planilha Excel, o bot navega pelo sistema, preenche os dados de cada funcionário e gera os PDFs dos contratos automaticamente, eliminando trabalho manual repetitivo e reduzindo erros humanos.

---

## Funcionalidades

- **Importacao de planilha Excel** com lista de funcionários a processar
- **Troca automatica de empresa** via atalho F8 do Domínio Folha
- **Preenchimento automatico** do código do funcionário e tipo de contrato
- **Suporte a dois tipos de contrato**: Experiência (`E`) e Prazo Indeterminado (`I`)
- **Geração e salvamento de PDF** com nome configurável por planilha
- **Controles de execução**: Iniciar, Pausar/Retomar e Parar
- **Dashboard em tempo real** com contadores de Sucesso, Erros, Empresa atual e Tempo decorrido
- **Barra de progresso** visual e percentual
- **Log colorido** em tempo real com timestamps
- **Export de logs** para arquivo `.txt`
- **Logs em arquivo** separados por data (sucesso e erros)
- **Retomada de processamento** a partir de qualquer linha do Excel
- **Tratamento inteligente de diálogos** de erro do sistema Domínio
- **Reconexão automática** em caso de perda de foco/handle da janela
- **Interface Dark Mode** moderna com tema personalizado

---

## Pré-requisitos

| Requisito | Versão Mínima |
|-----------|---------------|
| Python | 3.10+ |
| Sistema Operacional | Windows 10/11 |
| Domínio Folha | Qualquer versão instalada e aberta |

---

## Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/DomBot_Admiss.git
cd DomBot_Admiss
```

### 2. Crie e ative o ambiente virtual (recomendado)

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Instale as dependências

```bash
pip install customtkinter pandas pywinauto pywin32 Pillow openpyxl
```

---

## Estrutura do Projeto

```
DomBot_Admiss/
├── DomBot_Admiss.py        # Aplicação principal
├── assets/
│   ├── DomBot_New.png      # Logo da aplicação
│   └── favicon.ico         # Ícone da janela
├── logs/
│   ├── success_YYYY-MM-DD.log   # Log de sucessos por data
│   └── error_YYYY-MM-DD.log     # Log de erros por data
└── README.md
```

---

## Formato da Planilha Excel

A planilha de entrada deve conter obrigatoriamente as seguintes colunas:

| Coluna | Tipo | Descrição | Exemplo |
|--------|------|-----------|---------|
| `Nº` | Inteiro | Número da empresa no Domínio | `1042` |
| `Cod.Funcionário` | Inteiro | Código do funcionário | `315` |
| `Tipo de Contrato` | Texto | `E` = Experiência / `I` = Indeterminado | `E` |
| `Documento` | Texto | Nome do arquivo PDF a ser salvo | `CONTRATO_JOAO_SILVA` |

> **Obs:** Colunas adicionais são ignoradas. A coluna `Funcionário` é opcional e usada apenas para exibição nos logs.

**Exemplo de planilha:**

| Nº | Cod.Funcionário | Funcionário | Tipo de Contrato | Documento |
|----|-----------------|-------------|-----------------|-----------|
| 1042 | 315 | João Silva | E | CONTRATO_JOAO_SILVA |
| 2087 | 201 | Maria Costa | I | CONTRATO_MARIA_COSTA |

---

## Como Usar

1. **Abra o Domínio Folha** e faça login normalmente antes de iniciar o bot
2. **Execute o DomBot:**
   ```bash
   python DomBot_Admiss.py
   ```
3. **Selecione o arquivo Excel** clicando em `Procurar`
4. **Defina a linha inicial** (padrão: `2` — primeira linha após o cabeçalho)
5. Clique em **`▶ Iniciar`** para começar o processamento

### Controles durante a execução

| Botão | Ação |
|-------|------|
| `▶ Iniciar` | Inicia o processamento |
| `⏸ Pausar` | Pausa após a linha atual; clique novamente para retomar |
| `⏹ Parar` | Interrompe a execução com segurança |

---

## Arquitetura

```
┌─────────────────────────────────────────────────┐
│                 AutomacaoGUI                    │
│  Interface gráfica (CustomTkinter)              │
│  - Controles de execução                        │
│  - Dashboard de estatísticas                    │
│  - Log visual colorido                          │
│  - Preview da planilha Excel                    │
└──────────────────────┬──────────────────────────┘
                       │ Thread separada
                       ▼
┌─────────────────────────────────────────────────┐
│              DominioAutomation                  │
│  Engine de automação (pywinauto + win32)        │
│  - connect_to_dominio()                         │
│  - handle_empresa_change()                      │
│  - processar_relatorio_admissional()            │
│  - salvar_pdf()                                 │
│  - handle_error_dialogs()                       │
│  - cleanup_windows()                            │
└─────────────────────────────────────────────────┘
```

A GUI roda na thread principal e a automação em uma **thread daemon** separada, garantindo que a interface permaneça responsiva durante o processamento.

---

## Logs

Todos os eventos são registrados em dois destinos:

**Interface visual** — Log colorido em tempo real com ícones:
- `✅` Sucesso
- `❌` Erro
- `⚠️` Aviso
- `ℹ️` Informação
- `⏳` Processando

**Arquivos de log** em `./logs/`:
- `success_YYYY-MM-DD.log` — registros de processamentos bem-sucedidos
- `error_YYYY-MM-DD.log` — registros detalhados de erros

---

## Dependências

```
customtkinter    # Interface gráfica moderna (dark mode)
pandas           # Leitura e manipulação de planilhas Excel
pywinauto        # Automação de janelas Windows (backend UIA)
pywin32          # APIs nativas do Windows (win32gui, win32con)
Pillow           # Processamento de imagens (logo circular)
openpyxl         # Engine para leitura de .xlsx pelo pandas
```

---

## Observacoes Tecnicas

- O bot utiliza o backend **UIA (UI Automation)** do pywinauto para maior compatibilidade com o Domínio Folha
- A janela do Domínio Folha deve estar **aberta e visível** antes de iniciar o bot
- O campo `Linha inicial` permite **retomar** um processamento interrompido sem reprocessar linhas já concluídas
- O `smart_sleep` é um sleep interruptível que respeita pausa e parada imediatamente
- O `wait_for_condition` é um polling genérico com timeout configurável por operação

---

## Licenca

Este projeto é de uso **interno e proprietário**.
Todos os direitos reservados a **Hugo L. Almeida**.

---

<div align="center">

Desenvolvido por **Hugo L. Almeida**

</div>
