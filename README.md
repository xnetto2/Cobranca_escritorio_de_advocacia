# Controle Financeiro

Sistema em Python para **cadastro de clientes**, **controle de parcelas contratuais**, **alerta de vencimentos** e **envio de mensagens de cobrança via WhatsApp**.

O sistema utiliza interface gráfica simples com `tkinter`, registra os dados em planilha Excel e permite acompanhar parcelas vencidas e vincendas de contratos jurídicos.

## Funcionalidades

- Cadastro de clientes
- Registro do número do processo
- Registro do valor total do contrato
- Cálculo automático do valor de cada parcela
- Definição da data da primeira parcela
- Geração automática das datas das demais parcelas
- Registro do número de celular do cliente
- Leitura de base existente em Excel
- Verificação automática de parcelas vencidas e parcelas vencendo no dia
- Envio de mensagens automáticas de cobrança via WhatsApp
- Salvamento dos dados atualizados em planilha Excel

## Tecnologias utilizadas

- Python
- pandas
- openpyxl
- pywhatkit
- tkinter

## Estrutura sugerida do projeto

```text
controle-financeiro/
├── controlefinanceiro.py
├── clientes.xlsx
├── README.md
├── requirements.txt
└── .gitignore
```

## Requisitos

- Python 3.10 ou superior
- WhatsApp Web configurado no navegador padrão
- Conexão com internet
- Conta do WhatsApp ativa no computador
- Biblioteca `tkinter` disponível no ambiente Python

## Instalação

Clone ou baixe o projeto e instale as dependências:

```bash
pip install -r requirements.txt
```

## Execução

Execute o arquivo principal:

```bash
python controlefinanceiro.py
```

## Como o sistema funciona

1. O programa abre uma interface simples usando caixas de diálogo.
2. Se existir a planilha `clientes.xlsx`, o sistema lê os registros já salvos.
3. O programa verifica parcelas vencidas e parcelas que vencem no dia.
4. Caso existam vencimentos, o usuário pode optar por enviar mensagens de cobrança.
5. Em seguida, o sistema permite cadastrar novos clientes.
6. Ao final, os dados são gravados novamente na planilha Excel.

## Formato dos dados cadastrados

O sistema registra, entre outros, os seguintes campos:

- Nome
- Processo
- Valor Total do Contrato
- Valor de Cada Parcela
- Data Primeira Parcela
- Parcelas
- Celular
- Data Parcela 1, Data Parcela 2, etc.

## Observações importantes

### 1. Caminho do arquivo Excel
No código atual, o caminho da planilha está fixado manualmente. Exemplo:

```python
caminho_diretorio = "clientes.xlsx"
```

### 2. Envio de WhatsApp
O envio de mensagens depende do `pywhatkit`, que normalmente abre o WhatsApp Web no navegador e tenta enviar a mensagem no horário programado.

### 3. Datas das parcelas
No código atual, as parcelas são geradas com incrementos de 30 dias. Isso é uma aproximação mensal e pode divergir do calendário real em meses com 28, 29 ou 31 dias.

## Melhorias futuras sugeridas

- Criar interface gráfica mais completa
- Validar melhor entradas numéricas e datas
- Tratar erros de digitação no cadastro
- Registrar status de pagamento de cada parcela
- Permitir edição e exclusão de clientes
- Adicionar relatórios financeiros
- Gerar executável com PyInstaller
