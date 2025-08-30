# RPA - Automação de Medição de Horas (Projeto ZRC)

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)

Este robô é uma versão especializada para automatizar o processo de medição de horas exclusivamente para o **projeto ZRC**. Ele possui uma lógica de extração com período de datas customizado.

## 🚀 Funcionalidades Principais

- **Automação Web**: Acessa a plataforma PSOffice, realiza login e navega até a tela de relatórios.
- **Período de Datas Customizado**: Calcula e aplica dinamicamente o filtro de data do **dia 15 do mês anterior até o dia 14 do mês atual**.
- **Extração de Dados**: Baixa o relatório de horas para o período customizado.
- **Desbloqueio e Processamento**: Habilita a edição do arquivo Excel baixado, filtra os dados apenas para o projeto ZRC e processa os apontamentos relevantes.
- **Geração de Relatório Detalhado**:
  - Cria uma pasta para o mês da medição (ex: `relatorios/medicao/Agosto-2025/`).
  - Gera um único arquivo Excel (`ZRC...-relatorio-horas-Mes-Ano.xlsx`).
  - O arquivo contém abas separadas para cada profissional, com uma linha de "Total de Horas" no topo.
- **Envio de E-mail**: Envia um único e-mail para os stakeholders do projeto ZRC, contendo uma tabela resumo no corpo e o relatório Excel detalhado em anexo.
- **Monitoramento**: Ao final da execução, envia um e-mail de status com o tempo de cada etapa e detalhes de eventuais erros.
- **Agendamento**: Projetado para ser executado de forma autônoma no dia 15 de cada mês.

## 🛠️ Pré-requisitos

- [Python](https://www.python.org/downloads/) (versão 3.8 ou superior)
- [Git](https://git-scm.com/downloads/) (opcional)
- Google Chrome
- Microsoft Excel instalado

## ⚙️ Instalação

1.  **Clone o repositório:**

    ```bash
    git clone https://github.com/phlioni/robo-medicao-horas-zrc.git
    cd robo-medicao-horas-zrc
    ```

2.  **Crie e ative o ambiente virtual:**

    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

## 📝 Configuração

1.  **Crie o arquivo `config.py`**: Se não existir, faça uma cópia do arquivo `config.py.template` e renomeie-a para `config.py`.
2.  **Preencha `config.py`**: Abra o arquivo e preencha todas as variáveis:
    - `SITE_LOGIN` e `SITE_SENHA`.
    - `EMAIL_REMETENTE` e `EMAIL_SENHA`.
    - `ZRC_DESTINATARIO_PRINCIPAL` e `ZRC_DESTINATARIOS_COPIA`.
    - `STATUS_EMAIL_DESTINATARIO`.
    - Dados da sua assinatura.
3.  **Pasta de Imagens**: Crie uma pasta chamada `imagens` e coloque os logos `mosten.png` e `selos.png` dentro dela.

## ▶️ Execução

Para executar o robô, utilize o script de inicialização `run_robot.bat`.

- **Execução Manual:** Dê um duplo clique no arquivo `run_robot.bat`.
- **Execução Silenciosa:** Edite o arquivo `automacao_web.py` e remova o `#` da linha `options.add_argument("--headless")`.

## 🗓️ Agendamento (Windows)

- **Ação**: Aponte a tarefa no Agendador do Windows para o arquivo `run_robot.bat` dentro da pasta `rpa_zrc`.
- **Disparador (Gatilho)**: Configure o disparador para **"Mensalmente"**, nos **"Dias:" `15`**.

## 📂 Estrutura do Projeto
rpa_zrc/
├── imagens/
│   ├── mosten.png
│   └── selos.png
├── relatorios/             (criada pelo robô)
├── venv/                   (ambiente virtual)
├── .gitignore
├── automacao_web.py
├── config.py               (com suas senhas - ignorado pelo Git)
├── config.py.template      (modelo seguro para o Git)
├── envio_email.py
├── excel_handler.py
├── main.py
├── processamento_dados.py
├── README.md
├── requirements.txt
└── run_robot.bat