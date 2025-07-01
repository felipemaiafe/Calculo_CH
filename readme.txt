# Calculadora de Carga Horária (CH)

Esta é uma aplicação de desktop com interface gráfica (GUI) desenvolvida em Python para automatizar o processo de cálculo da média de Carga Horária (CH) de servidores públicos. A ferramenta combina dados históricos de arquivos PDF (fichas financeiras) com dados recentes extraídos via web scraping de um portal governamental, gerando um relatório final em HTML.

O objetivo principal deste projeto é eliminar a necessidade de realizar um processo manual, repetitivo e sujeito a erros, economizando horas de trabalho e garantindo maior precisão nos resultados.

---

## 🚀 Principais Funcionalidades

*   **🖥️ Interface Gráfica Amigável:** Uma interface simples e intuitiva construída com Tkinter, permitindo que usuários não-técnicos utilizem a ferramenta facilmente.
*   **📄 Análise de PDF:** Extrai automaticamente dados financeiros e de cargo de múltiplos anos a partir de um arquivo PDF analítico.
*   **🤖 Automação Web com Selenium:** Realiza login em um portal web seguro, navega por diferentes menus e extrai dados históricos mês a mês, tudo em modo *headless* (sem exibir o navegador).
*   **⚙️ Gerenciamento Automático do Driver:** Utiliza `webdriver-manager` para baixar e gerenciar automaticamente a versão correta do ChromeDriver, eliminando a necessidade de atualizações manuais por parte do usuário.
*   **🧩 Instalação Automática de Dependências:** Verifica se a biblioteca `webdriver-manager` está instalada e, caso não esteja, tenta instalá-la automaticamente.
*   **📊 Lógica de Negócio Inteligente:**
    *   Mapeia anos com dados ausentes para anos de referência válidos.
    *   Compara valores do PDF com tabelas em um arquivo Excel para encontrar a CH correta.
    *   Inclui uma lógica de fallback para verificar cargos anteriores em caso de inconsistências salariais.
*   **📜 Log de Eventos em Tempo Real:** Exibe o progresso e possíveis erros em uma caixa de log na própria interface.
*   **📋 Relatório Final em HTML:** Consolida todos os dados coletados em uma tabela HTML bem formatada, que é salva localmente e aberta automaticamente no navegador padrão ao final do processo.

---

## 📋 Pré-requisitos

Antes de executar a aplicação, certifique-se de que você tem os seguintes itens instalados/configurados:

1.  **Python 3.8+**
2.  **Google Chrome:** O navegador precisa estar instalado, pois o Selenium irá controlá-lo.
3.  **Arquivo Excel de Vencimentos:** Um arquivo crucial chamado `VENCIMENTOS MAGISTÉRIO_1993-2014.xlsx` precisa estar acessível. O caminho para este arquivo está definido na constante `EXCEL_FILE_PATH` dentro do script. **A aplicação não funcionará sem ele.**

---

## 🛠️ Instalação e Execução

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/seu-usuario/seu-repositorio.git
    cd seu-repositorio
    ```

2.  **(Opcional, mas recomendado) Crie e ative um ambiente virtual:**
    ```bash
    python -m venv venv
    # Windows
    venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **Instale as dependências:**
    O script tentará instalar `webdriver-manager` automaticamente se estiver faltando. Para instalar todas as outras dependências, execute:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Execute o script:**
    ```bash
    python Calculo_CH_GEMINI.py
    ```

---

## 📖 Como Usar

1.  **Inicie a Aplicação:** Execute o script `Calculo_CH_GEMINI.py`.
2.  **Credenciais:** Preencha os campos `Login RHNet`, `Senha RHNet` e `CPF do Servidor`.
3.  **Selecionar Ficha Financeira:** Clique no botão "Selecionar PDF" e escolha o arquivo da ficha financeira anual analítica que deseja processar.
4.  **Calcular:** Clique no botão verde "CALCULAR". A aplicação começará o processo de automação. Você pode acompanhar o progresso no "Log de Eventos".
5.  **Cancelar:** Se necessário, clique no botão vermelho "CANCELAR" para interromper o processo.
6.  **Resultado:** Ao final, uma janela para salvar o arquivo HTML aparecerá. Após salvar, o arquivo será aberto automaticamente em seu navegador.

---


## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.