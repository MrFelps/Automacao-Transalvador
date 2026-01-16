# Projeto Doctor Strange: Automação Transalvador (RPA) 🤖

## 📋 Descrição do Projeto

Este projeto é uma solução de **RPA (Robotic Process Automation)** desenvolvida em Python para realizar consultas em massa de veículos (RENAVAM) no portal TransOnline. 

Diferente de bots tradicionais que buscam elementos ocultos no HTML (Web Scraping puro), o **Doctor Strange** utiliza uma abordagem híbrida com **Visão Computacional**. Ele "enxerga" a tela do computador como um humano, identifica campos de formulário e desafios de Captcha visualmente, e interage através de simulação de mouse humanizada.

O objetivo principal é eliminar o trabalho repetitivo da equipe de Operações da Frota 162, garantindo velocidade, resiliência contra bloqueios (WAF) e precisão de dados.

---

## ⚙️ Como Funciona? (Fluxo de Execução)

O robô opera em um loop contínuo inteligente, seguindo as etapas abaixo para cada RENAVAM da lista:

1.  **Captura de Tela (Vision):** O script utiliza a biblioteca `Pillow (ImageGrab)` para tirar um "snapshot" da tela em tempo real.
2.  **Processamento OCR (Leitura):** O `Tesseract OCR` analisa a imagem e extrai textos e suas coordenadas espaciais (onde está cada palavra na tela).
3.  **Busca Semântica & Regex:**
    * Localiza o campo "RENAVAM" usando `RapidFuzz` (para tolerar pequenas variações de fonte/renderização).
    * Identifica o Captcha usando **Regex Bilíngue**, detectando tanto *"Não sou um robô"* quanto *"I'm not a robot"*.
4.  **Movimentação Humanizada:** A biblioteca `mousekey` calcula curvas de Bezier para mover o mouse de forma natural (não retilínea) e com velocidade variável, enganando sistemas anti-bot.
5.  **Resiliência (Modo Demolidor):** Caso o OCR falhe na leitura do texto do Captcha, o robô ativa um algoritmo de **Inferência de Posição**, calculando onde o clique *deveria* ocorrer com base na posição do campo anterior, garantindo que o processo não pare.

---

## 🛠️ Tecnologias e Bibliotecas Utilizadas

* **Linguagem:** Python 3.x
* **Visão Computacional:** `Tesseract OCR` (Engine) + `pytesseract` (Wrapper)
* **Manipulação de Imagem:** `Pillow` (PIL)
* **Análise de Dados:** `Pandas` (ETL do Excel) + `NumPy`
* **Interação de Interface (GUI):** `MouseKey` (Movimento Humano) + `PyAutoGUI` (Teclado/Scroll)
* **Busca por Similaridade:** `RapidFuzz` + `Re` (Expressões Regulares)
* **Interface de Comando (CLI):** `Rich` (Logs coloridos e Barra de Progresso)

---

## 🧠 Diferenciais de Engenharia

### 🛡️ Proteção de IP e Humanização
Para evitar o "banimento" do IP da empresa, o código implementa:
* **Fadiga Simulada:** O robô fica "cansado" após X consultas, aumentando os intervalos de espera.
* **Variação de Movimento:** Nunca clica no mesmo pixel exato; usa `random.uniform` para criar micro-variações nas coordenadas.

### 🎯 Teste de Mira (Calibragem)
Inclui um módulo auxiliar (`TesteMira.py`) que permite validar as coordenadas em um ambiente estático (Print Screen) antes da execução real, prevenindo cliques erráticos em produção.

### 🔒 Segurança de Dados
O projeto segue rigorosas práticas de segurança:
* Dados sensíveis (`lista_renavam.xlsx`) são ignorados pelo Git.
* Logs de execução são sanitizados.
* Implementação de *Kill-Switch* (`Ctrl + E`) para parada de emergência imediata.

---

## 🚀 Como Usar o Projeto

### Pré-requisitos
* Python 3.10+
* Tesseract OCR instalado no Windows (`C:\Program Files\Tesseract-OCR\tesseract.exe`)

### Instalação

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/MrFelps/Automacao-Transalvador.git](https://github.com/MrFelps/Automacao-Transalvador.git)
    ```

2.  **Prepare o Ambiente Virtual:**
    ```powershell
    python -m venv .venv
    .\.venv\Scripts\Activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Execute o Robô:**
    * Abra o site Transalvador em tela cheia.
    * Rode o comando e **clique na página** imediatamente para dar foco.
    ```bash
    python CaptchaClick.py
    ```
