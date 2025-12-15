# Projeto-CAPTCHA 🤖

## Descrição do Projeto

Este projeto é um robô de automação (Bot) desenvolvido em Python que utiliza uma abordagem de **Visão Computacional (Computer Vision)** para localizar e interagir com caixas de seleção de CAPTCHA do tipo "Não sou um robô".

Diferente de métodos que tentam "quebrar" a lógica do CAPTCHA, este robô simula a interação humana: ele **lê a tela**, **encontra o texto** relevante e **move o mouse de forma natural** para clicar na caixa de seleção. Sua principal vantagem é a capacidade de operar em qualquer site, contanto que o elemento do CAPTCHA esteja visível na tela no momento da execução.

## Como Funciona?

O fluxo de automação segue os seguintes passos:

1.  **Captura de Tela:** O robô tira um "screenshot" da tela principal usando a biblioteca `Pillow`.
2.  **Reconhecimento de Caracteres (OCR):** Utiliza o `Tesseract OCR` para extrair todo o texto visível da imagem capturada.
3.  **Análise de Dados:** O texto extraído é carregado em um DataFrame do `Pandas` para facilitar a manipulação.
4.  **Busca por Similaridade:** Com o `RapidFuzz`, o robô procura pela palavra-chave (ex: "sou" ou "robô") na lista de textos extraídos, encontrando a correspondência mais provável.
5.  **Localização e Clique:** Uma vez que a palavra-chave é encontrada, o script obtém suas coordenadas (x, y) na tela.
6.  **Simulação Humana:** Por fim, a biblioteca `mousekey` é utilizada para mover o cursor do mouse até o alvo de forma natural e fluida (não instantânea) e realizar o clique.

## Tecnologias e Bibliotecas Utilizadas

- **Linguagem:** Python
- **Visão Computacional (OCR):** `Tesseract OCR`, `pytesseract`
- **Manipulação de Imagem:** `Pillow`
- **Análise de Dados:** `Pandas`, `NumPy`
- **Busca por Similaridade:** `RapidFuzz`
- **Automação e Simulação de Interação:** `mousekey`
- **Visualização no Terminal:** `rich`

## Como Usar o Projeto

1.  Clone este repositório: `git clone https://github.com/MrFelps/Projeto-CAPTCHA.git`
2.  Instale o Tesseract OCR no seu sistema (e o pacote de idioma Português).
3.  Crie e ative um ambiente virtual: `python -m venv .venv` e `.venv\Scripts\Activate.ps1`
4.  Instale as dependências: `pip install -r requirements.txt`
5.  Execute o script principal: `python CaptchaClick.py`
