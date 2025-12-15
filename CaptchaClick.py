import pandas as pd
import pytesseract
from rich import print
import time
import numpy as np
from PIL import ImageGrab
import mousekey
import rapidfuzz
import pyperclip
import os
from datetime import datetime
import pyautogui  # Para o Scroll

# --- Configurações e Instâncias ---
mkey = mousekey.MouseKey()
mkey.enable_failsafekill('ctrl+e')

# Caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# --- Nossas Funções ---
def get_screenshot_tesser(minlen=2, lang='eng'):
    try:
        img_pil = ImageGrab.grab()
        img = np.array(img_pil)
        # Convertendo para DataFrame com dados do OCR
        df = pytesseract.image_to_data(img, lang=lang, output_type=pytesseract.Output.DATAFRAME)
        df = df.dropna(subset=["text"])
        df['text'] = df['text'].astype(str)
        df = df.loc[df['text'].str.len() >= minlen].reset_index(drop=True)
        return df
    except Exception as e:
        print(f"[ERRO na captura/OCR] {e}")
        return pd.DataFrame()

# --- Execução Principal do Robô ---
if __name__ == "__main__":
    try:
        # Lê a lista de Renavams
        df_renavams = pd.read_excel("lista_renavam.xlsx")
        lista_de_renavams = df_renavams["RENAVAM"].tolist()
        print(f"[INFO] {len(lista_de_renavams)} RENAVAMs carregados para processamento.")
    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha ao ler o Excel: {e}")
        exit()

    # --- PREPARAÇÃO DO ARQUIVO DE RELATÓRIO ---
    if not os.path.exists("Relatorios"):
        os.makedirs("Relatorios")

    data_hora_safe = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo_geral = os.path.join("Relatorios", f"Relatorio_Geral_{data_hora_safe}.txt")
    
    renavams_com_erro = []

    # --- CONFIGURAÇÃO DA HUMANIZAÇÃO (NOVA) ---
    # Mantendo X = -110 fixo e variando levemente o Y (-11, -10, -12)
    variacoes_clique = [
        (-110, -11), 
        (-110, -10), 
        (-110, -12)
    ]

    for i, renavam_atual in enumerate(lista_de_renavams):
        print(f"\n--- Iniciando Processo para RENAVAM: {renavam_atual} (Item {i+1}) ---")
        time.sleep(1)

        try:
            # --- ETAPA 1: Ler a tela e achar o campo ---
            df_tela = get_screenshot_tesser(lang='eng')
            if df_tela.empty: raise Exception("Tela vazia ou Tesseract falhou.")
            
            palavra_alvo, pontuacao, indice = rapidfuzz.process.extractOne("Renavam", df_tela['text'])
            if pontuacao < 80: raise Exception("Campo 'Renavam' não encontrado na tela.")
            
            info_renavam_label = df_tela.iloc[indice]
            
            # Clica no campo
            coord_x_campo = info_renavam_label['left'] + info_renavam_label['width'] + 100
            coord_y_campo = info_renavam_label['top'] + (info_renavam_label['height'] // 2) + 30
            
            mkey.move_to_natural(coord_x_campo, coord_y_campo)
            mkey.left_click()
            time.sleep(0.5)
            
            # Digita o Renavam
            for char in str(renavam_atual): 
                mkey.press_key(char)
                time.sleep(0.02)
            time.sleep(0.5)

            # --- ETAPA 2: Achar e Clicar no CAPTCHA (COM HUMANIZAÇÃO) ---
            ponto_referencia_x = info_renavam_label['left']
            ponto_referencia_y = info_renavam_label['top']
            
            candidatos_captcha = df_tela[df_tela['text'].str.contains('robot', case=False, na=False)]
            if candidatos_captcha.empty: raise Exception("Caixa 'Não sou um robô' não encontrada.")
            
            menor_distancia = float('inf')
            info_captcha_correto = None
            
            for index, row in candidatos_captcha.iterrows():
                captcha_x = row['left']
                captcha_y = row['top']
                distancia = np.sqrt((captcha_x - ponto_referencia_x)**2 + (captcha_y - ponto_referencia_y)**2)
                if distancia < menor_distancia:
                    menor_distancia = distancia
                    info_captcha_correto = row
            
            # === APLICAÇÃO DA VARIAÇÃO ===
            indice_variacao = i % 3 
            offset_x_escolhido, offset_y_escolhido = variacoes_clique[indice_variacao]
            
            print(f"[HUMANIZAÇÃO] Variação {indice_variacao + 1}: Offset X={offset_x_escolhido}, Y={offset_y_escolhido}")

            coord_x_captcha_box = info_captcha_correto['left'] + offset_x_escolhido
            coord_y_captcha_box = info_captcha_correto['top'] + offset_y_escolhido
            
            mkey.move_to_natural(coord_x_captcha_box, coord_y_captcha_box)
            mkey.left_click()
            print("[INFO] Checkbox do CAPTCHA clicado.")
            
            time.sleep(1)

            # --- ETAPA 2.5: CLIQUE INTERMEDIÁRIO ---
            offset_vertical_extra = 170   
            offset_horizontal_extra = 180 

            coord_x_extra = coord_x_captcha_box + offset_horizontal_extra
            coord_y_extra = coord_y_captcha_box + offset_vertical_extra

            print(f"Realizando clique intermediário...")
            mkey.move_to_natural(coord_x_extra, coord_y_extra)
            mkey.left_click()
            
            time.sleep(2)

            # --- ETAPA 3: Clicar em Consultar ---
            offset_vertical_consultar = 165 
            
            coord_x_final = coord_x_captcha_box
            coord_y_final = coord_y_captcha_box + offset_vertical_consultar

            print(f"Clicando em 'Consultar'...")
            mkey.move_to_natural(coord_x_final, coord_y_final)
            mkey.left_click()
            
            # --- ETAPA 4: ESPERAR CARREGAR E COPIAR ---
            print("Aguardando carregamento da página (5s)...") 
            time.sleep(5) 
            
            mkey.move_to_natural(500, 500)
            mkey.left_click()
            
            print("Copiando dados (Ctrl+A -> Ctrl+C)...")
            mkey.press_keys_simultaneously(['ctrl', 'a'])
            time.sleep(0.5)
            mkey.press_keys_simultaneously(['ctrl', 'c'])
            time.sleep(0.5)
            
            texto_copiado = pyperclip.paste()
            
            if not texto_copiado or len(texto_copiado) < 20:
                raise Exception("Falha ao copiar: Área de transferência vazia.")

            # --- ETAPA 5: SALVAR NO ARQUIVO GERAL ---
            print(f"[SUCESSO] Salvando dados do Renavam {renavam_atual}.")
            
            with open(nome_arquivo_geral, "a", encoding="utf-8") as arquivo:
                arquivo.write("\n" + "="*60 + "\n")
                arquivo.write(f" DATA: {datetime.now()} | RENAVAM CONSULTADO: {renavam_atual}\n")
                arquivo.write("="*60 + "\n\n")
                arquivo.write(texto_copiado)
                arquivo.write("\n\n")

        except Exception as e:
            print(f"[ERRO] Problema com Renavam {renavam_atual}: {e}")
            renavams_com_erro.append(f"{renavam_atual} - Motivo: {e}")

        finally:
            # --- ETAPA 6: RECARREGAR A PÁGINA ---
            print("Reiniciando a página para o próximo (Ctrl+L -> Enter)...")
            mkey.press_keys_simultaneously(['ctrl', 'l'])
            time.sleep(0.5)
            mkey.press_key('enter')
            
            time.sleep(4)

            # --- ETAPA 7: SCROLL DOWN ---
            print("Rolando a tela para posicionar o formulário...")
            pyautogui.scroll(-400) 
            time.sleep(1)

    # --- FASE FINAL ---
    print("\n--- Processamento finalizado. Gerando resumo de erros... ---")
    
    if renavams_com_erro:
        with open(nome_arquivo_geral, "a", encoding="utf-8") as arquivo:
            arquivo.write("\n\n" + "#"*60 + "\n")
            arquivo.write(" RESUMO DE ERROS / RENAVAMS NÃO PROCESSADOS \n")
            arquivo.write("#"*60 + "\n")
            for erro in renavams_com_erro:
                arquivo.write(f"- {erro}\n")
        print(f"[ATENÇÃO] {len(renavams_com_erro)} Renavams falharam. Veja o final do arquivo de texto.")
    else:
        with open(nome_arquivo_geral, "a", encoding="utf-8") as arquivo:
            arquivo.write("\n\n" + "#"*60 + "\n")
            arquivo.write(" SUCESSO TOTAL: Todos os Renavams foram processados.\n")
            arquivo.write("#"*60 + "\n")
        print("[SUCESSO] Todos os Renavams foram processados sem erros!")

    print(f"\nRelatório completo salvo em: {nome_arquivo_geral}")