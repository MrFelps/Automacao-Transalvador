import pandas as pd
import pytesseract
from rich import print
import time
import numpy as np
from PIL import ImageGrab
import mousekey
import rapidfuzz
import pyperclip

# --- Configurações e Instâncias ---
mkey = mousekey.MouseKey()
mkey.enable_failsafekill('ctrl+e')
pytesseract.pytesseract.tesseract_cmd = r"C:\Arquivos de Programas\Tesseract-OCR\tesseract.exe"

# --- Nossas Funções ---
def get_screenshot_tesser(minlen=2, lang='eng'):
    try:
        img_pil = ImageGrab.grab()
        img = np.array(img_pil)
        df = pytesseract.image_to_data(img, lang=lang, output_type=pytesseract.Output.DATAFRAME)
        df = df.dropna(subset=["text"])
        # Garante que 'text' é string antes de filtrar
        df['text'] = df['text'].astype(str)
        df = df.loc[df['text'].str.len() >= minlen].reset_index(drop=True)
        return df
    except Exception as e:
        print(f"[ERRO na captura/OCR] {e}")
        return pd.DataFrame()

# --- Execução Principal do Robô ---
if __name__ == "__main__":
    try:
        df_renavams = pd.read_excel("lista_renavam.xlsx")
        lista_de_renavams = df_renavams["RENAVAM"].tolist()
        print(f"[INFO] {len(lista_de_renavams)} RENAVAMs carregados.")
    except FileNotFoundError:
        print("[ERRO] Arquivo 'lista_renavam.xlsx' não encontrado!")
        exit()
    except KeyError:
        print("[ERRO] Coluna 'RENAVAM' não encontrada.")
        exit()
    except Exception as e:
        print(f"[ERRO] Falha ao ler o Excel: {e}")
        exit()

    resultados_completos = []

    for renavam_atual in lista_de_renavams:
        print(f"\n--- Processando RENAVAM: {renavam_atual} ---")
        print("Iniciando em 1 segundo...")
        time.sleep(1)
        texto_resultado_atual = "ERRO NA EXECUCAO"

        try:
            # ETAPAS 1, 2 e 3 (INTOCADAS)
            df_tela = get_screenshot_tesser(lang='eng')
            if df_tela.empty: raise Exception("Tesseract não encontrou texto na tela.")
            palavra_alvo, pontuacao, indice = rapidfuzz.process.extractOne("Renavam", df_tela['text'])
            if pontuacao < 85: raise Exception("Rótulo 'Renavam' não encontrado.")
            info_renavam_label = df_tela.iloc[indice]
            coord_x_campo = info_renavam_label['left'] + info_renavam_label['width'] + 100
            coord_y_campo = info_renavam_label['top'] + (info_renavam_label['height'] // 2) + 30
            mkey.move_to_natural(coord_x_campo, coord_y_campo); mkey.left_click(); time.sleep(1)
            for char in str(renavam_atual): mkey.press_key(char); time.sleep(0.02)
            time.sleep(1)
            ponto_referencia_x = info_renavam_label['left']; ponto_referencia_y = info_renavam_label['top']
            candidatos_captcha = df_tela[df_tela['text'].str.contains('robot', case=False, na=False)]
            if candidatos_captcha.empty: raise Exception("Nenhum texto 'robot' encontrado.")
            menor_distancia = float('inf'); info_captcha_correto = None
            for index, row in candidatos_captcha.iterrows():
                captcha_x = row['left']; captcha_y = row['top']
                distancia = np.sqrt((captcha_x - ponto_referencia_x)**2 + (captcha_y - ponto_referencia_y)**2)
                if distancia < menor_distancia: menor_distancia = distancia; info_captcha_correto = row
            # Guardamos as coordenadas exatas do CLIQUE 2 (Caixa do CAPTCHA)
            coord_x_captcha_box = info_captcha_correto['left'] - 80
            coord_y_captcha_box = info_captcha_correto['top'] - 15
            mkey.move_to_natural(coord_x_captcha_box, coord_y_captcha_box); mkey.left_click();
            print("[INFO] CAPTCHA (2º clique) Clicado.")
            time.sleep(1)
            coord_x_acessibilidade = 499
            coord_y_acessibilidade = 693
            print(f"Movendo para o clique de acessibilidade (3º clique)...")
            mkey.move_to_natural(coord_x_acessibilidade, coord_y_acessibilidade); mkey.left_click()
            print("[INFO] Clique de acessibilidade realizado.")
            time.sleep(1)

            # --- ETAPA 4: CLIQUE FINAL RELATIVO AO CAPTCHA BOX (VOLTAMOS A ESTA LÓGICA) ---
            print("Calculando posição do clique final (relativo ao CAPTCHA)...")
            time.sleep(2) # Pausa antes do último clique

            # === SEU PAINEL DE CONTROLE (4º CLIQUE - ABAIXO DO CAPTCHA BOX) ===
            # Ajuste este número para descer mais (+) ou menos (-)
            offset_vertical_4 = 150 # Exemplo: 150 pixels ABAIXO do clique do CAPTCHA Box
            # (Opcional) Ajuste horizontal se necessário (+ Direita, - Esquerda)
            offset_horizontal_4 = 0
            # =============================================================

            # Usa as coordenadas do CLIQUE 2 (CAPTCHA Box) como referência
            coord_x_final = coord_x_captcha_box + offset_horizontal_4
            coord_y_final = coord_y_captcha_box + offset_vertical_4

            print(f"Movendo para o clique final (4º clique) em ({coord_x_final}, {coord_y_final})...")
            mkey.move_to_natural(coord_x_final, coord_y_final)
            mkey.left_click()
            print("[INFO] Clique final ('Consultar') realizado.")

            # --- ETAPA 5: COPIAR TODO O TEXTO DA PÁGINA ---
            print("Aguardando 5 segundos para a página de resultados carregar...")
            time.sleep(5)
            print("Simulando Ctrl+A e Ctrl+C...")
            mkey.press_keys_simultaneously(['ctrl', 'a']); time.sleep(1)
            mkey.press_keys_simultaneously(['ctrl', 'c']); time.sleep(1)
            texto_resultado_atual = pyperclip.paste()
            if not texto_resultado_atual: texto_resultado_atual = "NENHUM TEXTO FOI COPIADO"
            print(f"[SUCESSO] Texto da consulta para {renavam_atual} copiado.")

        except Exception as e:
            print(f"[ERRO] Falha no processo para {renavam_atual}: {e}")
            texto_resultado_atual = f"ERRO: {e}"

        finally:
            resultados_completos.append(f"--- RENAVAM: {renavam_atual} ---\n{texto_resultado_atual}\n\n")
            print("Aguardando 3 segundos para o próximo RENAVAM...")
            time.sleep(3)

    # --- FASE FINAL ---
    print("\n--- Automação da lista concluída! Salvando relatório... ---")
    nome_arquivo_relatorio = "relatorio_consultas.txt"
    with open(nome_arquivo_relatorio, "w", encoding="utf-8") as f:
        f.writelines(resultados_completos)
    print(f"Relatório salvo com sucesso como '{nome_arquivo_relatorio}'")