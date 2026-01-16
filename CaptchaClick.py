import pandas as pd
import pytesseract
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn
from rich.theme import Theme
import time
import numpy as np
from PIL import ImageGrab
import mousekey
import rapidfuzz
import pyperclip
import os
import sys
from datetime import datetime
import pyautogui 
import random 

# Configuração visual do terminal
custom_theme = Theme({
    "info": "dim cyan",
    "warning": "magenta",
    "danger": "bold red",
    "success": "bold green"
})
console = Console(theme=custom_theme)

# =================================================================================
#               CONFIGURAÇÕES GERAIS 
# =================================================================================

VARIACOES_RENAVAM = [(0, 0), (15, 5), (-15, -5)]
VARIACOES_CAPTCHA = [(-100, 5), (-110, -5), (-105, -2)]
VARIACOES_CONSULTAR = [(0, 0), (0, 5), (0, -3)]

OFFSET_EXTRA_X = 180
OFFSET_EXTRA_Y = 170
OFFSET_CONSULTAR_Y = 165
SCROLL_AMOUNT = -400

# =================================================================================
#                       FUNÇÕES AUXILIARES
# =================================================================================

mkey = mousekey.MouseKey()
mkey.enable_failsafekill('ctrl+e')
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def espera_humana(minimo, maximo):
    time.sleep(random.uniform(minimo, maximo))

def mover_com_hesitacao(x, y):
    espera_humana(0.1, 0.3)
    mkey.move_to_natural(x, y)
    espera_humana(0.2, 0.4)

def get_screenshot_tesser(minlen=2, lang='eng'):
    try:
        img_pil = ImageGrab.grab()
        img = np.array(img_pil)
        # O lang='eng' costuma ler bem português também para palavras simples como 'robô'
        df = pytesseract.image_to_data(img, lang=lang, output_type=pytesseract.Output.DATAFRAME)
        df = df.dropna(subset=["text"])
        df['text'] = df['text'].astype(str)
        df = df.loc[df['text'].str.len() >= minlen].reset_index(drop=True)
        return df
    except Exception as e:
        return pd.DataFrame()

# =================================================================================
#                           EXECUÇÃO PRINCIPAL
# =================================================================================
if __name__ == "__main__":
    console.clear()
    console.print("[bold green]INICIANDO ROBÔ - MODO DIAGNÓSTICO BILÍNGUE[/bold green]")
    console.print("[dim]Pressione 'Ctrl + E' para parar. CLIQUE NO SITE AO INICIAR![/dim]\n")

    # --- 1. CARREGAMENTO DO EXCEL ---
    try:
        df_renavams = pd.read_excel("lista_renavam.xlsx")
        lista_de_renavams = df_renavams["RENAVAM"].tolist()
        total_renavams = len(lista_de_renavams)
        console.print(f"[info]INFO:[/info] {total_renavams} RENAVAMs carregados.\n")
    except Exception as e:
        console.print(f"[danger]ERRO CRÍTICO:[/danger] Falha ao ler Excel: {e}")
        sys.exit()

    if not os.path.exists("Relatorios"):
        os.makedirs("Relatorios")

    data_hora_safe = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo_geral = os.path.join("Relatorios", f"Relatorio_Geral_{data_hora_safe}.txt")
    renavams_com_erro = []

    # --- 2. LOOP COM BARRA DE PROGRESSO ---
    with Progress(
        SpinnerColumn(),
        TextColumn("[bold blue]{task.description}"),
        BarColumn(bar_width=None, complete_style="green", finished_style="green"),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeRemainingColumn(),
        console=console,
        transient=False 
    ) as progress:
        
        task_principal = progress.add_task(f"[Aguardando...]", total=total_renavams)

        for i, renavam_atual in enumerate(lista_de_renavams):
            
            progress.update(task_principal, description=f"Processando {i+1}/{total_renavams} | Renavam: {renavam_atual}")
            
            # --- CORREÇÃO VISUAL: Linha simples que não trava ---
            progress.console.print(f"\n[dim]----------------------------------------------------------------[/dim]")
            progress.console.print(f"[bold cyan]▶ INICIANDO RENAVAM: {renavam_atual}[/bold cyan]")

            # Define Velocidade (Lógica de Fadiga)
            tempo_min, tempo_max = (5.5, 7.0) if i < 10 else (10.0, 14.0)
            
            # Sorteio Combo
            idx_ren = random.randint(0, 2)
            idx_cap = random.randint(0, 2)
            idx_con = random.randint(0, 2)
            mapa = {0: 'A', 1: 'B', 2: 'C'}
            
            progress.console.print(f"   [info]Variável Escolhida:[/info] Renavam[{mapa[idx_ren]}] | Captcha[{mapa[idx_cap]}] | Consultar[{mapa[idx_con]}]")

            espera_humana(1.0, 2.0)

            try:
                # --- [A] SCAN ---
                progress.console.print("   [info]🔍 Escaneando tela (OCR)...[/info]")
                info_renavam_label = None
                df_tela = pd.DataFrame() 

                for tentativa in range(1, 4): 
                    df_tela = get_screenshot_tesser(lang='eng')
                    if not df_tela.empty:
                        palavra_alvo, pontuacao, indice = rapidfuzz.process.extractOne("Renavam", df_tela['text'])
                        if pontuacao >= 80:
                            info_renavam_label = df_tela.iloc[indice]
                            break 
                    time.sleep(1.0)

                if info_renavam_label is None: 
                    progress.console.print(f"\n[danger]⛔ ERRO FATAL: Campo 'Renavam' não encontrado![/danger]")
                    progress.console.print("[warning]Dica: Clique na janela do site imediatamente após iniciar.[/warning]")
                    sys.exit() 

                # --- [B] DIGITAÇÃO ---
                base_x = info_renavam_label['left'] + info_renavam_label['width'] + 100
                base_y = info_renavam_label['top'] + (info_renavam_label['height'] // 2) + 30
                vx, vy = VARIACOES_RENAVAM[idx_ren]
                
                progress.console.print(f"   [info]🖱️ Clicando no campo (Variação {mapa[idx_ren]})...[/info]")
                mover_com_hesitacao(base_x + vx, base_y + vy)
                mkey.left_click()
                mkey.left_click()
                
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.press('backspace')
                
                progress.console.print(f"   [info]⌨️ Digitando Renavam...[/info]")
                for char in str(renavam_atual): 
                    mkey.press_key(char)
                    time.sleep(random.uniform(0.02, 0.09)) 
                espera_humana(0.5, 1.0)

                # --- [C] CAPTCHA (AGORA BILÍNGUE) ---
                p_ref_x = info_renavam_label['left']
                p_ref_y = info_renavam_label['top']
                
                # Procura por: "robot" OU "robô" OU "robo"
                # regex=True ativa a busca inteligente
                padrao_busca = r"rob[oô]|robot"
                cands = df_tela[df_tela['text'].str.contains(padrao_busca, case=False, regex=True, na=False)]
                
                if cands.empty: 
                    progress.console.print(f"\n[danger]⛔ ERRO FATAL: Captcha ('robô' ou 'robot') não achado![/danger]")
                    sys.exit()

                best_dist = float('inf')
                best_row = None
                for _, row in cands.iterrows():
                    d = np.sqrt((row['left'] - p_ref_x)**2 + (row['top'] - p_ref_y)**2)
                    if d < best_dist:
                        best_dist = d
                        best_row = row
                
                vx_c, vy_c = VARIACOES_CAPTCHA[idx_cap]
                fx_cap = best_row['left'] + vx_c
                fy_cap = best_row['top'] + vy_c
                
                progress.console.print(f"   [info]🖱️ Clicando no Captcha (Variação {mapa[idx_cap]})...[/info]")
                mover_com_hesitacao(fx_cap, fy_cap)
                mkey.left_click()
                
                progress.console.print("   [info]⏳ Aguardando verificação do Captcha...[/info]")
                espera_humana(1.5, 2.5) 

                # --- [D] INTERMEDIÁRIO ---
                mover_com_hesitacao(fx_cap + OFFSET_EXTRA_X, fy_cap + OFFSET_EXTRA_Y)
                mkey.left_click()
                espera_humana(1.0, 1.5)

                # --- [E] CONSULTAR ---
                vx_b, vy_b = VARIACOES_CONSULTAR[idx_con]
                fx_con = fx_cap + vx_b
                fy_con = fy_cap + OFFSET_CONSULTAR_Y + vy_b
                
                progress.console.print(f"   [info]🖱️ Clicando em Consultar (Variação {mapa[idx_con]})...[/info]")
                mover_com_hesitacao(fx_con, fy_con)
                mkey.left_click()
                
                # --- [F] ESPERA ---
                progress.console.print(f"   [warning]⏳ Aguardando carregamento (aprox {tempo_max}s)...[/warning]")
                espera_humana(tempo_min, tempo_max) 
                
                # --- [G] CÓPIA ---
                texto_final = ""
                for tentativa_copia in range(1, 3): 
                    progress.console.print(f"   [info]📋 Tentativa de Cópia {tentativa_copia}...[/info]")
                    mkey.move_to_natural(500, 500)
                    mkey.left_click()
                    
                    pyperclip.copy("") 
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.5)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.5)
                    
                    texto_copiado = pyperclip.paste()
                    
                    if len(texto_copiado) > 100:
                        texto_final = texto_copiado
                        progress.console.print(f"   [success]✔ Sucesso! {len(texto_final)} caracteres copiados.[/success]")
                        break 
                    else:
                        progress.console.print(f"   [warning]⚠️ Falha na cópia. Aguardando mais...[/warning]")
                        espera_humana(3.0, 5.0)
                
                if len(texto_final) < 100:
                    raise Exception("Site demorou demais ou bloqueio soft.")

                # --- [H] SALVAR ---
                try:
                    with open(nome_arquivo_geral, "a", encoding="utf-8") as arquivo:
                        arquivo.write("\n" + "="*60 + "\n")
                        arquivo.write(f" DATA: {datetime.now()} | RENAVAM: {renavam_atual}\n")
                        arquivo.write("="*60 + "\n\n")
                        arquivo.write(str(texto_final)) 
                        arquivo.write("\n\n")
                except Exception as e_save:
                     progress.console.print(f"[danger]ERRO DE ARQUIVO: Não foi possível salvar![/danger]")
                     sys.exit()

                # Limpa seleção
                mkey.move_to_natural(500, 500) 
                mkey.left_click()

            except KeyboardInterrupt:
                console.print("\n[danger]PARADA MANUAL PELO USUÁRIO[/danger]")
                sys.exit()

            except SystemExit:
                raise 

            except Exception as e:
                progress.console.print(f"[danger]✖ Erro neste Renavam: {e}[/danger]")
                renavams_com_erro.append(f"{renavam_atual} - {e}")

            finally:
                # --- [I] PREPARAR PRÓXIMO ---
                progress.console.print("   [dim]🔄 Preparando próxima consulta...[/dim]")
                mkey.move_to_natural(500, 500)
                mkey.left_click()
                pyautogui.hotkey('ctrl', 'l')
                time.sleep(0.5)
                pyautogui.press('enter')
                
                tempo_reload = 5.0 if i < 10 else 9.0
                espera_humana(tempo_reload, tempo_reload + 2.0) 

                pyautogui.scroll(SCROLL_AMOUNT)
                espera_humana(1.5, 2.5)

                progress.advance(task_principal)

    console.print("\n" + "="*60)
    console.print("[success]PROCESSAMENTO FINALIZADO![/success]")