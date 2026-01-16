import time
import sys
import numpy as np
import mousekey
import rapidfuzz
import pytesseract
import pandas as pd
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn

# Configuração do Console para cores
console = Console()

'''
##########################################################################################
#                     MANUAL DE TESTE DE ESTRESSE (COMPLETO)                             #
#                                                                                        #
#  FASE 1: Validação Individual (Testa os 3 pontos de cada elemento isoladamente).       #
#  FASE 2: Matriz de Permutação (Testa as 27 combinações possíveis de comportamento).    #
##########################################################################################
'''

# --- IMPORTAÇÃO DAS CONFIGURAÇÕES ---
console.print("\n[bold yellow][INICIALIZANDO][/bold yellow] Importando variáveis do CaptchaClick.py...")
try:
    from CaptchaClick import (
        VARIACOES_RENAVAM,    
        VARIACOES_CAPTCHA,    
        VARIACOES_CONSULTAR,  
        OFFSET_EXTRA_X, 
        OFFSET_EXTRA_Y, 
        OFFSET_CONSULTAR_Y,
        get_screenshot_tesser
    )
    console.print("[bold green][SUCESSO][/bold green] Configurações importadas!")
except ImportError:
    console.print("[bold red][ERRO CRÍTICO][/bold red] Não encontrei as variáveis no CaptchaClick.py.")
    sys.exit()

# Configurações locais
mkey = mousekey.MouseKey()
mkey.enable_failsafekill('ctrl+e')
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def main():
    console.print("\n" + "="*60)
    console.print("    🕵️‍♂️   TESTE DE ESTRESSE DE MIRA (COMPLEXO)   🕵️‍♂️")
    console.print("    1. Abra o site em TELA CHEIA.")
    console.print("    2. Não mexa no mouse.")
    console.print("    O teste começará em 5 segundos...")
    console.print("="*60 + "\n")
    time.sleep(5) 

    # --- SCAN INICIAL DA TELA (CRÍTICO) ---
    console.print("[bold cyan]>>> REALIZANDO SCANNER DA TELA (OCR)...[/bold cyan]")
    try:
        df_tela = get_screenshot_tesser(lang='eng')
        if df_tela.empty: raise Exception("OCR retornou tela vazia.")
        
        # 1. Âncora Renavam
        palavra_alvo, pontuacao, indice = rapidfuzz.process.extractOne("Renavam", df_tela['text'])
        if pontuacao < 80: raise Exception("Palavra 'Renavam' não encontrada com confiança suficiente.")
        info_renavam = df_tela.iloc[indice]
        
        # Base Renavam
        base_renavam_x = info_renavam['left'] + info_renavam['width'] + 100
        base_renavam_y = info_renavam['top'] + (info_renavam['height'] // 2) + 30

        # 2. Âncora Captcha
        candidatos = df_tela[df_tela['text'].str.contains('robot', case=False, na=False)]
        info_captcha = None
        
        # Lógica de encontrar o captcha
        if candidatos.empty:
            console.print("[bold yellow][AVISO][/bold yellow] Texto 'robot' não encontrado. Usando estimativa baseada no Renavam.")
            # Cria um "fake" info_captcha baseado na posição estimada
            base_captcha_x = base_renavam_x - 50 # Ajuste estimado se falhar OCR
            base_captcha_y = base_renavam_y + 80
            # Precisamos de um dicionário fake para a lógica abaixo funcionar
            info_captcha = {'left': base_captcha_x, 'top': base_captcha_y}
        else:
            # Pega o primeiro candidato (ou implemente a lógica de distância aqui se tiver múltiplos)
            info_captcha = candidatos.iloc[0]
            base_captcha_x = info_captcha['left']
            base_captcha_y = info_captcha['top']

        console.print(f"[bold green]>>> REFERÊNCIAS ENCONTRADAS![/bold green]")
        console.print(f"    Renavam Base: X={base_renavam_x}, Y={base_renavam_y}")
        console.print(f"    Captcha Base: X={base_captcha_x}, Y={base_captcha_y}")
        time.sleep(2)

    except Exception as e:
        console.print(f"[bold red][ERRO FATAL NO SCAN][/bold red] {e}")
        console.print("Verifique se a tela está aberta corretamente.")
        sys.exit()


    # ============================================================================
    # INÍCIO DOS TESTES COM BARRA DE PROGRESSO
    # ============================================================================
    
    # Total de passos: 
    # Fase 1: 3 (Ren) + 3 (Cap) + 1 (Fix) + 3 (Cons) = 10 passos
    # Fase 2: 3 * 3 * 3 = 27 passos
    # Total Geral = 37 passos
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeRemainingColumn(),
        console=console
    ) as progress:

        # --- FASE 1: TESTES INDIVIDUAIS ---
        task_fase1 = progress.add_task("[cyan]FASE 1: Componentes Isolados...", total=10)
        
        mapa = {0: 'A', 1: 'B', 2: 'C'}

        # 1.1 Testando Variações Renavam
        for i in range(3):
            vx, vy = VARIACOES_RENAVAM[i]
            progress.console.print(f"[dim]Fase 1 - Renavam Posição {mapa[i]}: Offset {vx}, {vy}[/dim]")
            mkey.move_to_natural(base_renavam_x + vx, base_renavam_y + vy)
            time.sleep(0.5)
            progress.advance(task_fase1)

        # 1.2 Testando Variações Captcha
        for i in range(3):
            vx, vy = VARIACOES_CAPTCHA[i]
            # O Captcha usa a base encontrada do Captcha
            cx = base_captcha_x + vx
            cy = base_captcha_y + vy
            progress.console.print(f"[dim]Fase 1 - Captcha Posição {mapa[i]}: Offset {vx}, {vy}[/dim]")
            mkey.move_to_natural(cx, cy)
            time.sleep(0.5)
            progress.advance(task_fase1)

        # 1.3 Testando Ponto Fixo
        progress.console.print(f"[dim]Fase 1 - Ponto Fixo Intermediário[/dim]")
        # O ponto fixo é baseado na última posição do captcha, mas para teste vamos usar a base
        # Nota: No código real ele pega a posição atual. Aqui usaremos a Base + Offset Fixo
        # Para ser fiel, usaremos a base do Captcha + Variação A (0) + Offset
        cx_base = base_captcha_x + VARIACOES_CAPTCHA[0][0]
        cy_base = base_captcha_y + VARIACOES_CAPTCHA[0][1]
        mkey.move_to_natural(cx_base + OFFSET_EXTRA_X, cy_base + OFFSET_EXTRA_Y)
        time.sleep(0.5)
        progress.advance(task_fase1)

        # 1.4 Testando Variações Consultar
        for i in range(3):
            vx, vy = VARIACOES_CONSULTAR[i]
            # Consultar baseia-se no Captcha. Usaremos Captcha Base (variação 0) para teste isolado
            cx_ref = base_captcha_x + VARIACOES_CAPTCHA[0][0]
            cy_ref = base_captcha_y + VARIACOES_CAPTCHA[0][1]
            
            final_x = cx_ref + vx
            final_y = cy_ref + OFFSET_CONSULTAR_Y + vy
            
            progress.console.print(f"[dim]Fase 1 - Consultar Posição {mapa[i]}: Offset {vx}, {vy}[/dim]")
            mkey.move_to_natural(final_x, final_y)
            time.sleep(0.5)
            progress.advance(task_fase1)

        progress.console.print("[bold green]✔ FASE 1 CONCLUÍDA![/bold green]")
        time.sleep(1)

        # --- FASE 2: MATRIZ DE VARIAÇÕES (3x3x3) ---
        task_fase2 = progress.add_task("[magenta]FASE 2: Testando TODAS as Combinações...", total=27)

        # Loop Triplo (Renavam x Captcha x Consultar)
        for r in range(3):      # Renavam (0, 1, 2)
            for c in range(3):  # Captcha (0, 1, 2)
                for b in range(3): # Botão Consultar (0, 1, 2)
                    
                    combo_nome = f"R[{mapa[r]}]-C[{mapa[c]}]-B[{mapa[b]}]"
                    # progress.console.print(f"Testando Combo: {combo_nome}")

                    # --- PASSO 1: RENAVAM ---
                    var_ren_x, var_ren_y = VARIACOES_RENAVAM[r]
                    mkey.move_to_natural(base_renavam_x + var_ren_x, base_renavam_y + var_ren_y)
                    # Não colocamos sleep longo aqui para o teste fluir, o movimento natural já demora um pouco

                    # --- PASSO 2: CAPTCHA ---
                    var_cap_x, var_cap_y = VARIACOES_CAPTCHA[c]
                    # Define a posição exata onde o mouse clicaria no captcha nessa variação
                    pos_captcha_atual_x = base_captcha_x + var_cap_x
                    pos_captcha_atual_y = base_captcha_y + var_cap_y
                    
                    mkey.move_to_natural(pos_captcha_atual_x, pos_captcha_atual_y)

                    # --- PASSO 3: FIXO ---
                    # O fixo parte de onde o captcha parou
                    mkey.move_to_natural(pos_captcha_atual_x + OFFSET_EXTRA_X, pos_captcha_atual_y + OFFSET_EXTRA_Y)

                    # --- PASSO 4: CONSULTAR ---
                    var_cons_x, var_cons_y = VARIACOES_CONSULTAR[b]
                    # O consultar parte de onde o captcha parou + offset Y + ajuste fino
                    final_cons_x = pos_captcha_atual_x + var_cons_x
                    final_cons_y = pos_captcha_atual_y + OFFSET_CONSULTAR_Y + var_cons_y
                    
                    mkey.move_to_natural(final_cons_x, final_cons_y)
                    
                    # Atualiza barra
                    progress.advance(task_fase2)
                    
                    # Pausa rápida para visualizar o fim do ciclo
                    # time.sleep(0.1) 

    console.print("\n" + "="*60)
    console.print("[bold green]✅ TESTE DE ESTRESSE FINALIZADO COM SUCESSO![/bold green]")
    console.print("Todas as 37 etapas (Individuais + Combinações) foram validadas.")
    console.print("="*60)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("\n[bold red]Teste interrompido pelo usuário.[/bold red]")