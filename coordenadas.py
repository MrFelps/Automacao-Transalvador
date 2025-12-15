import mousekey
import time

mkey = mousekey.MouseKey()
print("Movimente o mouse e veja as coordenadas. Pressione Ctrl+C no terminal para parar.")
try:
    while True:
        x, y = mkey.get_cursor_position()
        # O 'end='\r'' faz a linha ser atualizada em vez de criar novas linhas
        print(f"X: {x}, Y: {y}    ", end='\r')
        time.sleep(0.1)
except KeyboardInterrupt:
    print("\n\nPrograma finalizado.")