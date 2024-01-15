import pyautogui
import time

def main():
    while True:
        # Espera 60 segundos
        time.sleep(60)
        
        # Simula o clique na tecla "a"
        pyautogui.press('a')

if __name__ == "__main__":
    main()
