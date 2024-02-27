import pyautogui
import keyboard
import time

def perform_click():
    pyautogui.click()

# Establecer teclas para iniciar y detener el bot
start_key = 'f1'
stop_key = 'f2'

print(f"Presiona {start_key} para iniciar el clic, {stop_key} para detener.")

running = True
clicking = False

while running:
    try:
        if keyboard.is_pressed(start_key):
            clicking = True
        elif keyboard.is_pressed(stop_key):
            running = False

        if clicking:
            perform_click()
            time.sleep(0.1)  # Ajusta este valor para controlar la velocidad de clic

    except Exception as e:
        print(f"Ocurrió un error: {e}")
        running = False