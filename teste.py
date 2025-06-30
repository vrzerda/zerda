import pyautogui
import time
import random
import math
import sys
import platform
import threading
from datetime import datetime

class ContinuousMouseMover:
    def __init__(self):
        self.running = False
        self.center_x = 0
        self.center_y = 0
        self.radius = 50
        self.angle = 0
        self.movement_type = "circle"
        self.speed = 0.1
        self.action_count = 0
        
    def print_status(self, message):
        """Imprime status com timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")

    def prevent_sleep_windows(self):
        """Previne suspensão no Windows usando API nativa"""
        try:
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001 | 0x00000040)
            return True
        except:
            return False

    def setup_center_position(self):
        """Define a posição central para os movimentos"""
        screen_width, screen_height = pyautogui.size()
        self.center_x = screen_width // 2
        self.center_y = screen_height // 2
        self.print_status(f"Centro definido em ({self.center_x}, {self.center_y})")

    def circular_movement(self):
        """Movimento circular contínuo"""
        self.angle += self.speed
        if self.angle >= 2 * math.pi:
            self.angle = 0
            
        x = self.center_x + int(self.radius * math.cos(self.angle))
        y = self.center_y + int(self.radius * math.sin(self.angle))
        
        return x, y

    def figure_eight_movement(self):
        """Movimento em formato de 8"""
        self.angle += self.speed
        if self.angle >= 4 * math.pi:
            self.angle = 0
            
        x = self.center_x + int(self.radius * math.sin(self.angle))
        y = self.center_y + int(self.radius * math.sin(self.angle) * math.cos(self.angle))
        
        return x, y

    def random_smooth_movement(self):
        """Movimento aleatório suave"""
        current_x, current_y = pyautogui.position()
        
        # Gera movimento aleatório mas suave
        max_distance = 30
        angle_rand = random.uniform(0, 2 * math.pi)
        distance = random.uniform(5, max_distance)
        
        new_x = current_x + int(distance * math.cos(angle_rand))
        new_y = current_y + int(distance * math.sin(angle_rand))
        
        # Mantém dentro dos limites da tela
        screen_width, screen_height = pyautogui.size()
        new_x = max(50, min(new_x, screen_width - 50))
        new_y = max(50, min(new_y, screen_height - 50))
        
        return new_x, new_y

    def get_next_position(self):
        """Retorna próxima posição baseada no tipo de movimento"""
        if self.movement_type == "circle":
            return self.circular_movement()
        elif self.movement_type == "figure8":
            return self.figure_eight_movement()
        else:  # random
            return self.random_smooth_movement()

    def continuous_movement_thread(self):
        """Thread para movimento contínuo do mouse"""
        self.print_status(f"Iniciando movimento contínuo tipo: {self.movement_type}")
        
        while self.running:
            try:
                # Obtém próxima posição
                x, y = self.get_next_position()
                
                # Move o mouse suavemente
                pyautogui.moveTo(x, y, duration=0.01)
                
                self.action_count += 1
                
                # Log a cada 100 movimentos para não spam
                if self.action_count % 100 == 0:
                    self.print_status(f"Movimento #{self.action_count} - Posição ({x}, {y})")
                
                # Pequena pausa entre movimentos
                time.sleep(0.02)  # 50 movimentos por segundo
                
            except Exception as e:
                self.print_status(f"Erro no movimento: {e}")
                time.sleep(0.1)

    def keyboard_activity_thread(self):
        """Thread para atividade de teclado periódica - teclas modificadoras invisíveis"""
        keyboard_actions = [
            'shift',      # Shift (não faz nada visível sozinho)
            'ctrl',       # Ctrl (não faz nada visível sozinho)
            'alt',        # Alt (não faz nada visível sozinho)
            'win',        # Windows key (não abre menu se pressionado rapidamente)
            'capslock',   # Caps Lock (liga/desliga mas volta ao normal)
            'numlock',    # Num Lock (liga/desliga mas volta ao normal)
            'scrolllock', # Scroll Lock (não afeta nada moderno)
            'pause',      # Pause/Break (não faz nada na maioria dos casos)
            'insert',     # Insert (modo inserir/sobrescrever)
            'home',       # Home (só move cursor se houver foco em texto)
            'end',        # End (só move cursor se houver foco em texto)
            'pageup',     # Page Up (só funciona se houver foco)
            'pagedown',   # Page Down (só funciona se houver foco)
            'up', 'down', 'left', 'right'  # Setas (só funcionam com foco)
        ]
        
        action_index = 0
        
        while self.running:
            try:
                # Seleciona ação de teclado ciclicamente
                key = keyboard_actions[action_index % len(keyboard_actions)]
                
                # Para teclas que podem alterar estado, pressiona duas vezes
                if key in ['capslock', 'numlock', 'scrolllock', 'insert']:
                    pyautogui.press(key)
                    time.sleep(0.05)
                    pyautogui.press(key)  # Volta ao estado original
                    activity_type = f"{key} (duplo)"
                else:
                    # Para outras teclas, apenas pressiona uma vez
                    pyautogui.press(key)
                    activity_type = key
                
                action_index += 1
                
                # Log da atividade de teclado
                if action_index % 20 == 0:  # Log a cada 20 ações
                    self.print_status(f"Atividade de teclado: {activity_type} (#{action_index})")
                
                # Aguarda entre 2-4 segundos de forma aleatória
                wait_time = random.uniform(2, 4)
                time.sleep(wait_time)
                
            except Exception as e:
                self.print_status(f"Erro na atividade de teclado: {e}")
                time.sleep(1)

    def start_continuous_movement(self, movement_type="circle"):
        """Inicia movimento contínuo"""
        self.movement_type = movement_type
        self.running = True
        
        # Configura posição central
        self.setup_center_position()
        
        # Prevenção de suspensão (Windows)
        if platform.system() == "Windows":
            if self.prevent_sleep_windows():
                self.print_status("API de prevenção de suspensão ativada")
        
        # Inicia threads
        mouse_thread = threading.Thread(target=self.continuous_movement_thread, daemon=True)
        keyboard_thread = threading.Thread(target=self.keyboard_activity_thread, daemon=True)
        
        mouse_thread.start()
        keyboard_thread.start()
        
        self.print_status("Movimento contínuo iniciado!")
        return mouse_thread, keyboard_thread

    def stop_movement(self):
        """Para o movimento contínuo"""
        self.running = False
        
        # Restaura estado normal (Windows)
        if platform.system() == "Windows":
            try:
                import ctypes
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
                self.print_status("Estado de suspensão restaurado")
            except:
                pass

def main():
    """Função principal"""
    print("=" * 70)
    print("    MOVEDOR DE MOUSE CONTÍNUO - ANTI-SUSPENSÃO AVANÇADO")
    print("=" * 70)
    print("Este programa move o mouse continuamente em padrões suaves:")
    print("• Movimento circular constante")
    print("• Movimento em formato de 8") 
    print("• Movimento aleatório suave")
    print("• Atividade diversificada de teclado (teclas invisíveis)")
    print("• Prevenção de suspensão via API")
    print("\nPressione Ctrl+C para parar o programa.")
    print("=" * 70)
    
    # Escolha do tipo de movimento
    print("\nTipos de movimento disponíveis:")
    print("1. Circular (padrão)")
    print("2. Formato de 8")
    print("3. Aleatório suave")
    
    try:
        choice = input("Escolha o tipo (1/2/3) ou Enter para circular: ").strip()
        if choice == "2":
            movement_type = "figure8"
        elif choice == "3":
            movement_type = "random"
        else:
            movement_type = "circle"
    except:
        movement_type = "circle"
    
    print(f"\nTipo de movimento selecionado: {movement_type}")
    print(f"Sistema operacional: {platform.system()}")
    print(f"Resolução da tela: {pyautogui.size()}")
    
    # Configurações do pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.001  # Pausa mínima para performance
    
    # Cria e inicia o movedor
    mover = ContinuousMouseMover()
    
    try:
        threads = mover.start_continuous_movement(movement_type)
        
        # Mantém programa rodando
        while True:
            time.sleep(1)
            
    except KeyboardInterrupt:
        mover.print_status("Interrupção solicitada pelo usuário...")
        mover.stop_movement()
        time.sleep(0.5)  # Aguarda threads terminarem
        print("Programa finalizado com segurança!")
        
    except pyautogui.FailSafeException:
        mover.print_status("FailSafe ativado - mouse nos cantos da tela")
        mover.stop_movement()
        print("Programa finalizado por segurança!")
        
    except Exception as e:
        mover.print_status(f"Erro inesperado: {e}")
        mover.stop_movement()
        print("Programa finalizado!")

if __name__ == "__main__":
    main()