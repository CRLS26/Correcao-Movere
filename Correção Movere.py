import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import psutil
import ctypes
import sys
import winreg
import time
import shutil
import threading

icone = os.path.join(os.path.dirname(__file__), 'Correção Movere.ico')

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

def fechar_edge():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'msedge.exe':
            proc.terminate()
    time.sleep(2)

def limpar_cookies_edge():
    if not is_admin():
        run_as_admin()
        sys.exit()
    fechar_edge()
    edge_user_data_path = os.path.join(os.getenv('LOCALAPPDATA'), 'Microsoft', 'Edge', 'User Data')
    perfis = ['Default', 'Profile 1', 'Profile 2', 'Profile 3']
    for perfil in perfis:
        cookies_path = os.path.join(edge_user_data_path, perfil, 'Network')
        if os.path.exists(cookies_path):
            try:
                shutil.rmtree(cookies_path)
                os.mkdir(cookies_path)
            except Exception:
                fechar_edge()
                try:
                    shutil.rmtree(cookies_path)
                    os.mkdir(cookies_path)
                except Exception:
                    pass

def executar_programa():
    programa = os.path.join(os.getenv('USERPROFILE'), 'Desktop', 'ConfigurarMeuMovere.exe')
    if os.path.exists(programa):
        subprocess.Popen(programa)
    else:
        messagebox.showerror("Erro", "O configurador movere não foi encontrado!")

def redefinir_pagina_inicial_edge():
    try:
        caminho_registro = r"Software\Policies\Microsoft\Edge"
        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, caminho_registro)
        nova_pagina_inicial = "https://app1.moveresoftware.com/podiumpneus/"
        winreg.SetValueEx(key, "HomepageLocation", 0, winreg.REG_SZ, nova_pagina_inicial)
        winreg.SetValueEx(key, "HomepageIsNewTabPage", 0, winreg.REG_DWORD, 0)
        winreg.CloseKey(key)  
        return True 
    except Exception:
        return False  

def apagar_pasta_temp():
    pasta = os.path.join(os.getenv('USERPROFILE'), 'Temp')
    if os.path.exists(pasta):
        try:
            shutil.rmtree(pasta)
        except Exception as e:
            pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        if not is_admin():
            run_as_admin()
            sys.exit()
        self.title("Correção Movere")
        self.geometry("600x700")
        self.configure(bg='#f0f0f0')
        self.iconbitmap(icone)
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        title = ttk.Label(main_frame, 
                         text="Correção Movere", 
                         font=('Helvetica', 16, 'bold'))
        title.pack(pady=10)

        ttk.Label(main_frame, 
                 text="Selecione o erro que está enfrentando:", 
                 font=("Helvetica", 14)).pack(pady=10)

        options_frame = ttk.LabelFrame(main_frame, text="Opções de Correção")
        options_frame.pack(fill='x', pady=5)

        self.opcao_var = tk.StringVar()
        ttk.Radiobutton(options_frame, 
                       text="Senha de acesso bloqueada", 
                       variable=self.opcao_var, 
                       value="1").pack(anchor='w', padx=5, pady=2)
        
        ttk.Radiobutton(options_frame, 
                       text="Erro ao executar rotinas", 
                       variable=self.opcao_var, 
                       value="2").pack(anchor='w', padx=5, pady=2)

        self.log_text = tk.Text(main_frame, height=15, width=50)
        self.log_text.pack(pady=10)

        self.progress = ttk.Progressbar(main_frame, 
                                      mode='determinate',
                                      length=400)
        self.progress.pack(pady=10)

        self.execute_button = ttk.Button(main_frame,
                                       text="Executar",
                                       command=self.start_cleaning,
                                       padding=(20, 10))
        self.execute_button.pack(pady=10)

        copyright_label = ttk.Label(main_frame, 
                                  text="© 2025 Carlos Teixeira - Todos os direitos reservados", 
                                  font=("Arial", 8))
        copyright_label.pack(side="bottom", pady=5)

    def start_cleaning(self):
        def run_cleaning():
            try:
                self.execute_button.configure(state='disabled')
                opcao = self.opcao_var.get()
                
                if not opcao:
                    messagebox.showwarning("Aviso", "Selecione uma opção!")
                    self.execute_button.configure(state='normal')
                    return

                total_steps = 4
                current_step = 0

                self.log_text.insert(tk.END, "Iniciando processo de correção...\n")
                
                if opcao == "1":
                    self.log_text.insert(tk.END, "Limpando cookies do Edge...\n")
                    limpar_cookies_edge()
                    current_step += 1
                    self.progress['value'] = (current_step / total_steps) * 100
                    self.update_idletasks()
                elif opcao == "2":
                    self.log_text.insert(tk.END, "Redefinindo página inicial...\n")
                    redefinir_pagina_inicial_edge()
                    current_step += 1
                    self.progress['value'] = (current_step / total_steps) * 100
                    self.update_idletasks()

                    self.log_text.insert(tk.END, "Executando configurador...\n")
                    executar_programa()
                    current_step += 1
                    self.progress['value'] = (current_step / total_steps) * 100
                    self.update_idletasks()

                self.log_text.insert(tk.END, "Aguardando a conclusão do configurador...\n")
                time.sleep(30)
                current_step += 1
                self.progress['value'] = (current_step / total_steps) * 100
                self.update_idletasks()

                self.log_text.insert(tk.END, "\nProcesso concluído com sucesso!\n")
                self.log_text.insert(tk.END, "\nMovere estará pronto para uso, após a inicialização do navegador!\n")
                self.log_text.see(tk.END)

            except Exception as e:
                self.log_text.insert(tk.END, f"\nErro durante o processo: {str(e)}\n")
            finally:
                self.execute_button.configure(state='normal')
                self.log_text.see(tk.END)

        threading.Thread(target=run_cleaning, daemon=True).start()

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()