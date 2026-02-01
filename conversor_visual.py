"""
Conversor Visual Excel -> HAP
Interface grafica simples para converter ficheiros Excel para HAP 5.1
"""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import sys
import os
import threading

class ConversorHAP:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor Excel -> HAP 5.1")
        self.root.geometry("700x500")
        self.root.resizable(True, True)

        # Directorio base
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.template = os.path.join(self.base_dir, "Template_Limpo_RSECE.E3A")

        # Variaveis
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.criar_interface()

    def criar_interface(self):
        # Frame principal
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Titulo
        titulo = tk.Label(main_frame, text="Conversor Excel -> HAP 5.1", font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))

        # Frame para ficheiro Excel
        frame_excel = tk.LabelFrame(main_frame, text="1. Ficheiro Excel", padx=10, pady=10)
        frame_excel.pack(fill=tk.X, pady=5)

        tk.Entry(frame_excel, textvariable=self.excel_path, width=60).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(frame_excel, text="Procurar...", command=self.procurar_excel).pack(side=tk.LEFT)

        # Frame para ficheiro de saida
        frame_output = tk.LabelFrame(main_frame, text="2. Ficheiro de Saida (.E3A)", padx=10, pady=10)
        frame_output.pack(fill=tk.X, pady=5)

        tk.Entry(frame_output, textvariable=self.output_path, width=60).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(frame_output, text="Guardar como...", command=self.escolher_output).pack(side=tk.LEFT)

        # Frame para botoes
        frame_botoes = tk.Frame(main_frame)
        frame_botoes.pack(fill=tk.X, pady=20)

        self.btn_validar = tk.Button(frame_botoes, text="Validar Excel", command=self.validar,
                                      width=15, height=2, bg="#4CAF50", fg="white")
        self.btn_validar.pack(side=tk.LEFT, padx=5)

        self.btn_converter = tk.Button(frame_botoes, text="Converter para HAP", command=self.converter,
                                        width=18, height=2, bg="#2196F3", fg="white")
        self.btn_converter.pack(side=tk.LEFT, padx=5)

        self.btn_validar_converter = tk.Button(frame_botoes, text="Validar + Converter", command=self.validar_e_converter,
                                                width=18, height=2, bg="#FF9800", fg="white")
        self.btn_validar_converter.pack(side=tk.LEFT, padx=5)

        # Area de log
        frame_log = tk.LabelFrame(main_frame, text="Resultado", padx=10, pady=10)
        frame_log.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log = scrolledtext.ScrolledText(frame_log, height=12, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

        # Barra de estado
        self.status = tk.Label(main_frame, text="Pronto", anchor=tk.W, relief=tk.SUNKEN)
        self.status.pack(fill=tk.X, pady=(10, 0))

    def procurar_excel(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar ficheiro Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~/Downloads")
        )
        if filepath:
            self.excel_path.set(filepath)
            # Sugerir nome de saida
            base_name = os.path.splitext(os.path.basename(filepath))[0]
            output_dir = os.path.dirname(filepath)
            self.output_path.set(os.path.join(output_dir, f"{base_name}.E3A"))

    def escolher_output(self):
        filepath = filedialog.asksaveasfilename(
            title="Guardar ficheiro HAP",
            filetypes=[("HAP files", "*.E3A"), ("All files", "*.*")],
            defaultextension=".E3A",
            initialdir=os.path.expanduser("~/Downloads")
        )
        if filepath:
            self.output_path.set(filepath)

    def log_clear(self):
        self.log.delete(1.0, tk.END)

    def log_write(self, text):
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.root.update()

    def set_status(self, text):
        self.status.config(text=text)
        self.root.update()

    def set_buttons_state(self, state):
        self.btn_validar.config(state=state)
        self.btn_converter.config(state=state)
        self.btn_validar_converter.config(state=state)

    def validar(self):
        if not self.excel_path.get():
            messagebox.showwarning("Aviso", "Seleccione um ficheiro Excel primeiro!")
            return False

        self.log_clear()
        self.set_status("A validar...")
        self.set_buttons_state(tk.DISABLED)

        try:
            validador = os.path.join(self.base_dir, "validar_excel_hap.py")
            result = subprocess.run(
                [sys.executable, validador, self.excel_path.get()],
                capture_output=True,
                text=True,
                cwd=self.base_dir,
                input="n\n"  # Responder nao ao relatorio detalhado
            )

            self.log_write(result.stdout)
            if result.stderr:
                self.log_write("ERROS:\n" + result.stderr)

            if "FICHEIRO VALIDO" in result.stdout or "FICHEIRO VÁLIDO" in result.stdout:
                self.set_status("Validacao concluida - VALIDO")
                self.set_buttons_state(tk.NORMAL)
                return True
            elif "ERROS:" in result.stdout and "ERROS: 0" not in result.stdout:
                self.set_status("Validacao concluida - COM ERROS")
                self.set_buttons_state(tk.NORMAL)
                return False
            else:
                self.set_status("Validacao concluida")
                self.set_buttons_state(tk.NORMAL)
                return True

        except Exception as e:
            self.log_write(f"ERRO: {str(e)}")
            self.set_status("Erro na validacao")
            self.set_buttons_state(tk.NORMAL)
            return False

    def converter(self):
        if not self.excel_path.get():
            messagebox.showwarning("Aviso", "Seleccione um ficheiro Excel primeiro!")
            return False

        if not self.output_path.get():
            messagebox.showwarning("Aviso", "Escolha onde guardar o ficheiro de saida!")
            return False

        if not os.path.exists(self.template):
            messagebox.showerror("Erro", f"Template nao encontrado:\n{self.template}")
            return False

        self.log_clear()
        self.set_status("A converter...")
        self.set_buttons_state(tk.DISABLED)

        try:
            conversor = os.path.join(self.base_dir, "excel_to_hap.py")
            result = subprocess.run(
                [sys.executable, conversor, self.excel_path.get(), self.template, self.output_path.get()],
                capture_output=True,
                text=True,
                cwd=self.base_dir
            )

            self.log_write(result.stdout)
            if result.stderr:
                self.log_write("ERROS:\n" + result.stderr)

            if result.returncode == 0 and os.path.exists(self.output_path.get()):
                self.set_status("Conversao concluida com sucesso!")
                self.log_write("\n" + "="*50)
                self.log_write(f"FICHEIRO CRIADO: {self.output_path.get()}")
                self.log_write("="*50)
                messagebox.showinfo("Sucesso", f"Ficheiro criado com sucesso!\n\n{self.output_path.get()}")
                self.set_buttons_state(tk.NORMAL)
                return True
            else:
                self.set_status("Erro na conversao")
                self.set_buttons_state(tk.NORMAL)
                return False

        except Exception as e:
            self.log_write(f"ERRO: {str(e)}")
            self.set_status("Erro na conversao")
            self.set_buttons_state(tk.NORMAL)
            return False

    def validar_e_converter(self):
        self.log_clear()
        self.log_write("="*50)
        self.log_write("PASSO 1: VALIDACAO")
        self.log_write("="*50 + "\n")

        if self.validar():
            self.log_write("\n" + "="*50)
            self.log_write("PASSO 2: CONVERSAO")
            self.log_write("="*50 + "\n")
            self.converter()
        else:
            resposta = messagebox.askyesno("Aviso",
                "A validacao encontrou problemas.\nDeseja continuar com a conversao mesmo assim?")
            if resposta:
                self.log_write("\n" + "="*50)
                self.log_write("PASSO 2: CONVERSAO (com avisos)")
                self.log_write("="*50 + "\n")
                self.converter()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ConversorHAP()
    app.run()
