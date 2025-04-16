import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from modules.db_handler import obtener_bancos, obtener_referencias
from modules.file_loader import cargar_excel
from modules.painter import pintar_extracto, pintar_contabilidad
from modules.discrepancy import generar_discrepancias


def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)  # Ruta del .exe
    return os.path.dirname(os.path.abspath(__file__))  # Ruta del .py

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Conciliador Contable")

        self.db_path = os.path.join(obtener_ruta_base(), "DB_Contabilidad (1).db")
        self.bancos = obtener_bancos(self.db_path)
        self.banco = tk.StringVar()
        self.extracto_path = None
        self.contabilidad_path = None

        self.setup_gui()

    def setup_gui(self):
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack()

        # Lista desplegable de bancos
        tk.Label(frame, text="Banco:").grid(row=0, column=0, sticky="e")
        self.dropdown_bancos = ttk.Combobox(frame, textvariable=self.banco, values=self.bancos, state="readonly")
        self.dropdown_bancos.grid(row=0, column=1, pady=5, sticky="ew")
        if self.bancos:
            self.banco.set(self.bancos[0])

        # Selección de archivos
        tk.Button(frame, text="Cargar Extracto", command=self.cargar_extracto).grid(row=1, column=0, pady=5, sticky="ew")
        self.label_ext = tk.Label(frame, text="Archivo no cargado")
        self.label_ext.grid(row=1, column=1, padx=5)

        tk.Button(frame, text="Cargar Contabilidad", command=self.cargar_contabilidad).grid(row=2, column=0, pady=5, sticky="ew")
        self.label_cont = tk.Label(frame, text="Archivo no cargado")
        self.label_cont.grid(row=2, column=1, padx=5)

        # Botón para procesar
        tk.Button(frame, text="Procesar Archivos", command=self.procesar).grid(row=3, column=0, columnspan=2, pady=10)

    def cargar_extracto(self):
        self.extracto_path = filedialog.askopenfilename(title="Seleccionar archivo de Extracto", filetypes=[("Excel files", "*.xlsx")])
        if self.extracto_path:
            self.label_ext.config(text=os.path.basename(self.extracto_path))

    def cargar_contabilidad(self):
        self.contabilidad_path = filedialog.askopenfilename(title="Seleccionar archivo de Contabilidad", filetypes=[("Excel files", "*.xlsx")])
        if self.contabilidad_path:
            self.label_cont.config(text=os.path.basename(self.contabilidad_path))

    def procesar(self):
        if not all([self.extracto_path, self.contabilidad_path, self.banco.get()]):
            messagebox.showerror("Error", "Debe seleccionar todos los archivos y el banco.")
            return

        try:
            referencias = obtener_referencias(self.db_path, self.banco.get())
            df_extracto = cargar_excel(self.extracto_path)
            df_contabilidad = cargar_excel(self.contabilidad_path)

            df_extracto_coloreado = pintar_extracto(df_extracto, referencias)

            df_contabilidad_coloreado = pintar_contabilidad(df_contabilidad, referencias)

            output_discrepancias = os.path.join(os.path.dirname(self.extracto_path), "Discrepancias.xlsx")
            generar_discrepancias(df_extracto_coloreado, df_contabilidad_coloreado, output_discrepancias)

            messagebox.showinfo("Éxito", f"Archivos procesados correctamente.\nDiscrepancias guardadas en:\n{output_discrepancias}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
