import tkinter as tk
from tkinter import ttk, messagebox
import psycopg2
import openpyxl
from openpyxl.styles import Font
import json
import os
from cryptography.fernet import Fernet

class ConsultaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Consulta Factura")
        self.geometry("400x500")

        self.host_label = ttk.Label(self, text="Host:")
        self.host_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.host_combobox = ttk.Combobox(self)
        self.host_combobox.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        self.port_label = ttk.Label(self, text="Puerto:")
        self.port_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.port_combobox = ttk.Combobox(self)
        self.port_combobox.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        self.database_label = ttk.Label(self, text="Base de datos:")
        self.database_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.database_combobox = ttk.Combobox(self)
        self.database_combobox.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        self.usuario_label = ttk.Label(self, text="Usuario:")
        self.usuario_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.usuario_combobox = ttk.Combobox(self)
        self.usuario_combobox.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        self.password_label = ttk.Label(self, text="Contraseña:")
        self.password_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.password_combobox = ttk.Combobox(self, show="*")
        self.password_combobox.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        self.start_date_label = ttk.Label(self, text="Fecha de Inicio (YYYY-MM-DD):")
        self.start_date_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.start_date_combobox = ttk.Combobox(self)
        self.start_date_combobox.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        self.end_date_label = ttk.Label(self, text="Fecha de Fin (YYYY-MM-DD):")
        self.end_date_label.grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.end_date_combobox = ttk.Combobox(self)
        self.end_date_combobox.grid(row=6, column=1, padx=10, pady=5, sticky="ew")

        self.archivo_label = ttk.Label(self, text="Nombre del archivo:")
        self.archivo_label.grid(row=7, column=0, padx=10, pady=5, sticky="w")
        self.archivo_combobox = ttk.Combobox(self)
        self.archivo_combobox.grid(row=7, column=1, padx=10, pady=5, sticky="ew")

        self.ejecutar_button = ttk.Button(self, text="Realizar Consulta", command=self.realizar_consulta)
        self.ejecutar_button.grid(row=8, column=1, padx=10, pady=10, sticky="ew")

        self.grid_columnconfigure(1, weight=1)

        self.config_file = "config.json"
        self.key_file = "key.key"
        self.load_key()
        self.load_config()

    def load_key(self):
        if os.path.exists(self.key_file):
            with open(self.key_file, "rb") as file:
                self.key = file.read()
        else:
            self.key = Fernet.generate_key()
            with open(self.key_file, "wb") as file:
                file.write(self.key)
        self.cipher = Fernet(self.key)
    
    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "rb") as file:
                encrypted_data = file.read()
                decrypted_data = self.cipher.decrypt(encrypted_data)
                config = json.loads(decrypted_data.decode())
                self.host_combobox['values'] = config.get("host", [])
                self.port_combobox['values'] = config.get("port", [])
                self.database_combobox['values'] = config.get("database", [])
                self.usuario_combobox['values'] = config.get("usuario", [])
                self.password_combobox['values'] = config.get("password", [])
                self.start_date_combobox['values'] = config.get("start_date", [])
                self.end_date_combobox['values'] = config.get("end_date", [])
                self.archivo_combobox['values'] = config.get("archivo", [])

                # Establecer el valor actual de cada Combobox
                self.host_combobox.set(config.get("host", [""])[-1])
                self.port_combobox.set(config.get("port", [""])[-1])
                self.database_combobox.set(config.get("database", [""])[-1])
                self.usuario_combobox.set(config.get("usuario", [""])[-1])
                self.password_combobox.set(config.get("password", [""])[-1])
                self.start_date_combobox.set(config.get("start_date", [""])[-1])
                self.end_date_combobox.set(config.get("end_date", [""])[-1])
                self.archivo_combobox.set(config.get("archivo", [""])[-1])

    def save_config(self):
        config = {
            "host": list(set(list(self.host_combobox['values']) + [self.host_combobox.get()])),
            "port": list(set(list(self.port_combobox['values']) + [self.port_combobox.get()])),
            "database": list(set(list(self.database_combobox['values']) + [self.database_combobox.get()])),
            "usuario": list(set(list(self.usuario_combobox['values']) + [self.usuario_combobox.get()])),
            "password": list(set(list(self.password_combobox['values']) + [self.password_combobox.get()])),
            "start_date": list(set(list(self.start_date_combobox['values']) + [self.start_date_combobox.get()])),
            "end_date": list(set(list(self.end_date_combobox['values']) + [self.end_date_combobox.get()])),
            "archivo": list(set(list(self.archivo_combobox['values']) + [self.archivo_combobox.get()]))
        }
        config_data = json.dumps(config).encode()
        encrypted_data = self.cipher.encrypt(config_data)
        with open(self.config_file, "wb") as file:
            file.write(encrypted_data)
    
    def realizar_consulta(self):
        host = self.host_combobox.get()
        port = self.port_combobox.get()
        database = self.database_combobox.get()
        usuario = self.usuario_combobox.get()
        password = self.password_combobox.get()
        start_date = self.start_date_combobox.get()
        end_date = self.end_date_combobox.get()
        nombre_archivo = self.archivo_combobox.get()

        self.save_config()

        consulta = f"""
        SELECT 
           ...(Consulta omitida por integridad de datos)...
        """

        try:
            conn = psycopg2.connect(
                host=host,
                port=port,
                database=database,
                user=usuario,
                password=password
            )
            cursor = conn.cursor()
            cursor.execute(consulta)
            resultados = cursor.fetchall()

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Resultados"

            headers = [desc[0] for desc in cursor.description]
            sheet.append(headers)
            for cell in sheet["1:1"]:
                cell.font = Font(bold=True)

            for row in resultados:
                row_data = []
                for value, desc in zip(row, cursor.description):
                    if desc[1] == psycopg2.NUMBER:
                        row_data.append(value)
                    elif desc[1] == psycopg2.STRING:
                        row_data.append(value)
                    elif desc[1] == psycopg2.DATETIME:
                        row_data.append(value)
                    else:
                        row_data.append(value)
                sheet.append(row_data)

            workbook.save(f"{nombre_archivo}.xlsx")
            messagebox.showinfo("Éxito", f"Datos guardados en {nombre_archivo}.xlsx")

            cursor.close()
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = ConsultaApp()
    app.mainloop()
