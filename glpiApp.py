import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import requests
import json
from datetime import datetime
import os


class GLPIDeviceImporter:
    def __init__(self, root):
        self.root = root
        self.root.title("Importador de Dispositivos GLPI")
        icon = tk.PhotoImage(file="logo.png")
        self.root.iconphoto(True, icon)
        self.root.geometry("900x700")

        # Variables de control
        self.excel_path = tk.StringVar()
        self.base_url = tk.StringVar(value="http://10.14.35.145")
        # print("base url: ", self.base_url.get())
        self.app_token = tk.StringVar(value="l1TqW2dIF38lxSPZstMzFFtDkrqzptV2H5BhdRxh")
        self.user_token = tk.StringVar(value="T99P59LLi8n1DIzTQBFT9uFt19kNQ0w4jiwsLgOH")
        self.headers = {"App-Token": self.app_token.get()}
        self.session_token = None
        # Crear el marco principal
        self.create_main_frame()
        # Variables para almacenar datos
        self.excel_data = None
        self.log_messages = []

    def iniciar_sesion(self):
        """
        Inicia sesión en GLPI usando user_token
        """
        headers = {
            "Content-Type": "application/json",
            "App-Token": self.app_token.get(),
            "Authorization": f"user_token {self.user_token.get()}",
        }
        print("iniciar sesion: ", headers)
        response = requests.post(
            f"{self.base_url.get()}/apirest.php/initSession", headers=headers
        )
        if response.status_code == 200:
            self.session_token = str(response.json()["session_token"])
            print("session token generated", self.session_token)
            self.headers.clear()
            self.headers = {
                "Session-token": self.session_token,
                "App-Token": self.app_token.get(),
            }
            return True
        raise Exception(f"Error al iniciar sesión: {response.text}")

    def create_main_frame(self):
        # Marco de configuración
        config_frame = ttk.LabelFrame(self.root, text="Configuración", padding="10")
        config_frame.pack(fill="x", padx=10, pady=5)

        # URL
        ttk.Label(config_frame, text="URL GLPI:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        ttk.Entry(config_frame, textvariable=self.base_url, width=50).grid(
            row=0, column=1, padx=5
        )

        # App Token
        ttk.Label(config_frame, text="App Token:").grid(
            row=1, column=0, sticky="w", pady=5
        )
        ttk.Entry(config_frame, textvariable=self.app_token, width=50, show="*").grid(
            row=1, column=1, padx=5
        )

        # User Token
        ttk.Label(config_frame, text="User Token:").grid(
            row=2, column=0, sticky="w", pady=5
        )
        ttk.Entry(config_frame, textvariable=self.user_token, width=50, show="*").grid(
            row=2, column=1, padx=5
        )

        # Marco de selección de archivo
        file_frame = ttk.LabelFrame(self.root, text="Archivo Excel", padding="10")
        file_frame.pack(fill="x", padx=10, pady=5)

        ttk.Entry(file_frame, textvariable=self.excel_path, width=70).pack(
            side="left", padx=5
        )
        ttk.Button(file_frame, text="Examinar", command=self.browse_file).pack(
            side="left", padx=5
        )

        # Marco de vista previa
        preview_frame = ttk.LabelFrame(self.root, text="Vista Previa", padding="10")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Crear Treeview para vista previa
        self.tree = ttk.Treeview(preview_frame, selectmode="browse")
        self.tree.pack(fill="both", expand=True)

        # Scrollbar para Treeview
        scrollbar = ttk.Scrollbar(
            preview_frame, orient="vertical", command=self.tree.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Marco de botones
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(button_frame, text="Cargar Excel", command=self.load_excel).pack(
            side="left", padx=5
        )
        ttk.Button(button_frame, text="Validar Datos", command=self.validate_data).pack(
            side="left", padx=5
        )
        ttk.Button(
            button_frame, text="Importar a GLPI", command=self.import_to_glpi
        ).pack(side="left", padx=5)

        # Marco de log
        log_frame = ttk.LabelFrame(self.root, text="Log", padding="10")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=6)
        self.log_text.pack(fill="both", expand=True)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if filename:
            self.excel_path.set(filename)

    def load_excel(self):
        if not self.excel_path.get():
            messagebox.showerror("Error", "Por favor seleccione un archivo Excel")
            return

        try:
            self.excel_data = pd.read_excel(self.excel_path.get())

            # Limpiar Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Configurar columnas
            self.tree["columns"] = list(self.excel_data.columns)
            self.tree["show"] = "headings"

            for column in self.excel_data.columns:
                self.tree.heading(column, text=column)
                self.tree.column(column, width=100)

            # Insertar datos
            for idx, row in self.excel_data.iterrows():
                self.tree.insert("", "end", values=list(row))

            self.log_message("Excel cargado exitosamente")

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el Excel: {str(e)}")
            self.log_message(f"Error: {str(e)}")

    def validate_data(self):
        if self.excel_data is None:
            messagebox.showerror("Error", "Por favor cargue un archivo Excel primero")
            return

        required_columns = [
            "id",
            "marbete",
            "tipo",
            "fabricante",
            "modelo",
            "serie",
            "fecha_de_compra",
            "fecha_de_inicio",
            "fecha_de_puesta_en_marcha",
            "proveedor",
            "monto",
            "fecha_de_inicio_garantia",
            "duracion_de_garantia",
            "ubicacion",
            "nombre",
        ]  # Ajusta según tus necesidades
        missing_columns = [
            col for col in required_columns if col not in self.excel_data.columns
        ]

        if missing_columns:
            messagebox.showerror(
                "Error", f"Faltan columnas requeridas: {', '.join(missing_columns)}"
            )
            return

        # Validar datos
        validation_errors = []

        for idx, row in self.excel_data.iterrows():
            for col_name in required_columns:
                if pd.isna(row[col_name]):
                    validation_errors.append(f"Fila {idx + 2}: {col_name} vacío")

        if validation_errors:
            error_message = "\n".join(validation_errors)
            messagebox.showerror("Errores de Validación", error_message)
            self.log_message("Validación fallida")
        else:
            messagebox.showinfo("Éxito", "Validación exitosa")
            self.log_message("Validación exitosa")

    def crear_dispositivo(self, tipo_dispositivo, input_data):
        """
        Crea un nuevo dispositivo en GLPI

        Args:
            tipo_dispositivo (str): Tipo de dispositivo (Computer, Monitor, Printer, etc.)
            datos (dict): Datos del dispositivo
        """
        if not self.session_token:
            raise Exception("No hay sesión iniciada")

        headers = {
            "Session-Token": self.session_token,
            "App-Token": self.app_token,
            "Content-Type": "application/json",
        }
        data = {}
        for key, value in input_data.items():
            if isinstance(value, (int, bool)):
                data[key] = value
            else:
                data[key] = str(value)

        # Crear la estructura correcta para la API
        input_array = {"input": data}

        print("Enviando datos:", json.dumps(input_array, indent=2))

        response = requests.post(
            f"{self.base_url}/apirest.php/{tipo_dispositivo}",
            headers=headers,
            json=input_array,
        )
        print(response.content)

        if response.status_code in [200, 201]:
            return response.json()
        raise Exception(f"Error al crear dispositivo: {response.text}")

    def buscar_modelos(self, search, tipo_dispositivo):
        print("criterio de busqueda:", search)
        search_criteria = {}
        search_criteria["model"] = search

        model_endpoints = {
            "computer": "ComputerModel",
            "printer": "PrinterModel",
            "monitor": "MonitorModel",
            "network": "NetworkEquipmentModel",
            "peripheral": "PeripheralModel",
            "phone": "PhoneModel",
        }
        # Mapeo de tipos de dispositivos a sus endpoints de creación
        device_endpoints = {
            "computer": "Computer",
            "printer": "Printer",
            "monitor": "Monitor",
            "network": "NetworkEquipment",
            "peripheral": "Peripheral",
            "phone": "Phone",
        }

        if not self.session_token:
            raise Exception("No hay sesión iniciada")
        else:
            self.headers["Session_token"] = self.session_token

        if tipo_dispositivo.lower() not in model_endpoints:
            return {
                "exists": False,
                "model_id": None,
                "message": f'Tipo de dispositivo no válido. Opciones válidas: {", ".join(model_endpoints.keys())}',
            }

        # Construir los parámetros de búsqueda
        search_params = {"criteria[0][link]": "AND"}

        # Mapeo de campos para la búsqueda
        field_mapping = {
            "name": 1,  # Nombre
            "serial": 45,  # Número de serie
            "id": 2,  # ID del dispositivo
            "model": 40,  # Modelo
            "location": 3,  # Ubicación
            "user": 24,  # Usuario
            "manufacturer": 23,  # Fabricante
            "comment": 16,  # Comentarios
            "status": 31,  # Estado
            "type": 4,  # Tipo
            "inventory_number": 46,  # Número de inventario
        }
        # Agregar criterios de búsqueda a los parámetros
        criterion_index = 0
        for field, value in search_criteria.items():
            if field in field_mapping:
                search_params[f"criteria[{criterion_index}][field]"] = field_mapping[
                    field
                ]
                search_params[f"criteria[{criterion_index}][searchtype]"] = "contains"
                search_params[f"criteria[{criterion_index}][value]"] = value
                criterion_index += 1

        try:
            endpoint = model_endpoints[tipo_dispositivo.lower()]
            self.log_message(f"Buscando modelo de {tipo_dispositivo}: {search}")

            response = requests.get(
                f"{self.base_url.get()}/apirest.php/search/Computer/",
                headers=self.headers,
                params=search_params,
                verify=False,
            )

            if response.status_code == 200:
                modelos = response.json()
                # Verificar si se encontraron resultados
                if isinstance(modelos, list) and len(modelos) > 0:
                    # Buscar coincidencia exacta
                    for modelo in modelos:
                        if modelo.get("name", "").lower() == search.lower():
                            return {
                                "exists": True,
                                "model_id": modelo.get("id"),
                                "name": modelo.get("name"),
                                "message": "Modelo encontrado",
                            }

                    # Si no hay coincidencia exacta, retornar el primer resultado
                    return {
                        "exists": True,
                        "model_id": modelos[0].get("id"),
                        "name": modelos[0].get("name"),
                        "message": "Modelo similar encontrado",
                    }
                else:
                    return {
                        "exists": False,
                        "model_id": None,
                        "message": "Modelo no encontrado",
                    }
            else:
                return {
                    "exists": False,
                    "count": 0,
                    "data": [],
                    "message": f"Error en la búsqueda: {response.status_code}",
                }, "datos incorrectos..."

        except requests.exceptions.RequestException as e:
            return {
                "exists": False,
                "count": 0,
                "data": [],
                "message": f"Error de conexión: {str(e)}",
            }

    def cerrar_sesion(self):
        """
        Cierra la sesión actual
        """
        if self.session_token:
            headers = {
                "Session-Token": self.session_token,
                "App-Token": self.app_token.get(),
            }

            response = requests.post(
                f"{self.base_url.get()}/apirest.php/killSession", headers=headers
            )

            if response.status_code == 200:
                self.session_token = None
                return True
            raise Exception(f"Error al cerrar sesión: {response.text}")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_messages.append(f"[{timestamp}] {message}")

    def import_to_glpi(self):
        self.iniciar_sesion()

        if not all(
            [
                self.base_url.get(),
                self.app_token.get(),
                self.user_token.get(),
                self.session_token,
            ]
        ):
            messagebox.showerror(
                "Error", "Por favor complete todos los datos de conexión"
            )
            return

        if self.excel_data is None:
            messagebox.showerror("Error", "Por favor cargue un archivo Excel primero")
            return

        try:
            success_count = 0
            error_count = 0
            s = "HP ProBook 445 G9"
            modelo_id = self.buscar_modelos(s)
            self.log_message(f'{modelo_id["exists"]}')

            for idx, row in self.excel_data.iterrows():
                s = {"model": "HP ProBook 445 G7"}
                modelo_id = self.buscar_modelos(s)
                print(modelo_id)

                computer_data = {
                    "name": row["nombre"],
                    "serial": row["serie"],
                    "entities_id": 0,  # ID de la entidad (0 es la raíz)
                    "states_id": 1,  # ID del estado (por ejemplo, 1 puede ser "En uso")
                    "manufacturers_id": 1,  # ID del fabricante
                    "computermodels_id": 1,  # ID del modelo
                    "is_dynamic": 0,  # 0 para inventario manual
                }

                try:
                    response = requests.post(
                        f"{self.base_url.get()}/apirest.php/Computer/",
                        headers=self.headers,
                        json=computer_data,
                        verify=False,
                    )

                    if response.status_code in [200, 201]:
                        success_count += 1
                        self.log_message(f"Dispositivo creado: {row['name']}")
                    else:
                        error_count += 1
                        self.log_message(
                            f"Error al crear {row['name']}: {response.status_code}"
                        )
                except Exception as e:
                    error_count += 1
                    self.log_message(f"Error en {row['name']}: {str(e)}")
                finally:
                    self.cerrar_sesion()

            messagebox.showinfo(
                "Resultado",
                f"Importación completada\nExitosos: {success_count}\nErrores: {error_count}",
            )

        except Exception as e:
            messagebox.showerror("Error", f"Error en la importación: {str(e)}")
            self.log_message(f"Error general: {str(e)}")
        finally:
            self.cerrar_sesion()


def main():
    root = tk.Tk()
    app = GLPIDeviceImporter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
