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
        self.root.geometry("1000x650")

        # Variables de control
        self.excel_path = tk.StringVar()
        self.base_url = tk.StringVar(value="http://10.14.35.145")
        print("base url: ", self.base_url.get())
        self.app_token = tk.StringVar(value="l1TqW2dIF38lxSPZstMzFFtDkrqzptV2H5BhdRxh")
        self.user_token = tk.StringVar(value="T99P59LLi8n1DIzTQBFT9uFt19kNQ0w4jiwsLgOH")
        self.session_token = tk.StringVar()
        self.headers = {
            "Content-Type": "application/json",
            "App-Token": self.app_token.get(),
            "Authorization": f"user_token {self.user_token.get()}",
        }
        # Crear el marco principal
        self.create_main_frame()
        # Variables para almacenar datos
        self.excel_data = None
        self.log_messages = []

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
                print()
                """"
                if pd.isna(row[name]):
                    validation_errors.append(f"Fila {idx + 2}: {name} vacío")"""

        if validation_errors:
            error_message = "\n".join(validation_errors)
            messagebox.showerror("Errores de Validación", error_message)
            self.log_message("Validación fallida")
        else:
            messagebox.showinfo("Éxito", "Validación exitosa")
            self.log_message("Validación exitosa")

    def iniciar_sesion(self):
        """
        Inicia sesión en GLPI usando user_token
        """
        headers = {
            "Content-Type": "application/json",
            "App-Token": self.app_token,
            "Authorization": f"user_token {self.user_token}",
        }

        response = requests.post(
            f"{self.base_url}/apirest.php/initSession", headers=headers
        )

        if response.status_code == 200:
            self.session_token = str(response.json()["session_token"])
            return True
        raise Exception(f"Error al iniciar sesión: {response.text}")

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

    def obtener_fabricantes(self):
        """
        Obtiene la lista de fabricantes
        """
        if not self.session_token:
            raise Exception("No hay sesión iniciada")

        headers = {"Session-Token": self.session_token, "App-Token": self.app_token}

        response = requests.get(
            f"{self.base_url}/apirest.php/Manufacturer", headers=headers
        )
        print("content: ", response.content)

        if response.status_code == 200:
            # Si la respuesta es un string, intentamos parsearlo como JSON
            if isinstance(response.text, str):
                try:
                    return json.loads(response.text)
                except json.JSONDecodeError:
                    raise Exception("Error al decodificar la respuesta JSON")
            return response.json()
        raise Exception(
            f"Error al obtener fabricantes: {response.status_code} - {response.text}"
        )

    def buscar_modelos(self, search_criteria):
        print(search_criteria)
        """
        Busca un modelo específico de computadora por nombre

        Args:
            nombre_modelo (str): Nombre del modelo a buscar

        Returns:
            dict: Información del modelo si se encuentra, None si no existe
        """
        if not self.session_token:
            raise Exception("No hay sesión iniciada")

        # Construir los parámetros de búsqueda
        search_params = {"criteria[0][link]": "AND"}

        # Mapeo de campos para la búsqueda
        field_mapping = {
            "name": 1,  # ID del campo nombre
            "serial": 45,  # ID del campo serial
            "inventory": 31,  # ID del campo número de inventario
            "model": 40,  # ID del campo modelo
        }
        model_endpoints = {
            "computadores": "/apirest.php/Computer",
            "impresoras": "/apirest.php/Printer",
            "monitores": "/apirest.php/Monitor",
            "dispositivos_red": "/apirest.php/NetworkEquipment",
            "perifericos": "/apirest.php/Peripheral",
            "telefonos": "/apirest.php/Phone",
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
            print(self.base_url.get())
            response = requests.get(
                f"{self.base_url.get()}/apirest.php/search/Computer/",
                headers=self.headers,
                params=search_params,
                verify=False,
            )

            if response.status_code == 200:
                data = response.json()

                # Verificar si se encontraron resultados
                if data.get("totalcount", 0) > 0:
                    return {
                        "exists": True,
                        "count": data.get("totalcount", 0),
                        "data": data.get("data", []),
                        "message": "Computador(es) encontrado(s)",
                    }
                else:
                    return {
                        "exists": False,
                        "count": 0,
                        "data": [],
                        "message": "No se encontró ningún computador con los criterios especificados",
                    }

            else:
                return {
                    "exists": False,
                    "count": 0,
                    "data": [],
                    "message": f"Error en la búsqueda: {response.status_code}",
                }

        except requests.exceptions.RequestException as e:
            return {
                "exists": False,
                "count": 0,
                "data": [],
                "message": f"Error de conexión: {str(e)}",
            }

    def obtener_ubicaciones(self):
        """
        Obtiene la lista de ubicaciones
        """
        if not self.session_token:
            raise Exception("No hay sesión iniciada")

        headers = {"Session-Token": self.session_token, "App-Token": self.app_token}

        response = requests.get(
            f"{self.base_url}/apirest.php/Location", headers=headers
        )

        if response.status_code == 200:
            return response.json()
        raise Exception(f"Error al obtener ubicaciones: {response.text}")

    def obtener_estados(self):
        """
        Obtiene la lista de estados de dispositivos
        """
        if not self.session_token:
            raise Exception("No hay sesión iniciada")

        headers = {"Session-Token": self.session_token, "App-Token": self.app_token}

        response = requests.get(f"{self.base_url}/apirest.php/State", headers=headers)

        if response.status_code == 200:
            return response.json()
        raise Exception(f"Error al obtener estados: {response.text}")

    def cerrar_sesion(self):
        """
        Cierra la sesión actual
        """
        if self.session_token:
            headers = {"Session-Token": self.session_token, "App-Token": self.app_token}

            response = requests.post(
                f"{self.base_url}/apirest.php/killSession", headers=headers
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
        if not all([self.base_url.get(), self.app_token.get(), self.user_token.get()]):
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

            for idx, row in self.excel_data.iterrows():
                modelo_id = self.buscar_modelos({"model": row["modelo"]})
                print(modelo_id["exits"])

                computer_data = {
                    "name": row["nombre"],
                    "serial": row["serie"],
                    "entities_id": 0,  # ID de la entidad (0 es la raíz)
                    "states_id": 1,  # ID del estado (por ejemplo, 1 puede ser "En uso")
                    "manufacturers_id": 1,  # ID del fabricante
                    "computermodels_id": 1,  # ID del modelo
                    "is_dynamic": 0,  # 0 para inventario manual
                }
                """
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
                """
            messagebox.showinfo(
                "Resultado",
                f"Importación completada\nExitosos: {success_count}\nErrores: {error_count}",
            )

        except Exception as e:
            messagebox.showerror("Error", f"Error en la importación: {str(e)}")
            self.log_message(f"Error general: {str(e)}")


def main():
    root = tk.Tk()
    app = GLPIDeviceImporter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
