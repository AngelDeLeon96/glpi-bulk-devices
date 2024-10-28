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
        ancho_ventana = 900
        alto_ventana = 900
        # Calcular la posición para centrar la ventana
        ancho_pantalla = root.winfo_screenwidth()
        alto_pantalla = root.winfo_screenheight()
        pos_x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        pos_y = (alto_pantalla // 2) - (alto_ventana // 2)

        # Configurar geometría de la ventana (ancho x alto + x + y)
        self.root.geometry(f"{ancho_ventana}x{alto_ventana}+{pos_x}+{pos_y}")

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

    def buscar_dispositivos(
        self,
        tipo_dispositivo="Computer",
        criterios={},
        campos_mostrar=None,
    ):
        """
        Realiza una búsqueda avanzada de dispositivos en GLPI

        Args:
            tipo_dispositivo (str): Tipo de dispositivo (Computer, Printer, Monitor, etc.)
            criterios (dict): Criterios de búsqueda
            campos_mostrar (list): Lista de IDs de campos a mostrar en el resultado

        Returns:
            dict: Resultado de la búsqueda
        """
        print("headerws", self.headers)
        nombre_modelo = criterios["model"]
        # Mapeo de campos comunes de búsqueda
        campos_busqueda = {
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

        # Parámetros base de búsqueda
        search_params = {
            "is_deleted": 0,  # No mostrar elementos eliminados
            "as_map": 0,  # Formato detallado de respuesta
        }

        # Agregar criterios de búsqueda si existen
        if criterios:
            for idx, (campo, valor) in enumerate(criterios.items()):
                if campo in campos_busqueda:
                    search_params.update(
                        {
                            f"criteria[{idx}][link]": "AND" if idx > 0 else "",
                            f"criteria[{idx}][field]": campos_busqueda[campo],
                            f"criteria[{idx}][searchtype]": "contains",
                            f"criteria[{idx}][value]": valor,
                        }
                    )

        # Campos a mostrar en el resultado
        default_fields = [1, 2, 45, 40, 3, 24, 23]  # Campos por defecto
        fields_to_show = campos_mostrar if campos_mostrar else default_fields

        # Agregar campos a mostrar
        for idx, field_id in enumerate(fields_to_show):
            search_params[f"forcedisplay[{idx}]"] = field_id

        try:
            # Realizar la búsqueda
            response = requests.get(
                f"{self.base_url.get()}/apirest.php/search/{tipo_dispositivo}/",
                headers=self.headers,
                params=search_params,
                verify=False,
            )

            # Log de la petición para debugging
            print(f"URL de búsqueda: {response.url}")
            print(f"Parámetros: {search_params}")
            print(f"Código de respuesta: {response.status_code}")

            if response.status_code == 200:
                try:
                    modelos = response.json()
                    print("modelos encontrados: ", modelos)
                    # Formatear los resultados
                    # Si encontramos modelos
                    """
                    if isinstance(modelos, list):
                        # Buscar coincidencia exacta primero
                        for modelo in modelos:
                            if (
                                modelo.get("name", "").lower()
                                == criterios["model"].lower()
                            ):
                                return {
                                    "success": True,
                                    "model_id": modelo["id"],
                                    "name": modelo["name"],
                                    "message": "Modelo encontrado exactamente",
                                }

                        # Si no hay coincidencia exacta, buscar coincidencia parcial
                        for modelo in modelos:
                            if nombre_modelo.lower() in modelo.get("name", "").lower():
                                return {
                                    "success": True,
                                    "model_id": modelo["id"],
                                    "name": modelo["name"],
                                    "message": "Modelo similar encontrado",
                                }

                        return {
                            "success": False,
                            "model_id": None,
                            "message": f'No se encontró el modelo "{nombre_modelo}"',
                        }
                    """
                except json.JSONDecodeError as e:
                    self.log_message(f"Error decodificando JSON: {str(e)}")
                    self.log_message(f"Respuesta raw: {response.text}")
                    return {
                        "success": False,
                        "message": f"Error procesando la respuesta: {str(e)}",
                    }
            else:
                error_msg = f"Error en la búsqueda (Código {response.status_code})"
                try:
                    error_data = response.json()
                    error_msg += f": {error_data.get('message', '')}"
                except:
                    error_msg += f": {response.text}"

                return {"success": False, "message": error_msg}

        except requests.exceptions.RequestException as e:
            error_msg = f"Error de conexión: {str(e)}"
            self.log_message(error_msg)
            return {"success": False, "message": error_msg}

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
        dispositivos_dic = {}
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
            count = 0
            for idx, row in self.excel_data.iterrows():
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
                    count = count + 1
                    print("searching ", row["tipo"])

                    modelo_id = self.buscar_dispositivos(
                        tipo_dispositivo=row["tipo"], criterios={"model": row["modelo"]}
                    )

                    # print(modelo_id)
                    if count == 7:
                        break
                    """
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
                    """
                except Exception as e:
                    error_count += 1
                    print(e)
                    self.log_message(f"Error en {row['name']}: {str(e)}")

                """messagebox.showinfo(
                "Resultado",
                f"Importación completada\nExitosos: {success_count}\nErrores: {error_count}",
            )"""

        except Exception as e:
            messagebox.showerror("Error", f"Error en la importación: {str(e)}")
            self.log_message(f"Error general: {str(e)}")


def main():
    root = tk.Tk()
    app = GLPIDeviceImporter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
