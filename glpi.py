import json
import requests


class GLPIApi:
    def __init__(self, url_base, app_token, user_token):
        """
        Inicializa la conexión con GLPI usando user_token

        Args:
            url_base (str): URL base de GLPI (ejemplo: 'http://glpi.midominio.com')
            app_token (str): Token de aplicación de GLPI
            user_token (str): Token de usuario para autenticación
        """
        self.url_base = url_base.rstrip("/")
        self.app_token = str(app_token)
        self.user_token = str(user_token)
        self.session_token = None

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
            f"{self.url_base}/apirest.php/initSession", headers=headers
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
            f"{self.url_base}/apirest.php/{tipo_dispositivo}",
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
            f"{self.url_base}/apirest.php/Manufacturer", headers=headers
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
        """
        Busca un modelo específico de computadora por nombre

        Args:
            nombre_modelo (str): Nombre del modelo a buscar

        Returns:
            dict: Información del modelo si se encuentra, None si no existe
        """
        if not self.session_token:
            raise Exception("No hay sesión iniciada")

        headers = {"Session-Token": self.session_token, "App-Token": self.app_token}
        print("my headers", headers)
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
            response = requests.get(
                f"{self.url_base}/apirest.php/search/Computer/",
                headers=headers,
                params=search_params,
                verify=False,
            )
            print(response.url, response.headers)
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
            f"{self.url_base}/apirest.php/Location", headers=headers
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

        response = requests.get(f"{self.url_base}/apirest.php/State", headers=headers)

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
                f"{self.url_base}/apirest.php/killSession", headers=headers
            )

            if response.status_code == 200:
                self.session_token = None
                return True
            raise Exception(f"Error al cerrar sesión: {response.text}")


# Configuración
url_glpi = "http://10.14.35.145"
app_token = "l1TqW2dIF38lxSPZstMzFFtDkrqzptV2H5BhdRxh"
user_token = "T99P59LLi8n1DIzTQBFT9uFt19kNQ0w4jiwsLgOH"

# Crear instancia
glpi = GLPIApi(url_glpi, app_token, user_token)

try:
    # Iniciar sesión
    glpi.iniciar_sesion()
    # Ejemplo: Crear una computadora

    # Obtener modelos de computadora disponibles
    print("\nModelos de computadora disponibles:")
    search_model = {"model": "HP ProBook 445 G7"}

    modelos = glpi.buscar_modelos(search_model)

    print("model founded: ", modelos["exists"])

    datos_computadora = {
        "name": "PC-001",
        "serial": "SN123456",
        "entities_id": 0,  # ID de la entidad (0 es la raíz)
        "states_id": 1,  # ID del estado (por ejemplo, 1 puede ser "En uso")
        "manufacturers_id": 1,  # ID del fabricante
        "computermodels_id": 1,  # ID del modelo
        "is_dynamic": 0,  # 0 para inventario manual
    }

    # nueva_computadora = glpi.crear_dispositivo("Computer", datos_computadora)
    # print(f"Computadora creada con ID: {nueva_computadora}")

except Exception as e:
    print(f"Error: {str(e)}")
finally:
    glpi.cerrar_sesion()
