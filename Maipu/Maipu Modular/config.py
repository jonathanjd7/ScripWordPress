import os
from selenium.webdriver.chrome.options import Options

URL_LOGIN = "https://grupomaipu.com/wp-login.php"
URL_NUEVO_POST = "https://grupomaipu.com/wp-admin/post-new.php"
USUARIO_WP = os.environ.get("USUARIO_WP", "GrupoMaipu2024")  # Cambiar en entorno protegido
PASSWORD_WP = os.environ.get("PASSWORD_WP", "1eH4.2NI>/&;")    # Cambiar en entorno protegido

CARPETA_WORD = r"C:\Users\Jonathan JD\Desktop\pink\Jonathan\Maipu\Maipu Modular"

options = Options()
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--disable-gpu")
