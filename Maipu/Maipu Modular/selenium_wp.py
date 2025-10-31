import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

def iniciar_navegador(options=None):
    if options is None:
        options = Options()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)
    return driver, wait

def login_wordpress(driver, wait, url_login, usuario, password):
    try:
        driver.get(url_login)
        username_field = wait.until(EC.visibility_of_element_located((By.ID, "user_login")))
        password_field = wait.until(EC.visibility_of_element_located((By.ID, "user_pass")))
        username_field.clear()
        username_field.send_keys(usuario)
        password_field.clear()
        password_field.send_keys(password)
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "wp-submit")))
        login_button.click()
        time.sleep(3)
        return True
    except Exception as e:
        print(f"[ERROR] Error en login: {e}")
        return False

def crear_nueva_categoria_clasico(driver, wait, nombre_categoria):
    """Crea una nueva categoría en WordPress (Editor Clásico) si no existe"""
    try:
        print(f"[INFO] Intentando crear nueva categoría: '{nombre_categoria}'")
        
        # Paso 1: Hacer clic en el enlace "Añadir una nueva categoría"
        try:
            añadir_link = None
            
            # Selector 1: Por ID
            try:
                añadir_link = wait.until(EC.element_to_be_clickable((By.ID, 'category-add-toggle')))
                print("[DEBUG] Enlace encontrado por ID 'category-add-toggle'")
            except:
                pass
            
            # Selector 2: Por clase category-add-toggle
            if not añadir_link:
                try:
                    añadir_link = driver.find_element(By.CLASS_NAME, 'category-add-toggle')
                    print("[DEBUG] Enlace encontrado por clase 'category-add-toggle'")
                except:
                    pass
            
            # Selector 3: Por texto "Añadir"
            if not añadir_link:
                try:
                    añadir_link = driver.find_element(By.XPATH, '//a[contains(text(), "Añadir") and contains(@class, "category")]')
                    print("[DEBUG] Enlace encontrado por texto 'Añadir'")
                except:
                    pass
            
            if not añadir_link:
                print("[ERROR] No se encontro el enlace 'Anadir categoria' (probados 3 selectores)")
                return False
            
            # Hacer clic en el enlace
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", añadir_link)
            time.sleep(1)
            
            # Usar JavaScript click
            try:
                driver.execute_script("arguments[0].click();", añadir_link)
                print("[OK] Enlace 'Anadir categoria' clickeado (JavaScript)")
            except:
                añadir_link.click()
                print("[OK] Enlace 'Anadir categoria' clickeado (click directo)")
            
            time.sleep(3)  # Dar más tiempo para que aparezca el formulario
            
        except Exception as e:
            print(f"[ERROR] Error al hacer clic en 'Anadir categoria': {e}")
            return False
        
        # Paso 2: Buscar el campo de texto para el nombre de la nueva categoría
        try:
            # Intentar múltiples selectores
            campo_nombre = None
            
            # Selector 1: ID estándar
            try:
                campo_nombre = wait.until(EC.visibility_of_element_located((By.ID, 'newcategory')))
                print("[DEBUG] Campo encontrado por ID 'newcategory'")
            except:
                pass
            
            # Selector 2: Por nombre
            if not campo_nombre:
                try:
                    campo_nombre = driver.find_element(By.NAME, 'newcategory')
                    print("[DEBUG] Campo encontrado por NAME 'newcategory'")
                except:
                    pass
            
            # Selector 3: Buscar input en el área de categorías que apareció
            if not campo_nombre:
                try:
                    # Buscar inputs de texto visibles en el área de categorías
                    campos = driver.find_elements(By.XPATH, '//div[@id="category-adder"]//input[@type="text"]')
                    if campos:
                        campo_nombre = campos[0]
                        print("[DEBUG] Campo encontrado en div category-adder")
                except:
                    pass
            
            # Selector 4: Cualquier input visible recién aparecido
            if not campo_nombre:
                try:
                    campos = driver.find_elements(By.XPATH, '//input[@type="text" and @aria-label]')
                    for campo in campos:
                        if campo.is_displayed():
                            campo_nombre = campo
                            print("[DEBUG] Campo encontrado por visibilidad")
                            break
                except:
                    pass
            
            if not campo_nombre:
                print("[ERROR] No se encontro el campo para el nombre de la categoria (probados 4 selectores)")
                return False
            
            # Ingresar el nombre de la categoría
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_nombre)
            time.sleep(0.5)
            campo_nombre.click()
            campo_nombre.clear()
            campo_nombre.send_keys(nombre_categoria)
            print(f"[OK] Nombre de categoria '{nombre_categoria}' ingresado")
            time.sleep(1)
        except Exception as ex:
            print(f"[ERROR] Error al buscar/ingresar campo de categoria: {ex}")
            return False
        
        # Paso 3: Hacer clic en el botón "Añadir nueva categoría"
        try:
            confirmar_btn = None
            
            # Selector 1: Por ID
            try:
                confirmar_btn = wait.until(EC.element_to_be_clickable((By.ID, 'category-add-submit')))
                print("[DEBUG] Boton encontrado por ID 'category-add-submit'")
            except:
                pass
            
            # Selector 2: Por clase o texto
            if not confirmar_btn:
                try:
                    confirmar_btn = driver.find_element(By.XPATH, '//input[@type="button" and contains(@value, "Añadir")]')
                    print("[DEBUG] Boton encontrado por XPATH con 'Añadir'")
                except:
                    pass
            
            # Selector 3: Cualquier botón en el área de categorías
            if not confirmar_btn:
                try:
                    confirmar_btn = driver.find_element(By.XPATH, '//div[@id="category-adder"]//input[@type="button"]')
                    print("[DEBUG] Boton encontrado en div category-adder")
                except:
                    pass
            
            if not confirmar_btn:
                print("[ERROR] No se encontro el boton de confirmar (probados 3 selectores)")
                # Intentar con Enter como último recurso
                try:
                    campo_nombre.send_keys(Keys.RETURN)
                    print("[OK] Enter presionado como alternativa")
                    time.sleep(2)
                    return True
                except:
                    return False
            
            # Hacer clic en el botón
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", confirmar_btn)
            time.sleep(0.5)
            
            # Usar JavaScript click
            try:
                driver.execute_script("arguments[0].click();", confirmar_btn)
                print("[OK] Boton 'Anadir nueva categoria' clickeado (JavaScript)")
            except:
                confirmar_btn.click()
                print("[OK] Boton 'Anadir nueva categoria' clickeado (click directo)")
            
            time.sleep(2)
            
            print(f"[SUCCESS] Categoria '{nombre_categoria}' creada exitosamente")
            return True
            
        except Exception as e:
            print(f"[ERROR] Error al confirmar creacion de categoria: {e}")
            return False
        
    except Exception as e:
        print(f"[ERROR] Error general al crear categoria '{nombre_categoria}': {e}")
        return False

def insertar_descripcion_classic_editor(driver, wait, descripcion):
    """Inserta descripción en Classic Editor - VERSIÓN MEJORADA CON SINCRONIZACIÓN"""
    try:
        print("[INFO] Configurando editor clásico para descripción...")
        
        # Paso 1: Cambiar a pestaña HTML
        try:
            html_btn = wait.until(EC.element_to_be_clickable((By.ID, 'content-html')))
            html_btn.click()
            print("[OK] Cambiado a modo HTML")
            time.sleep(2)
        except Exception as e:
            print(f"[ERROR] No se pudo cambiar a modo HTML: {e}")
            return False
        
        # Paso 2: Insertar contenido en el textarea
        try:
            # Buscar el textarea directamente
            textarea_editor = wait.until(EC.presence_of_element_located((By.ID, 'content')))
            
            # Limpiar contenido previo
            driver.execute_script("arguments[0].value = '';", textarea_editor)
            time.sleep(0.5)
            
            # Insertar la descripción usando JavaScript
            driver.execute_script("arguments[0].value = arguments[1];", textarea_editor, descripcion)
            
            # Disparar eventos para que WordPress detecte el cambio
            driver.execute_script("""
                var element = arguments[0];
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
            """, textarea_editor)
            
            print(f"[OK] Descripción insertada mediante JavaScript ({len(descripcion)} caracteres)")
            time.sleep(1)
            
            # Verificar que se insertó correctamente
            valor_actual = driver.execute_script("return arguments[0].value;", textarea_editor)
            if len(valor_actual) > 100:
                print(f"[OK] Verificación: Contenido presente ({len(valor_actual)} caracteres)")
            else:
                print(f"[WARNING] Verificación: Contenido parece vacío o corto ({len(valor_actual)} caracteres)")
            
        except Exception as e:
            print(f"[ERROR] Error al insertar en textarea: {e}")
            return False
        
        # Paso 3: NO cambiar a modo visual - Quedarse en HTML para preservar formatos
        print("[INFO] Manteniendo modo HTML para preservar formatos...")
        time.sleep(2)
        
        # Verificar una última vez que el contenido sigue en el textarea
        try:
            valor_final = driver.execute_script("return document.getElementById('content').value;")
            if len(valor_final) > 100:
                print(f"[OK] Contenido confirmado en modo HTML ({len(valor_final)} caracteres)")
                print(f"[OK] HTML contiene: {valor_final.count('<strong>')} negritas, {valor_final.count('<ol>')} listas numeradas")
            else:
                print(f"[WARNING] Contenido parece haberse perdido")
        except:
            print("[WARNING] No se pudo verificar contenido final")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error insertando descripción en editor clásico: {e}")
        # Asegurarse de volver al contenido principal en caso de error
        try:
            driver.switch_to.default_content()
        except:
            pass
        return False

def guardar_borrador(driver, wait):
    """Función para guardar el post como borrador - VERSIÓN MEJORADA CON JAVASCRIPT"""
    try:
        print("[INFO] Guardando como borrador...")
        
        # Método 1: Usar JavaScript para hacer click directamente (evita problemas de intercepción)
        try:
            driver.execute_script("document.getElementById('save-post').click();")
            print("[OK] Clic en botón 'Guardar borrador' mediante JavaScript")
        except Exception as e:
            print(f"[WARNING] Método JavaScript falló: {e}")
            
            # Método 2: Buscar botón y usar JavaScript con el elemento
            try:
                guardar_btn = wait.until(EC.presence_of_element_located((By.ID, 'save-post')))
                driver.execute_script("arguments[0].click();", guardar_btn)
                print("[OK] Clic en botón 'Guardar borrador' (Método 2)")
            except Exception as e2:
                print(f"[ERROR] Todos los métodos fallaron: {e2}")
                return False
        
        # Esperar a que se procese el guardado
        time.sleep(5)
        
        # Verificar mensaje de éxito
        try:
            exito_element = wait.until(EC.presence_of_element_located((By.ID, 'message')))
            mensaje = exito_element.text.lower()
            print(f"[INFO] Mensaje del sistema: {mensaje}")
            
            if any(palabra in mensaje for palabra in ['guardado', 'borrador', 'actualizado', 'saved', 'draft']):
                print("[OK] Guardado exitoso - Mensaje confirmado")
                return True
            else:
                print(f"[WARNING] Mensaje inesperado: {mensaje}")
                return True  # Continuamos aunque el mensaje no sea el esperado
                
        except TimeoutException:
            print("[WARNING] No se encontró mensaje de confirmación, pero continuamos")
            return True
            
    except Exception as e:
        print(f"[ERROR] Error al guardar: {e}")
        return False
