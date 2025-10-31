from extraccion_doc import (
    obtener_archivos_word_carpeta, 
    extraer_titulo_doc, 
    frase_clave_doc, 
    titulo_seo_doc, 
    meta_description_doc, 
    leer_etiquetas_doc, 
    leer_categorias_doc,
    extraer_descripcion_con_formato_doc
)
from selenium_wp import (
    iniciar_navegador, 
    login_wordpress, 
    crear_nueva_categoria_clasico, 
    insertar_descripcion_classic_editor, 
    guardar_borrador
)
from utilidades import configurar_logging, pause
import config
from docx import Document
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import time
import os

def procesar_archivo(archivo_word, driver, wait):
    """Procesa un archivo Word individual y lo guarda como borrador en WordPress"""
    print(f"\n{'='*60}")
    print(f"[INFO] Procesando: {os.path.basename(archivo_word)}")
    print(f"{'='*60}")
    
    try:
        # Navegar a la página de nuevo post
        driver.get(config.URL_NUEVO_POST)
        time.sleep(4)

        doc = Document(archivo_word)
        titulo = extraer_titulo_doc(doc)
        descripcion = extraer_descripcion_con_formato_doc(doc)
        frase_obj = frase_clave_doc(doc)
        tit_seo = titulo_seo_doc(doc)
        meta_desc = meta_description_doc(doc)
        etiquetas = leer_etiquetas_doc(doc)
        categorias = leer_categorias_doc(doc)

        print(f"[INFO] Título: {titulo if titulo else '[VACÍO - REVISAR DOCUMENTO]'}")
        if descripcion:
            print(f"[INFO] Descripción: {descripcion[:100]}... ({len(descripcion)} caracteres)")
        else:
            print(f"[WARNING] Descripción: [VACÍA - REVISAR DOCUMENTO]")
        print(f"[INFO] Frase clave: {frase_obj}")
        print(f"[INFO] Título SEO: {tit_seo}")
        print(f"[INFO] Meta desc: {meta_desc}")
        print(f"[INFO] Etiquetas: {etiquetas}")
        print(f"[INFO] Categorías: {categorias}")

        # ---------------- INSERTAR TÍTULO ----------------
        print("[INFO] Insertando título...")
        title_field = wait.until(EC.element_to_be_clickable((By.ID, 'title')))
        title_field.click()
        title_field.clear()
        title_field.send_keys(titulo)
        time.sleep(1)

        # ---------------- INSERTAR DESCRIPCIÓN ----------------
        if descripcion.strip():
            print(f"[INFO] Insertando descripción en editor clásico ({len(descripcion)} caracteres)...")
            print(f"[DEBUG] La descripción contiene {descripcion.count('<strong>')} negritas, {descripcion.count('<ol>')} listas numeradas")
            
            if not insertar_descripcion_classic_editor(driver, wait, descripcion):
                print("[ERROR] No se pudo insertar la descripción")
                return False
            
            # Espera adicional para que WordPress procese la descripción
            print("[INFO] Esperando a que WordPress procese la descripción...")
            time.sleep(3)
        else:
            print("[WARNING] No hay descripción para insertar")

        # ---------------- INSERTAR FRASE CLAVE ----------------
        print("[INFO] Insertando frase clave...")
        try:
            frase_field = wait.until(EC.visibility_of_element_located((By.ID, 'focus-keyword-input-metabox')))
            frase_field.clear()
            frase_field.send_keys(Keys.CONTROL, 'a')
            frase_field.send_keys(Keys.DELETE)
            frase_field.send_keys(frase_obj)
            print("[OK] Frase clave insertada")
        except TimeoutException:
            print("[WARNING] Campo de frase clave no encontrado")

        # ---------------- INSERTAR TÍTULO SEO ----------------
        print("[INFO] Insertando título SEO...")
        try:
            # Este campo es más complejo, intentamos enfoque diferente
            titulo_seo_area = wait.until(EC.element_to_be_clickable((By.ID, 'yoast-google-preview-title-metabox')))
            titulo_seo_area.click()
            titulo_seo_area.send_keys(Keys.CONTROL, 'a')
            titulo_seo_area.send_keys(Keys.DELETE)
            titulo_seo_area.send_keys(tit_seo)
            print("[OK] Título SEO insertado")
        except TimeoutException:
            print("[WARNING] Campo de título SEO no encontrado")

        # ---------------- INSERTAR META DESCRIPCIÓN ----------------
        print("[INFO] Insertando meta descripción...")
        try:
            meta_desc_field = wait.until(EC.element_to_be_clickable((By.ID, 'yoast-google-preview-description-metabox')))
            meta_desc_field.click()
            meta_desc_field.send_keys(Keys.CONTROL, 'a')
            meta_desc_field.send_keys(Keys.DELETE)
            meta_desc_field.send_keys(meta_desc)
            print("[OK] Meta descripción insertada")
        except TimeoutException:
            print("[WARNING] Campo de meta descripción no encontrado")

        # ---------------- INSERTAR CATEGORÍAS ----------------
        print("[INFO] Configurando categorías...")
        print(f"[DEBUG] Total de categorías detectadas: {len(categorias)}")
        if categorias:
            print(f"[DEBUG] Categorías a procesar: {', '.join(categorias)}")
        
        try:
            for categoria in categorias:
                categoria_limpia = categoria.strip()
                if not categoria_limpia:
                    continue
                    
                print(f"[DEBUG] Procesando categoría: '{categoria_limpia}'")
                checkbox_encontrado = False
                
                try:
                    # Método 1: Buscar por texto exacto del label
                    xpath = f"//label[normalize-space(text()) = '{categoria_limpia}']/input[@type='checkbox']"
                    checkbox = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                    
                    # Hacer scroll para que sea visible
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                    time.sleep(0.5)
                    
                    if not checkbox.is_selected():
                        # Usar JavaScript click para evitar interceptación
                        try:
                            driver.execute_script("arguments[0].click();", checkbox)
                            print(f"[OK] Categoría seleccionada: {categoria_limpia}")
                        except:
                            checkbox.click()
                            print(f"[OK] Categoría seleccionada (click directo): {categoria_limpia}")
                    else:
                        print(f"[INFO] Categoría ya estaba seleccionada: {categoria_limpia}")
                    checkbox_encontrado = True
                    time.sleep(0.5)
                    
                except TimeoutException:
                    print(f"[WARNING] No se encontró la categoría '{categoria_limpia}'")
                    
                    # Método 2: Intentar búsqueda más flexible
                    try:
                        xpath_flex = f"//label[contains(normalize-space(text()), '{categoria_limpia}')]/input[@type='checkbox']"
                        checkbox = driver.find_element(By.XPATH, xpath_flex)
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                        time.sleep(0.5)
                        
                        if not checkbox.is_selected():
                            try:
                                driver.execute_script("arguments[0].click();", checkbox)
                                print(f"[OK] Categoría seleccionada (búsqueda flexible): {categoria_limpia}")
                            except:
                                checkbox.click()
                                print(f"[OK] Categoría seleccionada (búsqueda flexible - click directo): {categoria_limpia}")
                        checkbox_encontrado = True
                        time.sleep(0.5)
                    except:
                        print(f"[WARNING] Categoría '{categoria_limpia}' no existe, intentando crearla...")
                        
                        # Método 3: Crear la categoría si no existe
                        if crear_nueva_categoria_clasico(driver, wait, categoria_limpia):
                            print(f"[OK] Categoría '{categoria_limpia}' creada exitosamente")
                            
                            # Esperar y buscar la categoría recién creada para seleccionarla
                            time.sleep(2)
                            try:
                                xpath = f"//label[normalize-space(text()) = '{categoria_limpia}']/input[@type='checkbox']"
                                checkbox = driver.find_element(By.XPATH, xpath)
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                                time.sleep(0.5)
                                
                                if not checkbox.is_selected():
                                    try:
                                        driver.execute_script("arguments[0].click();", checkbox)
                                        print(f"[OK] Categoría '{categoria_limpia}' seleccionada después de crearla")
                                    except:
                                        checkbox.click()
                                        print(f"[OK] Categoría '{categoria_limpia}' seleccionada después de crearla (click directo)")
                                checkbox_encontrado = True
                            except Exception as e:
                                print(f"[WARNING] Se creó la categoría pero no se pudo seleccionar: {e}")
                        else:
                            print(f"[ERROR] No se pudo crear la categoría '{categoria_limpia}'")
                
                except Exception as e:
                    print(f"[ERROR] Error inesperado con la categoría '{categoria_limpia}': {e}")
                    
        except Exception as e:
            print(f"[WARNING] Error general en categorías: {e}")

        # ---------------- INSERTAR ETIQUETAS ----------------
        print("[INFO] Configurando etiquetas...")
        try:
            if etiquetas.strip():
                etiq_field = wait.until(EC.visibility_of_element_located((By.ID, 'new-tag-post_tag')))
                etiq_field.clear()
                etiq_field.send_keys(etiquetas)
                etiq_field.send_keys(Keys.ENTER)
                print("[OK] Etiquetas insertadas")
        except Exception as e:
            print(f"[WARNING] Error insertando etiquetas: {e}")

        # ---------------- PREPARAR PÁGINA PARA GUARDAR ----------------
        print("[INFO] Preparando página para guardar...")
        time.sleep(3)
        
        # Hacer scroll para asegurar que todos los elementos estén cargados
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # Verificación final: asegurar que la descripción está en el textarea
        if descripcion.strip():
            try:
                print("[DEBUG] Verificación final de descripción antes de guardar...")
                textarea_content = driver.execute_script("return document.getElementById('content').value;")
                if len(textarea_content) > 100:
                    print(f"[OK] Descripción confirmada en textarea ({len(textarea_content)} caracteres)")
                else:
                    print(f"[WARNING] Descripción parece vacía en textarea, reinsertando...")
                    # Reinsertar si está vacío
                    driver.execute_script("document.getElementById('content').value = arguments[0];", descripcion)
                    print(f"[OK] Descripción reinsertada ({len(descripcion)} caracteres)")
            except Exception as e:
                print(f"[WARNING] No se pudo verificar descripción: {e}")

        # ---------------- GUARDAR COMO BORRADOR ----------------
        print("[INFO] Guardando como borrador...")

        if guardar_borrador(driver, wait):
            print(f"[SUCCESS] Post '{titulo}' guardado como borrador con éxito!")
            return True
        else:
            print(f"[ERROR] No se pudo guardar '{titulo}' como borrador")
            return False

    except Exception as e:
        print(f"[ERROR] Error procesando {archivo_word}: {e}")
        return False

def main():
    # Obtener todos los archivos Word de la carpeta
    archivos = obtener_archivos_word_carpeta(config.CARPETA_WORD)
    
    if not archivos:
        print(f"[ERROR] No se encontraron archivos .docx en la carpeta: {config.CARPETA_WORD}")
        return
    
    print(f"[INFO] Se encontraron {len(archivos)} archivos Word:")
    for i, archivo in enumerate(archivos, 1):
        print(f"  {i}. {os.path.basename(archivo)}")
    
    # Inicializar el navegador una sola vez
    print("[INFO] Inicializando navegador...")
    driver, wait = iniciar_navegador(config.options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    try:
        # Login una sola vez
        print("[INFO] Iniciando sesión en WordPress...")
        if not login_wordpress(driver, wait, config.URL_LOGIN, config.USUARIO_WP, config.PASSWORD_WP):
            print("[ERROR] No se pudo hacer login en WordPress")
            return
        
        # Esperar a que cargue el dashboard
        wait.until(EC.presence_of_element_located((By.ID, 'wpadminbar')))
        print("[OK] Login exitoso")

        # Procesar cada archivo
        exitosos = 0
        total = len(archivos)
        
        for i, archivo in enumerate(archivos, 1):
            print(f"\n[INFO] Progreso: {i}/{total}")
            if procesar_archivo(archivo, driver, wait):
                exitosos += 1
            pause(5)  # Espera entre guardados
        
        print(f"\n{'='*60}")
        print(f"[SUCCESS] PROCESO COMPLETADO")
        print(f"{'='*60}")
        print(f"[OK] Archivos guardados como borrador: {exitosos}/{total}")
        print(f"[ERROR] Archivos fallidos: {total - exitosos}")
        print(f"{'='*60}")

    except Exception as e:
        print(f"[ERROR] Error en el proceso principal: {e}")
        
    finally:
        time.sleep(3)
        driver.quit()
        print("[INFO] Navegador cerrado.")

if __name__ == "__main__":
    main()
