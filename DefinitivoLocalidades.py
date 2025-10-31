import os
import pandas as pd
from docxtpl import DocxTemplate
import datetime
import locale
import re
import random

# ============================
# 1. CONFIGURACIÓN INICIAL
# ============================

# Configurar la localización para fechas en español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except:
        print("[AVISO] No se pudo configurar el idioma español para fechas")

# ============================
# ⚙️ CONTROL DE PRUEBA - EDITA AQUÍ
# ============================
# Cambia NUM_FICHAS_PRUEBA para controlar cuántas fichas generar:
#   None o 0  --> Genera TODAS las fichas del archivo
#   10        --> Genera solo las primeras 10 fichas (PRUEBA)
#   50        --> Genera solo las primeras 50 fichas
NUM_FICHAS_PRUEBA = 10  # <--- CAMBIA ESTE NÚMERO AQUÍ
# ============================

# Variables globales
SECTOR = 'Empresas de la localidad '  # Verificar que 'SECTOR' coincida con el nombre del sector
MES = datetime.datetime.now().strftime("%B")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERBOSE = False  # Controla el nivel de salida por consola

# Rutas relativas (PORTABLE - no necesita modificación)
OUTPUT_DIRECTORY_CSV = os.path.join(BASE_DIR, 'csv_creados')
ARCHIVO_EXCEL_LOCALIDADES = os.path.join(BASE_DIR, 'datos_maestros', '1_maestro_localidades_nw_20251028_2.xlsx')
PLANTILLA_LOCALIDADES = os.path.join(BASE_DIR, 'docsparascript', 'plantilla_localidades.docx')
PLANTILLAS_CREADAS_DIR = os.path.join(BASE_DIR, 'Plantillas_creadas')

# Crear directorios necesarios
os.makedirs(OUTPUT_DIRECTORY_CSV, exist_ok=True)
os.makedirs(PLANTILLAS_CREADAS_DIR, exist_ok=True)

# Validación de rutas
if not os.path.exists(ARCHIVO_EXCEL_LOCALIDADES):
    raise FileNotFoundError(f"El archivo {ARCHIVO_EXCEL_LOCALIDADES} no existe.")

if not os.path.exists(PLANTILLA_LOCALIDADES):
    raise FileNotFoundError(f"La plantilla de localidades {PLANTILLA_LOCALIDADES} no existe.")

if VERBOSE:
    print("Configuracion cargada correctamente.")


# ============================
# 2. FUNCIONES
# ============================

def eliminar_acentos_slash(text):
    """Elimina acentos y caracteres especiales para nombres de archivo"""
    replacements = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ä': 'a', 'ã': 'a', 'å': 'a',
        'Á': 'A', 'À': 'A', 'Â': 'A', 'Ä': 'A', 'Ã': 'A', 'Å': 'A',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'É': 'E', 'È': 'E', 'Ê': 'E', 'Ë': 'E',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'Í': 'I', 'Ì': 'I', 'Î': 'I', 'Ï': 'I',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'ö': 'o', 'õ': 'o', 'ø': 'o',
        'Ó': 'O', 'Ò': 'O', 'Ô': 'O', 'Ö': 'O', 'Õ': 'O', 'Ø': 'O',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
        'Ú': 'U', 'Ù': 'U', 'Û': 'U', 'Ü': 'U',
        'ý': 'y', 'ÿ': 'y', 'Ý': 'Y', 'Ÿ': 'Y',
        'ñ': 'n', 'Ñ': 'n', '/': '_', ' ': '_', '-': '_'
    }
    text = str(text)
    text = re.sub(r'[áàâäãåéèêëíìîïóòôöõøúùûüýÿñÁÀÂÄÃÅÉÈÊËÍÌÎÏÓÒÔÖÕØÚÙÛÜÝŸÑ/ -]', lambda x: replacements[x.group()], text)
    return re.sub(r'_+', '_', text)


def capitalizar_localidad(text):
    """Capitaliza correctamente nombres de localidades"""
    # Palabras que deben ir en minúscula (excepto al inicio)
    palabras_minusculas = ['de', 'del', 'la', 'el', 'los', 'las', 'y']
    
    palabras = text.split()
    resultado = []
    
    for i, palabra in enumerate(palabras):
        if i == 0:  # Primera palabra siempre capitalizada
            resultado.append(palabra.capitalize())
        elif palabra.lower() in palabras_minusculas:
            resultado.append(palabra.lower())
        else:
            resultado.append(palabra.capitalize())
    
    return ' '.join(resultado)


def normalizar_texto_para_url(text):
    """Normaliza texto para generar URLs: elimina acentos, convierte a minúsculas, 
    reemplaza espacios con guiones, elimina paréntesis y caracteres especiales"""
    if not text or pd.isna(text):
        return ''
    
    text = str(text).strip()
    
    # PRIMERO: Tomar solo la parte antes del primer paréntesis (si existe)
    # Esto maneja casos como "A Aldea De Arriba (Santa Marta...)" -> "A Aldea De Arriba"
    # IMPORTANTE: Esto debe ir ANTES del caso especial para evitar confundir " - " dentro de paréntesis
    if '(' in text:
        text = text.split('(')[0].strip()
    
    # DESPUÉS: Caso especial: "A Coruña - La Coruña" -> solo usar "La Coruña"
    # Esto solo debe aplicarse si NO hay paréntesis (ya eliminados arriba)
    if ' - ' in text:
        text = text.split(' - ')[-1].strip()
    
    # Manejar paréntesis de cierre sueltos al final (sin apertura antes)
    # Ejemplo: "A Barcala Cambre)" -> "A Barcala"
    if ')' in text and '(' not in text:
        # Encontrar el último espacio antes del ')' y eliminar desde ahí
        idx_cierre = text.rfind(')')
        idx_espacio = text.rfind(' ', 0, idx_cierre)
        if idx_espacio != -1:
            text = text[:idx_espacio].strip()
        else:
            # Si no hay espacio, eliminar el ')' y todo lo que esté junto
            text = text[:idx_cierre].strip()
    
    # Eliminar comas y otros caracteres especiales al inicio/final
    text = re.sub(r'^[,\s]+|[,\s]+$', '', text)
    
    # Eliminar cualquier paréntesis suelto restante (por si acaso quedó alguno)
    text = text.replace('(', '').replace(')', '').strip()
    
    # Reemplazos de acentos y caracteres especiales
    replacements = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ä': 'a', 'ã': 'a', 'å': 'a',
        'Á': 'a', 'À': 'a', 'Â': 'a', 'Ä': 'a', 'Ã': 'a', 'Å': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'É': 'e', 'È': 'e', 'Ê': 'e', 'Ë': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'Í': 'i', 'Ì': 'i', 'Î': 'i', 'Ï': 'i',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'ö': 'o', 'õ': 'o', 'ø': 'o',
        'Ó': 'o', 'Ò': 'o', 'Ô': 'o', 'Ö': 'o', 'Õ': 'o', 'Ø': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
        'Ú': 'u', 'Ù': 'u', 'Û': 'u', 'Ü': 'u',
        'ý': 'y', 'ÿ': 'y', 'Ý': 'y', 'Ÿ': 'y',
        'ñ': 'n', 'Ñ': 'n',
    }
    
    # Convertir a minúsculas primero
    text = text.lower()
    
    # Reemplazar acentos
    for original, reemplazo in replacements.items():
        text = text.replace(original, reemplazo)
    
    # Reemplazar espacios con guiones
    text = text.replace(' ', '-')
    
    # Eliminar caracteres no permitidos en URLs (mantener solo letras, números y guiones)
    text = re.sub(r'[^a-z0-9-]', '', text)
    
    # Eliminar guiones múltiples consecutivos
    text = re.sub(r'-+', '-', text)
    
    # Eliminar guiones al inicio y final
    text = text.strip('-')
    
    return text


def generar_url_localidad(localidad, provincia):
    """Genera la URL correcta para una localidad según el formato:
    https://gsas.es/producto/base-de-datos-de-empresas-de-la-localidad-de-[localidad]-en-la-provincia-de-[provincia]/
    """
    localidad_normalizada = normalizar_texto_para_url(localidad)
    provincia_normalizada = normalizar_texto_para_url(provincia)
    
    if not localidad_normalizada or not provincia_normalizada:
        return ''
    
    url = f'https://gsas.es/producto/base-de-datos-de-empresas-de-la-localidad-de-{localidad_normalizada}-en-la-provincia-de-{provincia_normalizada}/'
    return url


def get_random_urls(ARCHIVO_EXCEL_LOCALIDADES, num_urls=10):
    """Obtiene URLs aleatorias de la columna 'url' del archivo de localidades"""
    df = pd.read_excel(ARCHIVO_EXCEL_LOCALIDADES)
    
    # Buscar la columna 'url' sin distinguir mayúsculas/minúsculas
    columna_url = None
    for col in df.columns:
        if col.lower().strip() == 'url':
            columna_url = col
            break
    
    if columna_url is None:
        columnas_disponibles = ', '.join(df.columns.tolist())
        raise ValueError(f"El archivo de localidades no contiene la columna 'url'. Columnas disponibles: {columnas_disponibles}")
    
    urls = df[columna_url].dropna().astype(str).tolist()
    # Filtrar URLs vacías o que solo contengan espacios
    urls = [url.strip() for url in urls if url.strip()]
    
    if len(urls) < num_urls:
        raise ValueError(f"No hay suficientes URLs. Solo hay {len(urls)} disponibles, se necesitan {num_urls}")
    urls_seleccionadas = random.sample(urls, num_urls)
    return {f'URL_{i+1}': url for i, url in enumerate(urls_seleccionadas)}


def generar_plantilla_localidad(ARCHIVO_EXCEL_LOCALIDADES):
    """Genera la plantilla específica para localidades con placeholders genéricos"""
    urls_diccionario = get_random_urls(ARCHIVO_EXCEL_LOCALIDADES, num_urls=10)
    
    context = {
        # Placeholders genéricos que se rellenarán después
        'Localidad': '{{Localidad}}',
        'Provincia': '{{Provincia}}',
        'Precio': '{{Precio}}',
        'registros': '{{registros}}',
        'dir': '{{dir}}',
        'phone': '{{phone}}',
        'mail': '{{mail}}',
        'cat': '{{cat}}',
        'web': '{{web}}',
        'Mes': MES,
        'Sector': SECTOR,
        # URLs del blog
        'URL_1': urls_diccionario['URL_1'],
        'URL_2': urls_diccionario['URL_2'],
        'URL_3': urls_diccionario['URL_3'],
        'URL_4': urls_diccionario['URL_4'],
        'URL_5': urls_diccionario['URL_5'],
        'URL_6': urls_diccionario['URL_6'],
        'URL_7': urls_diccionario['URL_7'],
        'URL_8': urls_diccionario['URL_8'],
        'URL_9': urls_diccionario['URL_9'],
        'URL_10': urls_diccionario['URL_10']
    }
    
    doc_generico = DocxTemplate(PLANTILLA_LOCALIDADES)
    
    # Crear directorio para plantillas
    sector_directory = os.path.join(PLANTILLAS_CREADAS_DIR, eliminar_acentos_slash(SECTOR), f'Plantillas_{eliminar_acentos_slash(SECTOR)}')
    os.makedirs(sector_directory, exist_ok=True)
    
    # Nombre de la plantilla
    nombre_archivo = 'Plantilla_localidad.docx'
    ruta_archivo = os.path.join(sector_directory, nombre_archivo)
    
    doc_generico.render(context)
    doc_generico.save(ruta_archivo)
    if VERBOSE:
        print(f'Plantilla "{ruta_archivo}" generada con exito.')
    
    return DocxTemplate(ruta_archivo)


def generar_documento_localidad(context, doc_plantilla, nombre_archivo, sector_directory):
    """Genera un documento individual para una localidad"""
    doc_plantilla.render(context)
    ruta_archivo = os.path.join(sector_directory, nombre_archivo)
    doc_plantilla.save(ruta_archivo)
    if VERBOSE:
        print(f'Documento "{ruta_archivo}" generado con exito.')


def generador_csv_con_url_localidades(sector_directory_seleccionar, df_localidades_procesadas):
    """Genera CSV con URLs desde el DataFrame procesado.
    
    Args:
        sector_directory_seleccionar: Directorio donde guardar el CSV
        df_localidades_procesadas: DataFrame con las localidades procesadas (debe tener columna 'url')
    """
    try:
        # Obtener URLs del DataFrame (ya generadas anteriormente)
        urls = df_localidades_procesadas['url'].dropna().astype(str).tolist()
        urls = [url.strip() for url in urls if url.strip()]
        
        if not urls:
            print("[AVISO] No se encontraron URLs válidas en el DataFrame.")
            return

        output_file_url = os.path.join(sector_directory_seleccionar, "urls_localidades.csv")
        pd.Series(urls).to_csv(output_file_url, index=False, header=False, encoding='utf-8')
        print(f"Archivo CSV de URLs generado: {output_file_url}")
        print(f"   Total de URLs exportadas: {len(urls)}")
    except Exception as e:
        print(f"Error al escribir el CSV de URLs: {e}")


# ============================
# 3. PROCESO PRINCIPAL
# ============================

print("\n" + "="*60)
print("GENERADOR DE FICHAS POR LOCALIDADES")
print("="*60 + "\n")

# Leer datos del archivo maestro de localidades
print(f"Leyendo archivo de localidades: {ARCHIVO_EXCEL_LOCALIDADES}")
df_localidades = pd.read_excel(ARCHIVO_EXCEL_LOCALIDADES)

total_localidades = len(df_localidades)
print(f"Total de localidades en el archivo: {total_localidades}")

# Aplicar límite si NUM_FICHAS_PRUEBA está configurado
if NUM_FICHAS_PRUEBA and NUM_FICHAS_PRUEBA > 0:
    df_localidades = df_localidades.head(NUM_FICHAS_PRUEBA)
    print(f"[MODO PRUEBA] Generando solo las primeras {NUM_FICHAS_PRUEBA} fichas")
    print(f"[CONSEJO] Para generar todas las fichas ({total_localidades}), cambia NUM_FICHAS_PRUEBA = None en la linea 25\n")
else:
    print(f"[MODO COMPLETO] Generando TODAS las {total_localidades} fichas\n")

if VERBOSE:
    print(f"Columnas disponibles: {df_localidades.columns.tolist()}\n")

# Generar plantilla específica para localidades
if VERBOSE:
    print("Generando plantilla para localidades...")
doc_localidad = generar_plantilla_localidad(ARCHIVO_EXCEL_LOCALIDADES)

# Crear directorio para documentos del sector
sector_directory = os.path.join(PLANTILLAS_CREADAS_DIR, eliminar_acentos_slash(SECTOR))
os.makedirs(sector_directory, exist_ok=True)

# Guardar CSV con los datos procesados (se actualizará después con las URLs)
csv_output = os.path.join(OUTPUT_DIRECTORY_CSV, 'Generador_csv_localidades.csv')

# Contador de documentos generados
contador = 0
errores = 0

print("Generando documentos por localidad...\n")

# Generar documentos para cada localidad
for index, fila in df_localidades.iterrows():
    try:
        # Obtener y limpiar datos
        localidad = str(fila['localidad']).strip() if pd.notna(fila['localidad']) else ''
        provincia = str(fila['provincia']).strip() if pd.notna(fila['provincia']) else ''
        
        if not localidad or not provincia:
            print(f"[AVISO] Fila {index}: Localidad o provincia vacia, saltando...")
            errores += 1
            continue
        
        # Generar URL correcta para esta localidad
        url_generada = generar_url_localidad(localidad, provincia)
        if url_generada:
            # Actualizar el DataFrame con la URL generada
            if 'url' not in df_localidades.columns:
                df_localidades['url'] = ''
            df_localidades.at[index, 'url'] = url_generada
        
        # Capitalizar nombres correctamente
        localidad_formateada = capitalizar_localidad(localidad)
        provincia_formateada = capitalizar_localidad(provincia)
        
        # Preparar contexto con los datos de la localidad
        context = {
            'Localidad': localidad_formateada,
            'Provincia': provincia_formateada,
            'registros': '{:.0f}'.format(fila['registros']) if pd.notna(fila['registros']) else '0',
            'dir': '{:.0f}'.format(fila['Direccion']) if pd.notna(fila['Direccion']) else '0',
            'phone': '{:.0f}'.format(fila['Telefono']) if pd.notna(fila['Telefono']) else '0',
            'mail': '{:.0f}'.format(fila['Mail']) if pd.notna(fila['Mail']) else '0',
            'cat': '{:.0f}'.format(fila['Category']) if pd.notna(fila['Category']) else '0',
            'web': '{:.0f}'.format(fila['Website']) if pd.notna(fila['Website']) else '0',
            'Precio': '{:.2f}'.format(fila['precio']) if pd.notna(fila['precio']) else '0.00'
        }
        
        # Nombre del archivo: Sector_Localidad_de_Provincia.docx
        localidad_archivo = eliminar_acentos_slash(localidad)
        provincia_archivo = eliminar_acentos_slash(provincia)
        sector_prefix = eliminar_acentos_slash(SECTOR).strip('_')
        nombre_archivo = f'{sector_prefix}_{localidad_archivo}_de_{provincia_archivo}.docx'
        
        # Generar el documento
        generar_documento_localidad(context, doc_localidad, nombre_archivo, sector_directory)
        contador += 1
        
        # Mostrar progreso cada 50 documentos (o cada 5 si son menos de 50)
        intervalo_progreso = 5 if len(df_localidades) <= 50 else 50
        if contador % intervalo_progreso == 0:
            if VERBOSE:
                print(f"   [{contador}/{len(df_localidades)}] documentos generados...")
        
    except Exception as e:
        print(f"[ERROR] Error procesando fila {index} ({localidad} - {provincia}): {e}")
        errores += 1
        continue

print("\n" + "="*60)
print(f"[COMPLETADO] Proceso completado!")
print(f"   Documentos generados: {contador}")
print(f"   Errores: {errores}")
print(f"   Total procesado: {contador + errores}")
print("="*60 + "\n")

# Guardar CSV con los datos procesados (incluye URLs generadas)
df_localidades.to_csv(csv_output, index=False, encoding='utf-8')
print(f"CSV de localidades guardado en: {csv_output}\n")

# Generar CSV con URLs (generadas desde localidad y provincia)
print("Generando CSV con URLs...")
# Crear DataFrame solo con las localidades procesadas exitosamente
df_procesadas = df_localidades[df_localidades['url'].notna() & (df_localidades['url'] != '')].copy()
generador_csv_con_url_localidades(sector_directory, df_procesadas)

print("\nProceso finalizado con exito!")

