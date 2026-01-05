import pandas as pd
import subprocess
import os

# --- Configuración ---
# Asegúrate de que este sea el nombre de tu archivo Excel
NOMBRE_ARCHIVO_EXCEL = 'programasInicioIracing.xlsx'

# Nombres de las columnas en tu archivo Excel (columna A y B)
COLUMNA_RUTA = 'Ruta'
COLUMNA_EJECUTABLE = 'Ejecutable'

# --- Función Principal ---

def ejecutar_programas_desde_excel(nombre_archivo):
    """
    Lee un archivo Excel con rutas de programas y los ejecuta secuencialmente.

    Args:
        nombre_archivo (str): El nombre del archivo Excel a leer.
    """
    print(f"✅ Leyendo el archivo Excel: {nombre_archivo}...")

    try:
        # Cargar el archivo Excel en un DataFrame
        # 'header=None' si no tienes encabezados y quieres usar índices
        # Pero es mejor usar los nombres de columna si existen.
        df = pd.read_excel(nombre_archivo)
        
    except FileNotFoundError:
        print(f"❌ ERROR: El archivo '{nombre_archivo}' no fue encontrado.")
        return
    except Exception as e:
        print(f"❌ ERROR al leer el archivo Excel: {e}")
        return

    # Comprobar si las columnas requeridas existen
    if COLUMNA_RUTA not in df.columns or COLUMNA_EJECUTABLE not in df.columns:
        print(f"❌ ERROR: El archivo Excel debe contener las columnas '{COLUMNA_RUTA}' (Columna A) y '{COLUMNA_EJECUTABLE}' (Columna B).")
        return

    print(f"📚 {len(df)} programas encontrados en la lista.")
    print("-" * 30)

    # Iterar sobre cada fila del DataFrame
    for index, row in df.iterrows():
        # Obtener los valores de las columnas
        ruta_directorio = str(row[COLUMNA_RUTA]).strip()
        nombre_ejecutable = str(row[COLUMNA_EJECUTABLE]).strip()
        
        # Combinar la ruta del directorio y el nombre del ejecutable
        ruta_completa_programa = os.path.join(ruta_directorio, nombre_ejecutable)
        
        print(f"▶️ Ejecutando programa {index + 1}: {ruta_completa_programa}")

        try:
            # Usar subprocess.Popen para ejecutar el programa. 
            # Esto permite que el script continúe después de iniciar el programa.
            # Si quieres que espere a que cada programa termine, usa subprocess.run()
            # Popen es mejor para iniciar varias aplicaciones de escritorio.
            
            # Nota importante: En Windows, para iniciar ejecutables (.exe) que 
            # están fuera del directorio de trabajo, a menudo se usa shell=True 
            # o se pasa la ruta completa directamente.
            
            subprocess.Popen(ruta_completa_programa, shell=True)
            print(f"   Iniciado con éxito.")
            
        except FileNotFoundError:
            print(f"   ⚠️ ADVERTENCIA: Programa no encontrado en la ruta: {ruta_completa_programa}")
        except Exception as e:
            print(f"   ❌ ERROR al intentar ejecutar: {e}")
        
        print("-" * 30)

    print("🏁 Fin del script. Todos los programas han sido iniciados (o se intentó iniciarlos).")

# --- Ejecución ---
if __name__ == "__main__":
    ejecutar_programas_desde_excel(NOMBRE_ARCHIVO_EXCEL)