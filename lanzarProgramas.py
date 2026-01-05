import openpyxl
import subprocess
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import psutil

def lanzar_programas():
    # 1. Seleccionar el archivo Excel
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecciona el Excel de configuración",
        filetypes=[("Archivos de Excel", "*.xlsx")]
    )

    if not file_path:
        return

    try:
        # 2. Cargar el libro de Excel
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active # Lee la hoja activa

        # 3. Recorrer filas empezando desde la fila 2 (min_row=2)
        # Columna 1 (A): Ruta | Columna 2 (B): Ejecutable
        lanzados = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            ruta = str(row[0]).strip() if row[0] else None
            ejecutable = str(row[1]).strip() if row[1] else None

            # Si la fila está vacía, saltar
            if not ruta or not ejecutable or ruta == "None":
                continue

            full_path = os.path.join(ruta, ejecutable)
            
            if os.path.exists(full_path):
                print(f"Iniciando: {ejecutable}")
                subprocess.Popen([full_path], cwd=ruta, shell=True)
                lanzados += 1
                time.sleep(3) # Espera de 3 segundos entre programas
            else:
                print(f"No se encontró: {full_path}")

        messagebox.showinfo("Finalizado", f"Se han intentado abrir {lanzados} programas.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al leer el Excel: {e}")

if __name__ == "__main__":
    lanzar_programas()