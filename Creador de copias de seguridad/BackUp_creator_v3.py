import os
import shutil
import json
from tqdm import tqdm
from plyer import notification
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

CONFIG_FILE = "backup_config.json"
LOG_FILE = "backup_log.txt"

def log(mensaje):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ahora}] {mensaje}\n")

def seleccionar_varias_carpetas(titulo):
    carpetas = []
    while True:
        root = tk.Tk()
        root.withdraw()
        carpeta = filedialog.askdirectory(title=titulo)
        root.destroy()
        if carpeta:
            if carpeta not in carpetas:
                carpetas.append(carpeta)
                print(f"Carpeta añadida: {carpeta}")
            else:
                print("Esa carpeta ya está añadida.")
            seguir = input("¿Añadir otra carpeta? (s/n): ")
            if seguir.lower() != "s":
                break
        else:
            break
    return carpetas

def guardar_config(origenes, destino, excluidos):
    conf = {"origenes": origenes, "destino": destino, "excluidos": excluidos}
    with open(CONFIG_FILE, "w") as file:
        json.dump(conf, file)

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as file:
            return json.load(file)
    return None

def sync_mirror(origen, destino, excluidos):
    archivos_a_procesar = []
    archivos_copiados = 0
    archivos_eliminados = 0
    errores = []

    for carpeta, subcarpetas, archivos in os.walk(origen):
        if any(excl in carpeta for excl in excluidos):
            continue
        for archivo in archivos:
            archivos_a_procesar.append(os.path.join(carpeta, archivo))

    pbar = tqdm(total=len(archivos_a_procesar), desc=f"Copiando archivos desde {origen}", unit="archivo")
    for src_file in archivos_a_procesar:
        try:
            rel_path = os.path.relpath(src_file, origen)
            dst_file = os.path.join(destino, os.path.basename(origen), rel_path)
            dst_dir = os.path.dirname(dst_file)
            if not os.path.exists(dst_dir):
                os.makedirs(dst_dir)
            if (not os.path.exists(dst_file) or
                os.path.getmtime(src_file) > os.path.getmtime(dst_file)):
                shutil.copy2(src_file, dst_file)
                archivos_copiados += 1
            pbar.update(1)
        except Exception as e:
            errores.append(f"Error copiando {src_file}: {e}")
    pbar.close()

    dest_base = os.path.join(destino, os.path.basename(origen))
    for carpeta, subcarpetas, archivos in os.walk(dest_base, topdown=False):
        for archivo in archivos:
            dst_file = os.path.join(carpeta, archivo)
            try:
                rel_path = os.path.relpath(dst_file, dest_base)
                src_file = os.path.join(origen, rel_path)
                if not os.path.exists(src_file):
                    os.remove(dst_file)
                    archivos_eliminados += 1
            except Exception as e:
                errores.append(f"Error eliminando {dst_file}: {e}")
        if not os.listdir(carpeta):
            try:
                os.rmdir(carpeta)
            except Exception as e:
                errores.append(f"Error eliminando carpeta {carpeta}: {e}")

    return archivos_copiados, archivos_eliminados, errores

# Carga o pide configuración
config = cargar_config()
if config:
    origenes = config["origenes"]
    destino = config["destino"]
    excluidos = config["excluidos"]

    print(f"Carpetas origen actuales: {origenes}")
    respuesta = input("¿Quieres añadir nuevas carpetas a la copia? (s/n): ")
    if respuesta.lower() == "s":
        nuevas = seleccionar_varias_carpetas("Selecciona carpetas ORIGEN para añadir")
        origenes.extend([x for x in nuevas if x not in origenes])

    print(f"Patrones excluidos actuales: {excluidos}")
    respuesta_exc = input("¿Quieres modificar las exclusiones? (s/n): ")
    if respuesta_exc.lower() == "s":
        excluidos_input = input("Introduce los nombres separados por coma, o deja vacío para no excluir nada: ")
        excluidos = [e.strip() for e in excluidos_input.split(",")] if excluidos_input.strip() else []

else:
    print("Configuración inicial. Selecciona carpetas de origen para backup:")
    origenes = seleccionar_varias_carpetas("Selecciona carpetas ORIGEN")
    destino = filedialog.askdirectory(title="Selecciona la carpeta DESTINO EN DISCO EXTERNO")
    excluidos = []
    print("¿Quieres excluir alguna carpeta o patrón? (ejemplo: venv,node_modules,temp)")
    excluidos_input = input("Introduce los nombres separados por coma, o deja vacío: ")
    if excluidos_input.strip():
        excluidos = [e.strip() for e in excluidos_input.split(",")]

guardar_config(origenes, destino, excluidos)

log(f"INICIO de copia - Orígenes: {origenes} - Destino: {destino} - Exclusiones: {excluidos}")

total_copiados = 0
total_eliminados = 0
all_errores = []

print("Sincronizando y copiando todas las carpetas origen...")

for origen in origenes:
    copiados, eliminados, errores = sync_mirror(origen, destino, excluidos)
    total_copiados += copiados
    total_eliminados += eliminados
    all_errores.extend(errores)
    log(f"Copia desde {origen} - Archivos copiados: {copiados}, eliminados: {eliminados}, errores: {len(errores)}")

mensaje_final = f"Copia finalizada. Archivos copiados: {total_copiados}, eliminados: {total_eliminados}."
if all_errores:
    mensaje_final += f" Errores: {len(all_errores)} (ver log para detalles)."
    for err in all_errores:
        log(err)
else:
    mensaje_final += " Sin errores."

log(mensaje_final)

notification.notify(
    title="Copia de seguridad completada",
    message=mensaje_final,
    timeout=10
)

print(mensaje_final)
print("Registro guardado en", LOG_FILE)
