import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import os
import unicodedata
import threading 
import datetime
import csv
import time

# --- Configuración ---
RUTA_EXCEL_DEFECTO = r"D:\Inplantalariak\01_Inplantalari\00_Recursos Inplantalariak\001_Soft AUTOMATIZACIONES\Gestion de notas_tiempo_docs\Buscador de soluciones ticnologicas\F_Visualizar_Aplicaciones.xlsx"
df = None
loading_thread = None
ultimo_filtrado = None

# --- Columnas a preseleccionar y mostrar por defecto ---
COLUMNAS_PRESELECCIONADAS_POR_DEFECTO = [] 

# --- Lista global para almacenar las líneas seleccionadas de forma persistente ---
lineas_seleccionadas = []

# --- Variables para control de progreso ---
carga_activa = False
progreso_actual = 0

# --- Funciones de Utilidad ---

def normalize_text(text):
    """Normaliza el texto quitando acentos y convirtiendo a minúsculas."""
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    return text

def actualizar_progreso(paso, total, mensaje=""):
    """Actualiza la barra de progreso desde el hilo secundario"""
    global progreso_actual
    progreso_actual = (paso / total) * 100
    ventana.after(0, lambda p=progreso_actual: progress_bar.config(value=p))
    if mensaje:
        ventana.after(0, lambda m=mensaje: loading_label.config(text=m))
    time.sleep(0.1)

def load_excel_data_in_thread(file_path):
    """Función para cargar los datos de Excel en un hilo separado."""
    global df, carga_activa
    carga_activa = True
    
    try:
        if not os.path.exists(file_path):
            ventana.after(0, lambda: on_excel_load_complete(False, file_path, error="Archivo no encontrado."))
            return
        
        try:
            actualizar_progreso(1, 5, "Iniciando carga del archivo...")
            actualizar_progreso(2, 5, "Leyendo datos Excel...")
            temp_df = pd.read_excel(file_path) 
            actualizar_progreso(3, 5, "Procesando datos...")
            df = temp_df
            actualizar_progreso(4, 5, "Preparando interfaz...")
            time.sleep(0.5)
            actualizar_progreso(5, 5, "¡Carga completada!")
            time.sleep(0.5)
            
            ventana.after(0, lambda: on_excel_load_complete(True, file_path))
            
        except Exception as e:
            ventana.after(0, lambda: on_excel_load_complete(False, file_path, error=f"Error: {e}"))
        
    except Exception as e:
        ventana.after(0, lambda: on_excel_load_complete(False, file_path, error=f"Error al iniciar carga: {e}"))
    finally:
        carga_activa = False

def on_excel_load_complete(success, file_path, error=None):
    """Callback que se ejecuta en el hilo principal cuando la carga de Excel termina."""
    global loading_thread
    
    progress_bar['value'] = 100 if success else 0
    ventana.after(100, lambda: progress_bar.grid_remove())
    loading_thread = None

    if success:
        excel_path_label.config(text=f"Archivo cargado: {os.path.basename(file_path)}", foreground="green")
        loading_label.config(text="¡Listo para buscar! Escribe en el cuadro de búsqueda...")
        resultado_texto.delete("1.0", tk.END)
        resultado_texto.insert(tk.END, "✅ Archivo cargado correctamente. Escribe algo para buscar...")
        update_column_selector()
    else:
        excel_path_label.config(text="¡Error al cargar el archivo!", foreground="red")
        messagebox.showerror("Error de Carga", f"Error al cargar el archivo Excel:\n{error}")
        loading_label.config(text="Error al cargar. Intenta de nuevo.", foreground="red")
        column_listbox.delete(0, tk.END)

    boton_cargar_excel.config(state=tk.NORMAL)

def update_column_selector():
    """Actualiza el Listbox con las columnas del DataFrame cargado y preselecciona."""
    column_listbox.delete(0, tk.END)
    if df is not None:
        for col in df.columns:
            column_listbox.insert(tk.END, col)
        loading_label.config(text="¡Columnas cargadas! Selecciona las que deseas ver.")
    else:
        column_listbox.insert(tk.END, "Carga un archivo Excel para ver las columnas.")

def open_file_dialog_async():
    """Permite al usuario seleccionar un archivo Excel y lo carga en un hilo separado."""
    global loading_thread
    if loading_thread and loading_thread.is_alive():
        messagebox.showwarning("Advertencia", "Ya hay un archivo cargándose. Por favor, espera.")
        return

    file_path = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    if file_path:
        excel_path_label.config(text=f"Cargando: {os.path.basename(file_path)}...", foreground="blue")
        loading_label.config(text="Iniciando carga...", foreground="blue")
        progress_bar['value'] = 0
        progress_bar.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        boton_cargar_excel.config(state=tk.DISABLED)
        column_listbox.delete(0, tk.END)
        column_listbox.insert(tk.END, "Cargando columnas...")
        loading_thread = threading.Thread(target=load_excel_data_in_thread, args=(file_path,), daemon=True)
        loading_thread.start()

# --- Funciones de Búsqueda ---

def buscar_automatico(event=None):
    """Búsqueda automática mientras se escribe (con delay)"""
    if hasattr(ventana, '_busqueda_id'):
        ventana.after_cancel(ventana._busqueda_id)
    ventana._busqueda_id = ventana.after(500, realizar_busqueda)

def realizar_busqueda():
    """Realiza la búsqueda con los parámetros actuales"""
    global ultimo_filtrado
    
    if df is None:
        mostrar_resultados_texto("❌ Primero carga un archivo Excel.")
        resultado_count_label.config(text="Esperando datos...")
        return
        
    if not entrada.get().strip():
        mostrar_resultados_texto("Introduce un término de búsqueda...")
        resultado_count_label.config(text="Esperando búsqueda...")
        return

    palabra_original = entrada.get().strip()
    start_time = time.time()
    palabra_normalizada = normalize_text(palabra_original)

    # Obtener las columnas seleccionadas
    selected_indices = column_listbox.curselection()
    selected_columns = [column_listbox.get(i) for i in selected_indices]

    if not selected_columns:
        mostrar_resultados_texto("❌ Selecciona al menos una columna para buscar.")
        resultado_count_label.config(text="Sin columnas seleccionadas")
        return

    # Asegurarse de que las columnas seleccionadas existen
    valid_selected_columns = [col for col in selected_columns if col in df.columns]

    if not valid_selected_columns:
        mostrar_resultados_texto("❌ No se encontraron columnas válidas.")
        resultado_count_label.config(text="Columnas no válidas")
        return
    
    loading_label.config(text="Buscando...", foreground="blue")
    
    # Búsqueda en TODAS las columnas
    mask = df.apply(lambda row: any(palabra_normalizada in normalize_text(str(val)) for val in row), axis=1)
    filtrado = df[mask]
    ultimo_filtrado = filtrado

    elapsed_time = time.time() - start_time

    # MOSTRAR RESULTADOS
    if not filtrado.empty:
        mostrar_resultados_tabulados(filtrado, valid_selected_columns, elapsed_time)
    else:
        mostrar_resultados_texto("❌ No se encontraron coincidencias.")
        resultado_count_label.config(text=f"❌ 0 resultados | {elapsed_time:.2f}s")
    
    loading_label.config(text="Búsqueda completada", foreground="green")

def limpiar_treeview():
    """Limpia completamente el Treeview"""
    for widget in resultado_frame.winfo_children():
        widget.destroy()

def mostrar_resultados_texto(mensaje):
    """Muestra mensajes de texto en el área de resultados"""
    limpiar_treeview()
    resultado_texto.delete("1.0", tk.END)
    resultado_texto.insert(tk.END, mensaje)
    resultado_texto.pack(fill=tk.BOTH, expand=True)

def mostrar_resultados_tabulados(filtrado, columnas, tiempo_ejecucion):
    """Muestra los resultados en formato tabular con columnas reales"""
    limpiar_treeview()
    resultado_texto.pack_forget()  # Ocultar el área de texto
    
    # Crear frame contenedor para Treeview y scrollbar
    tree_container = ttk.Frame(resultado_frame)
    tree_container.pack(fill=tk.BOTH, expand=True)
    
    # Crear Treeview para mostrar datos en columnas
    tree = ttk.Treeview(tree_container, columns=columnas, show='headings', height=min(20, len(filtrado)))
    
    # Configurar columnas con auto-ajuste
    for col in columnas:
        tree.heading(col, text=col)
        # Ancho flexible basado en el contenido
        tree.column(col, width=120, anchor='w', minwidth=80, stretch=True)
    
    # Insertar datos
    for _, row in filtrado.iterrows():
        values = [str(row[col]) if pd.notna(row[col]) else "" for col in columnas]
        tree.insert('', tk.END, values=values)
    
    # Scrollbars
    v_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
    h_scrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
    
    # Empaquetar con grid para mejor control
    tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
    v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
    
    # Configurar pesos del grid
    tree_container.grid_rowconfigure(0, weight=1)
    tree_container.grid_columnconfigure(0, weight=1)
    
    # Configurar evento de clic para selección
    tree.bind("<Button-1>", lambda event: on_treeview_click(event, tree, columnas))
    
    resultado_count_label.config(text=f"✅ {len(filtrado)} resultados | {tiempo_ejecucion:.2f}s")

def on_treeview_click(event, tree, columnas):
    """Maneja el clic en el Treeview para seleccionar filas completas"""
    item = tree.identify_row(event.y)
    if item:
        valores = tree.item(item, 'values')
        linea_texto = " | ".join([f"{col}: {val}" for col, val in zip(columnas, valores)])
        
        if linea_texto not in lineas_seleccionadas:
            lineas_seleccionadas.append(linea_texto)
            actualizar_lista_seleccion()
            status_label.config(text="✅ Línea añadida a la selección", foreground="green")
            ventana.after(2000, lambda: status_label.config(text="", foreground="black"))

# --- Funciones de Gestión de Selección ---

def limpiar_campos():
    """Limpia el campo de entrada y el área de resultados."""
    global lineas_seleccionadas
    entrada.delete(0, tk.END)
    limpiar_treeview()
    resultado_texto.pack(fill=tk.BOTH, expand=True)
    resultado_texto.delete("1.0", tk.END)
    resultado_texto.insert(tk.END, "Listo para una nueva búsqueda...")
    resultado_count_label.config(text="Resultados encontrados: 0")
    
    if lineas_seleccionadas and messagebox.askyesno("Limpiar", "¿Quieres limpiar también la selección de líneas?"):
        lineas_seleccionadas = []
        actualizar_lista_seleccion()

def limpiar_seleccion():
    """Limpia la lista de líneas seleccionadas."""
    global lineas_seleccionadas
    if lineas_seleccionadas and messagebox.askyesno("Limpiar Selección", f"¿Eliminar {len(lineas_seleccionadas)} líneas?"):
        lineas_seleccionadas = []
        actualizar_lista_seleccion()

def actualizar_lista_seleccion():
    """Actualiza el Listbox de selección con las líneas almacenadas."""
    selection_listbox.delete(0, tk.END)
    for linea in lineas_seleccionadas:
        display_linea = linea[:97] + "..." if len(linea) > 100 else linea
        selection_listbox.insert(tk.END, display_linea)
    
    selection_count_label.config(text=f"Líneas seleccionadas: {len(lineas_seleccionadas)}")
    boton_exportar_seleccion.config(state=tk.NORMAL if lineas_seleccionadas else tk.DISABLED)
    boton_copiar_seleccion.config(state=tk.NORMAL if lineas_seleccionadas else tk.DISABLED)

def eliminar_linea_seleccionada():
    """Elimina la línea seleccionada de la lista de selección."""
    global lineas_seleccionadas
    selected_indices = selection_listbox.curselection()
    if selected_indices:
        index = selected_indices[0]
        lineas_seleccionadas.pop(index)
        actualizar_lista_seleccion()
        status_label.config(text="Línea eliminada", foreground="orange")
        ventana.after(2000, lambda: status_label.config(text="", foreground="black"))

# --- Funciones de Exportación ---

def obtener_columnas_seleccionadas():
    """Obtiene las columnas seleccionadas actualmente"""
    selected_indices = column_listbox.curselection()
    return [column_listbox.get(i) for i in selected_indices] if selected_indices else []

def exportar_seleccion():
    """Exporta SOLO las columnas seleccionadas a Excel/CSV"""
    if not lineas_seleccionadas:
        messagebox.showwarning("Exportar", "No hay líneas seleccionadas para exportar.")
        return

    selected_columns = obtener_columnas_seleccionadas()
    if not selected_columns:
        messagebox.showwarning("Exportar", "Selecciona al menos una columna para exportar.")
        return

    file_path = filedialog.asksaveasfilename(
        title="Exportar selección",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )
    
    if file_path:
        try:
            loading_label.config(text="Exportando...", foreground="blue")
            
            # Reconstruir DataFrame solo con las columnas seleccionadas
            datos_exportar = []
            for linea in lineas_seleccionadas:
                datos_fila = {}
                partes = linea.split(' | ')
                for parte in partes:
                    if ': ' in parte:
                        col, val = parte.split(': ', 1)
                        if col in selected_columns:
                            datos_fila[col] = val
                for col in selected_columns:
                    if col not in datos_fila:
                        datos_fila[col] = ""
                datos_exportar.append(datos_fila)
            
            export_df = pd.DataFrame(datos_exportar)
            export_df = export_df[selected_columns]
            
            if file_path.endswith('.xlsx'):
                export_df.to_excel(file_path, index=False)
            else:
                export_df.to_csv(file_path, index=False, encoding='utf-8')
            
            messagebox.showinfo("Éxito", f"✅ {len(lineas_seleccionadas)} líneas exportadas")
            loading_label.config(text="Exportación completada", foreground="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {e}")
            loading_label.config(text="Error en exportación", foreground="red")

def copiar_seleccion_portapapeles():
    """Copia SOLO las columnas seleccionadas al portapapeles (sin encabezados)"""
    if not lineas_seleccionadas:
        messagebox.showwarning("Copiar", "No hay líneas seleccionadas para copiar.")
        return
    
    selected_columns = obtener_columnas_seleccionadas()
    if not selected_columns:
        messagebox.showwarning("Copiar", "Selecciona al menos una columna para copiar.")
        return
    
    # Extraer solo los valores de las columnas seleccionadas
    texto_a_copiar = ""
    for linea in lineas_seleccionadas:
        valores_fila = []
        partes = linea.split(' | ')
        for parte in partes:
            if ': ' in parte:
                col, val = parte.split(': ', 1)
                if col in selected_columns:
                    valores_fila.append(val)
        texto_a_copiar += "\t".join(valores_fila) + "\n"
    
    ventana.clipboard_clear()
    ventana.clipboard_append(texto_a_copiar.strip())
    status_label.config(text=f"📋 {len(lineas_seleccionadas)} líneas copiadas", foreground="blue")
    ventana.after(3000, lambda: status_label.config(text="", foreground="black"))

def seleccionar_todo_resultados():
    """Selecciona todas las filas visibles en los resultados"""
    if ultimo_filtrado is not None and not ultimo_filtrado.empty:
        selected_columns = obtener_columnas_seleccionadas()
        if selected_columns:
            nuevas_lineas = 0
            for _, row in ultimo_filtrado.iterrows():
                linea_texto = " | ".join([f"{col}: {str(row[col]) if pd.notna(row[col]) else ''}" for col in selected_columns])
                if linea_texto not in lineas_seleccionadas:
                    lineas_seleccionadas.append(linea_texto)
                    nuevas_lineas += 1
            actualizar_lista_seleccion()
            status_label.config(text=f"✅ {nuevas_lineas} líneas añadidas", foreground="green")
            ventana.after(3000, lambda: status_label.config(text="", foreground="black"))

# --- INTERFAZ RESPONSIVE MEJORADA ---

ventana = tk.Tk()
ventana.title("Buscador de Aplicaciones TIC v3.5 - RESPONSIVE")
ventana.geometry("1400x800")  # Tamaño inicial más grande
ventana.minsize(1000, 600)   # Tamaño mínimo
ventana.resizable(True, True)

# Configurar estilo moderno
style = ttk.Style()
style.theme_use('clam')
style.configure('TFrame', background='#f0f0f0')
style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
style.configure('TButton', font=('Arial', 10, 'bold'), padding=6)
style.configure('TEntry', padding=6, font=('Arial', 10))
style.configure('Title.TLabel', font=('Arial', 14, 'bold'), foreground='#2c3e50')
style.configure('Subtitle.TLabel', font=('Arial', 11, 'bold'), foreground='#34495e')
style.configure('Treeview', font=('Arial', 9), rowheight=25)
style.configure('Treeview.Heading', font=('Arial', 9, 'bold'), background='#34495e', foreground='white')
style.configure('TLabelframe', font=('Arial', 10, 'bold'), background='#f0f0f0')
style.configure('TLabelframe.Label', font=('Arial', 10, 'bold'), foreground='#2c3e50')

# Configurar grid principal para responsive
ventana.grid_rowconfigure(0, weight=1)
ventana.grid_columnconfigure(0, weight=1)

# Frame principal con padding
main_frame = ttk.Frame(ventana, padding="10")
main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
main_frame.grid_rowconfigure(1, weight=1)  # La fila de resultados se expande
main_frame.grid_columnconfigure(0, weight=1)

# --- MARCO SUPERIOR RESPONSIVE ---
control_frame = ttk.LabelFrame(main_frame, text="CONTROLES PRINCIPALES", padding="15")
control_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(0, 10))
control_frame.grid_columnconfigure(1, weight=1)  # Columna central expansible

# Título
titulo_label = ttk.Label(control_frame, text="🔍 BUSCADOR AVANZADO TIC v3.5 - DISEÑO RESPONSIVE", style='Title.TLabel')
titulo_label.grid(row=0, column=0, columnspan=4, pady=(0, 15), sticky=tk.W)

# Información de archivo
excel_path_label = ttk.Label(control_frame, text="Esperando cargar archivo...", foreground="gray", font=('Arial', 9))
excel_path_label.grid(row=1, column=0, columnspan=4, pady=(0, 5), sticky=tk.W)

loading_label = ttk.Label(control_frame, text="", foreground="blue", font=('Arial', 9, 'italic'))
loading_label.grid(row=2, column=0, columnspan=4, pady=(0, 10), sticky=tk.W)

# Fila 1: Botones y progreso
boton_cargar_excel = ttk.Button(control_frame, text="📂 Cargar Excel", command=open_file_dialog_async)
boton_cargar_excel.grid(row=3, column=0, padx=(0, 10), pady=5, sticky=tk.W)

progress_bar = ttk.Progressbar(control_frame, orient="horizontal", mode="determinate", length=200)
progress_bar.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
progress_bar.grid_remove()

# Fila 2: Búsqueda
label_buscar = ttk.Label(control_frame, text="🔍 Término a buscar:", style='Subtitle.TLabel')
label_buscar.grid(row=4, column=0, sticky=tk.W, pady=5)

entrada = ttk.Entry(control_frame, width=35)
entrada.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
entrada.bind("<KeyRelease>", buscar_automatico)
entrada.bind("<Return>", lambda e: seleccionar_todo_resultados())

boton_limpiar = ttk.Button(control_frame, text="🧹 Limpiar", command=limpiar_campos)
boton_limpiar.grid(row=4, column=2, padx=5, pady=5, sticky=tk.W)

# Fila 3: Acciones rápidas
acciones_frame = ttk.Frame(control_frame)
acciones_frame.grid(row=5, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))

boton_seleccionar_todo = ttk.Button(acciones_frame, text="⭐ Seleccionar Todo Visible", command=seleccionar_todo_resultados)
boton_seleccionar_todo.pack(side=tk.LEFT, padx=(0, 10))

status_label = ttk.Label(control_frame, text="", font=('Arial', 9), foreground="green")
status_label.grid(row=6, column=0, columnspan=4, pady=(10, 0), sticky=tk.W)

# --- MARCO DE RESULTADOS COMPLETAMENTE RESPONSIVE ---
results_frame = ttk.Frame(main_frame)
results_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

# Configurar grid responsive para las 3 columnas
results_frame.grid_rowconfigure(0, weight=1)
results_frame.grid_columnconfigure(0, weight=1)  # Columna 1: Selector (ancho flexible)
results_frame.grid_columnconfigure(1, weight=3)  # Columna 2: Resultados (más ancha)
results_frame.grid_columnconfigure(2, weight=1)  # Columna 3: Selección (ancho flexible)

# Columna 1 - Selector de columnas (RESPONSIVE)
column_selector_frame = ttk.LabelFrame(results_frame, text="📊 COLUMNAS DISPONIBLES", padding="8")
column_selector_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=(0, 5))
column_selector_frame.grid_rowconfigure(0, weight=1)
column_selector_frame.grid_columnconfigure(0, weight=1)

# Frame interno con scrollbar
column_inner_frame = ttk.Frame(column_selector_frame)
column_inner_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
column_inner_frame.grid_rowconfigure(0, weight=1)
column_inner_frame.grid_columnconfigure(0, weight=1)

column_listbox = tk.Listbox(column_inner_frame, selectmode=tk.MULTIPLE, exportselection=False, font=("Arial", 9))
column_listbox.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

listbox_scrollbar = ttk.Scrollbar(column_inner_frame, orient="vertical", command=column_listbox.yview)
listbox_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
column_listbox.config(yscrollcommand=listbox_scrollbar.set)

# Columna 2 - Resultados de búsqueda (RESPONSIVE)
resultados_container = ttk.LabelFrame(results_frame, text="🔍 RESULTADOS DE BÚSQUEDA", padding="8")
resultados_container.grid(row=0, column=1, padx=5, sticky=(tk.N, tk.S, tk.E, tk.W))
resultados_container.grid_rowconfigure(1, weight=1)  # El área de resultados se expande
resultados_container.grid_columnconfigure(0, weight=1)

resultado_count_label = ttk.Label(resultados_container, text="Resultados encontrados: 0", 
                                 font=('Arial', 10, 'bold'), foreground="darkgreen")
resultado_count_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))

# Frame para resultados (alterna entre Treeview y texto)
resultado_frame = ttk.Frame(resultados_container)
resultado_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
resultado_frame.grid_rowconfigure(0, weight=1)
resultado_frame.grid_columnconfigure(0, weight=1)

# Área de texto para mensajes
resultado_texto = scrolledtext.ScrolledText(resultado_frame, wrap=tk.WORD, font=("Consolas", 10), 
                                          relief="sunken", borderwidth=1)
resultado_texto.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
resultado_texto.insert(tk.END, "Escribe algo para buscar...")

# Columna 3 - Selección (RESPONSIVE)
selection_frame = ttk.LabelFrame(results_frame, text="⭐ LÍNEAS SELECCIONADAS", padding="8")
selection_frame.grid(row=0, column=2, sticky=(tk.N, tk.S, tk.E, tk.W), padx=(5, 0))
selection_frame.grid_rowconfigure(1, weight=1)  # Listbox se expande
selection_frame.grid_columnconfigure(0, weight=1)

# Header con contador
selection_header = ttk.Frame(selection_frame)
selection_header.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))

selection_count_label = ttk.Label(selection_header, text="Líneas: 0", 
                                 font=('Arial', 10, 'bold'), foreground="darkred")
selection_count_label.pack(side=tk.LEFT)

# Botones de acción en grid responsive
botones_seleccion_frame = ttk.Frame(selection_frame)
botones_seleccion_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

# Botones en grid para mejor distribución
boton_exportar_seleccion = ttk.Button(botones_seleccion_frame, text="💾 Exportar", 
                                     command=exportar_seleccion, state=tk.DISABLED)
boton_exportar_seleccion.grid(row=0, column=0, padx=(0, 3), pady=2, sticky=tk.W)

boton_copiar_seleccion = ttk.Button(botones_seleccion_frame, text="📋 Copiar", 
                                   command=copiar_seleccion_portapapeles, state=tk.DISABLED)
boton_copiar_seleccion.grid(row=0, column=1, padx=3, pady=2, sticky=tk.W)

boton_eliminar_linea = ttk.Button(botones_seleccion_frame, text="🗑️ Quitar", 
                                 command=eliminar_linea_seleccionada)
boton_eliminar_linea.grid(row=1, column=0, padx=(0, 3), pady=2, sticky=tk.W)

boton_limpiar_seleccion = ttk.Button(botones_seleccion_frame, text="🧹 Limpiar", 
                                    command=limpiar_seleccion)
boton_limpiar_seleccion.grid(row=1, column=1, padx=3, pady=2, sticky=tk.W)

# Configurar grid de botones
botones_seleccion_frame.grid_columnconfigure(0, weight=1)
botones_seleccion_frame.grid_columnconfigure(1, weight=1)

# Listbox de selección con scrollbar
selection_inner_frame = ttk.Frame(selection_frame)
selection_inner_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
selection_inner_frame.grid_rowconfigure(0, weight=1)
selection_inner_frame.grid_columnconfigure(0, weight=1)

selection_listbox = tk.Listbox(selection_inner_frame, font=("Arial", 9))
selection_listbox.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

selection_scrollbar = ttk.Scrollbar(selection_inner_frame, orient="vertical", command=selection_listbox.yview)
selection_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
selection_listbox.config(yscrollcommand=selection_scrollbar.set)

# --- CONFIGURACIÓN FINAL RESPONSIVE ---

# Asegurar que todos los frames internos se expandan
column_selector_frame.grid_propagate(False)
resultados_container.grid_propagate(False)
selection_frame.grid_propagate(False)

# Cargar archivo por defecto al iniciar
if os.path.exists(RUTA_EXCEL_DEFECTO):
    excel_path_label.config(text=f"Cargando archivo predeterminado...", foreground="blue")
    loading_label.config(text="Preparando carga...", foreground="blue")
    progress_bar['value'] = 0
    progress_bar.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
    boton_cargar_excel.config(state=tk.DISABLED)
    column_listbox.delete(0, tk.END)
    column_listbox.insert(tk.END, "Cargando columnas...")
    loading_thread = threading.Thread(target=load_excel_data_in_thread, args=(RUTA_EXCEL_DEFECTO,), daemon=True)
    loading_thread.start()
else:
    excel_path_label.config(text="¡Archivo predeterminado no encontrado! Carga uno manualmente.", foreground="red")
    loading_label.config(text="Usa el botón 'Cargar Excel' para empezar")

# Ajustar inicial
ventana.update()
ventana.after(100, lambda: ventana.minsize(ventana.winfo_width(), ventana.winfo_height()))

ventana.mainloop()