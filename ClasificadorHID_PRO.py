import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import threading
import queue
import os
import subprocess 

# Configuración y Variables Globales
OUTPUT_FILE = "REGISTRO_HID.xlsx"
SCAN_QUEUE = queue.Queue()
PROCESS_RUNNING = False
STORES = ['DHL', '99MINUTOS', 'FEDEX', 'TERRESTRE', 'BLINK']

# Almacenamiento de datos y contadores
DATA_CACHE = {}
COUNTS = {store: 0 for store in STORES}
TOTAL_SCANS = 0
# Usamos el directorio donde se ejecuta el script para la ruta absoluta del archivo
FILE_PATH = os.path.join(os.getcwd(), OUTPUT_FILE) 

# ----------------- Lógica de Archivos y Datos (Paso a Paso) -----------------

def load_initial_data():
    """Carga los datos existentes del Excel al inicio del programa.
       Paso 1: Intentar leer el archivo si existe."""
    global DATA_CACHE, COUNTS, TOTAL_SCANS
    DATA_CACHE = {}
    COUNTS = {store: 0 for store in STORES}
    TOTAL_SCANS = 0
    
    try:
        if not os.path.exists(FILE_PATH):
            print(f"Archivo {OUTPUT_FILE} no encontrado. Creando nueva caché.")
            return

        xls = pd.ExcelFile(FILE_PATH)
        sheets = xls.sheet_names
        
        for sheet in sheets:
            df = xls.parse(sheet)
            DATA_CACHE[sheet] = df
            
            # Recalcular contadores
            ok_counts = df[df['Status'] == 'OK'].shape[0]
            if sheet in COUNTS:
                COUNTS[sheet] = ok_counts
            TOTAL_SCANS += ok_counts
            
        print(f"Datos cargados exitosamente desde {OUTPUT_FILE}.")
    except Exception as e:
        messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo Excel '{OUTPUT_FILE}'. Asegúrese de que no esté abierto.\nError: {e}")

def save_current_data():
    """Guarda los datos de la caché en el Excel de forma segura.
       Paso 2: Escribir los datos en el archivo Excel."""
    try:
        # Asegurarse de que el directorio exista (debería ser el CWD)
        os.makedirs(os.path.dirname(FILE_PATH) or '.', exist_ok=True)
        
        # Usamos el path absoluto para evitar problemas de ruta
        writer = pd.ExcelWriter(FILE_PATH, engine='xlsxwriter')
        all_sheets = set(STORES).union(DATA_CACHE.keys())
        
        for sheet in sorted(all_sheets):
            # Obtener el DataFrame o crear uno vacío si la tienda es nueva
            df = DATA_CACHE.get(sheet, pd.DataFrame(columns=['Timestamp', 'Code', 'Status']))
            df.to_excel(writer, sheet_name=sheet, index=False)
            
        writer.close()
        print(f"Datos guardados exitosamente en {OUTPUT_FILE}.")
        return True
    except PermissionError:
        messagebox.showerror("Error de Permiso", f"No se puede guardar el archivo '{OUTPUT_FILE}'. Asegúrese de que el archivo NO esté abierto en otra aplicación (Excel).")
        return False
    except Exception as e:
        messagebox.showerror("Error de Escritura", f"Error desconocido al guardar los datos: {e}")
        return False

def open_output_folder():
    """Abre la carpeta que contiene el archivo de Excel."""
    try:
        folder = os.path.dirname(FILE_PATH)
        if not folder:
            folder = os.getcwd() 

        if os.path.exists(folder):
            if os.name == 'nt': # Windows
                subprocess.Popen(['explorer', folder])
            elif os.uname()[0] == 'Darwin': # macOS
                subprocess.Popen(['open', folder])
            else: # Linux/Otros
                subprocess.Popen(['xdg-open', folder])
        else:
            # Si la ruta no existe, abre el directorio de trabajo actual
            subprocess.Popen(['explorer', os.getcwd()])
            
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la carpeta. Error: {e}")


# ----------------- Procesamiento de Scans (Worker Thread) -----------------

def process_worker():
    """Procesa los escaneos de la cola en un hilo separado.
       Paso 3: El hilo se ejecuta en segundo plano para procesar la cola."""
    global TOTAL_SCANS, COUNTS, DATA_CACHE
    
    # 3.1: Carga inicial de datos ANTES de procesar
    load_initial_data() 
    app.after(0, app.update_initial_interface)

    while PROCESS_RUNNING or not SCAN_QUEUE.empty():
        try:
            line = SCAN_QUEUE.get(timeout=0.1)
        except queue.Empty:
            continue

        try:
            # 3.2: Procesar la trama de la cámara (TIENDA,CÓDIGO)
            parts = line.split(',')
            if len(parts) != 2:
                continue

            store_name = parts[0].strip().upper()
            code = parts[1].strip()
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if store_name not in DATA_CACHE:
                DATA_CACHE[store_name] = pd.DataFrame(columns=['Timestamp', 'Code', 'Status'])

            df = DATA_CACHE[store_name]
            
            # 3.3: Lógica de Duplicados (DUP vs OK)
            if code in df['Code'].values:
                status = 'DUP'
            else:
                status = 'OK'
                TOTAL_SCANS += 1
                COUNTS[store_name] = COUNTS.get(store_name, 0) + 1
            
            # Añadir nuevo registro a la caché
            new_row = pd.DataFrame([{'Timestamp': timestamp, 'Code': code, 'Status': status}])
            DATA_CACHE[store_name] = pd.concat([df, new_row], ignore_index=True)
            
            # 3.4: Actualizar GUI desde el hilo principal (after(0, ...))
            app.after(0, app.update_scan_interface, store_name, code, status)
            
        except Exception as e:
            print(f"Error al procesar línea: {e}")
        finally:
            SCAN_QUEUE.task_done()
    
    # Paso 4: Guardar los datos al finalizar el hilo de trabajo
    app.after(0, app.save_button.config, {'state': tk.NORMAL})
    save_current_data()
    app.after(0, app.status_label.config, {'text': "STATUS: DETENIDO - Archivo Excel actualizado.", 'foreground': "red"})


# ----------------- Interfaz Gráfica (Tkinter con Ttk) -----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Clasificador HID PRO V4")
        self.geometry("650x450")
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.input_buffer = ""
        self.count_labels = {}
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind('<FocusIn>', self.handle_focus_in)
        self.bind('<FocusOut>', self.handle_focus_out)
        self.focus_force() 
        self.update_initial_interface() 

    def create_widgets(self):
        # Frame de control principal
        control_frame = ttk.Frame(self, padding="10")
        control_frame.pack(fill='x')
        
        ttk.Label(control_frame, text=f"Archivo de Salida: {OUTPUT_FILE}", font=('Arial', 9)).pack(side=tk.LEFT, padx=10)
        
        # Botones de Acción
        self.start_button = ttk.Button(control_frame, text="INICIAR ESCUCHA (HID)", command=self.start_process)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="DETENER", command=self.stop_process, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        self.save_button = ttk.Button(control_frame, text="GUARDAR AHORA", command=self.manual_save)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.folder_button = ttk.Button(control_frame, text="📂 Abrir Carpeta de Registros", command=open_output_folder)
        self.folder_button.pack(side=tk.RIGHT, padx=5)

        # Separador
        ttk.Separator(self, orient='horizontal').pack(fill='x', pady=5)
        
        # Estado
        self.status_label = ttk.Label(self, text="STATUS: LISTO. Esperando inicio...", foreground="gray", font=('Arial', 12, 'bold'))
        self.status_label.pack(pady=5)
        
        # Último Scan y Foco
        self.last_scan_label = ttk.Label(self, text="ÚLTIMO: N/A", foreground="gray", font=('Arial', 10))
        self.last_scan_label.pack(pady=5)
        
        # Indicador de Foco
        self.focus_indicator = ttk.Label(self, text="⚠️ VENTANA SIN FOCO (NO CAPTURARÁ)", foreground="red", font=('Arial', 9, 'bold'))
        self.focus_indicator.pack(pady=5)
        
        # Totalizador
        self.total_scans_label = ttk.Label(self, text="Total Registros Únicos: 0", font=('Arial', 14, 'bold'))
        self.total_scans_label.pack(pady=10)
        
        # Contadores por Tienda
        counts_frame = ttk.Frame(self)
        counts_frame.pack(pady=10)
        
        for i, store in enumerate(STORES):
            label = ttk.Label(counts_frame, text=f"{store}: 0", width=15, relief=tk.RIDGE, anchor='center', font=('Arial', 10))
            label.grid(row=0, column=i, padx=5, pady=5, ipadx=5, ipady=5)
            self.count_labels[store] = label

    # --- Métodos de Captura ---

    def handle_key_input(self, event):
        """Captura la entrada de teclado (cámara HID)."""
        global PROCESS_RUNNING
        if not PROCESS_RUNNING:
            return

        char = event.char
        
        if char == '\r' or char == '\n':
            if self.input_buffer:
                SCAN_QUEUE.put(self.input_buffer)
                self.input_buffer = ""
        else:
            self.input_buffer += char

    def handle_focus_in(self, event):
        """Maneja cuando la ventana gana el foco."""
        self.focus_indicator.config(text="✅ VENTANA CON FOCO (CAPTURA ACTIVA)", foreground="green")
        self.bind('<Key>', self.handle_key_input)

    def handle_focus_out(self, event):
        """Maneja cuando la ventana pierde el foco."""
        self.focus_indicator.config(text="⚠️ VENTANA SIN FOCO (NO CAPTURARÁ)", foreground="red")
        self.unbind('<Key>')

    # --- Métodos de Control ---

    def start_process(self):
        """Inicia el proceso de escucha y el hilo de procesamiento."""
        global PROCESS_RUNNING, worker_thread
        if PROCESS_RUNNING: return

        PROCESS_RUNNING = True
        self.input_buffer = ""
        self.status_label.config(text="STATUS: ACTIVO - Esperando lectura...", foreground="blue")
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        worker_thread = threading.Thread(target=process_worker)
        worker_thread.daemon = True
        worker_thread.start()
        
        self.focus_force() 

    def stop_process(self):
        """Detiene el proceso de escucha y el hilo de procesamiento."""
        global PROCESS_RUNNING
        if not PROCESS_RUNNING: return

        PROCESS_RUNNING = False
        self.status_label.config(text="STATUS: DETENIENDO - Guardando datos...", foreground="orange")
        self.stop_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        
        self.after(500, self.check_worker_end)

    def check_worker_end(self):
        global worker_thread
        if worker_thread and worker_thread.is_alive():
            self.after(500, self.check_worker_end)
        else:
            self.status_label.config(text="STATUS: DETENIDO - Proceso finalizado.", foreground="red")
            self.start_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)

    def manual_save(self):
        """Función para guardar los datos de forma manual."""
        if save_current_data():
             self.status_label.config(text="STATUS: Guardado Manual Exitoso.", foreground="darkgreen")

    def on_closing(self):
        """Maneja el evento de cierre de ventana."""
        if PROCESS_RUNNING:
            if messagebox.askyesno("Detener Proceso", "El proceso está activo. ¿Desea detenerlo y guardar antes de salir?"):
                self.stop_process()
                self.after(500, self.destroy)
            else:
                return
        else:
            self.destroy()

    # --- Métodos de Actualización ---

    def update_scan_interface(self, store, code, status):
        """Actualiza la interfaz con el último scan y los totales."""
        
        color = "blue" if status == 'OK' else "red"
        self.last_scan_label.config(text=f"ÚLTIMO: {store} ({code}) -> {status}", foreground=color)
        
        self.total_scans_label.config(text=f"Total Registros Únicos: {TOTAL_SCANS}")
        
        for name, count in COUNTS.items():
            if name in self.count_labels:
                self.count_labels[name].config(text=f"{name}: {count}")

    def update_initial_interface(self):
        """Actualiza la interfaz con los datos cargados al inicio."""
        self.total_scans_label.config(text=f"Total Registros Únicos: {TOTAL_SCANS}")
        for name, count in COUNTS.items():
            if name in self.count_labels:
                self.count_labels[name].config(text=f"{name}: {count}")


if __name__ == "__main__":
    # Inicializar la aplicación y cargar la interfaz
    app = App()
    worker_thread = None
    app.mainloop()