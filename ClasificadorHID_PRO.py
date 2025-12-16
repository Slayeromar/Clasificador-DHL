import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import threading
import queue
import os
import subprocess

# --- CONFIGURACIÓN GLOBAL ---
OUTPUT_FILE_NAME = "Clasificacion.xlsx"
OUTPUT_FOLDER_PATH = r"C:\Users\Omar Zambrano\Desktop\Final DHL"
FILE_PATH = os.path.join(OUTPUT_FOLDER_PATH, OUTPUT_FILE_NAME)

SCAN_QUEUE = queue.Queue()
PROCESS_RUNNING = False
STORES = ['DHL', '99MINUTOS', 'FEDEX', 'TERRESTRE', 'BLINK']

DATA_CACHE = {}
COUNTS = {store: 0 for store in STORES}
TOTAL_SCANS = 0

# Paleta de colores - Tema DHL
COLORS = {
    'bg_primary': '#ffffff',
    'bg_secondary': '#ffcc00',  # Amarillo DHL
    'bg_card': '#f4f4f4',       # Gris muy claro para contraste
    'accent': '#d40511',        # Rojo DHL
    'accent_hover': '#b8030f',
    'success': '#28a745',
    'warning': '#ff9800',
    'error': '#d40511',
    'text_primary': '#000000',
    'text_secondary': '#5d5d5d',
    'DHL': '#ffcc00',
    '99MINUTOS': '#007bff',
    'FEDEX': '#ff6b35',
    'TERRESTRE': '#4caf50',
    'BLINK': '#9c27b0'
}

# ----------------- Lógica de Archivos y Datos -----------------

def load_initial_data():
    global DATA_CACHE, COUNTS, TOTAL_SCANS
    DATA_CACHE = {}
    COUNTS = {store: 0 for store in STORES}
    TOTAL_SCANS = 0

    try:
        if not os.path.exists(FILE_PATH):
            return

        xls = pd.ExcelFile(FILE_PATH)
        sheets = xls.sheet_names

        for sheet in sheets:
            df = xls.parse(sheet, dtype={'Code': str})
            DATA_CACHE[sheet] = df
            ok_counts = df.shape[0]
            if sheet in COUNTS:
                COUNTS[sheet] = ok_counts
            TOTAL_SCANS += ok_counts

    except Exception as e:
        messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo Excel.\nError: {e}")

def save_current_data():
    try:
        os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)
        writer = pd.ExcelWriter(FILE_PATH, engine='xlsxwriter')
        workbook = writer.book
        
        fmt_header_base = {'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12}
        fmt_text = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'num_format': '@'})
        fmt_center = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})

        all_sheets = set(STORES).union(DATA_CACHE.keys())

        for sheet_name in sorted(all_sheets):
            df = DATA_CACHE.get(sheet_name, pd.DataFrame(columns=['Timestamp', 'Code', 'Status']))
            df_final = df[df['Status'] == 'OK'].copy()
            df_final.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            header_color = COLORS.get(sheet_name, '#dddddd')
            fmt_header = workbook.add_format(fmt_header_base)
            fmt_header.set_bg_color(header_color)

            worksheet.set_column('A:A', 20, fmt_center) 
            worksheet.set_column('B:B', 30, fmt_text)   
            worksheet.set_column('C:C', 10, fmt_center) 

            for col_num, value in enumerate(df_final.columns.values):
                worksheet.write(0, col_num, value, fmt_header)

        writer.close()
        return True

    except PermissionError:
        messagebox.showerror("Error de Permiso", f"No se puede guardar el archivo '{OUTPUT_FILE_NAME}'.\n\n⚠️ CIERRE EL EXCEL Y VUELVA A INTENTARLO.")
        return False
    except Exception as e:
        messagebox.showerror("Error de Escritura", f"Error crítico al guardar: {e}")
        return False

def open_output_folder():
    try:
        folder = OUTPUT_FOLDER_PATH
        if os.path.exists(folder):
            os.startfile(folder) if os.name == 'nt' else subprocess.Popen(['xdg-open', folder])
        else:
            messagebox.showinfo("Carpeta no encontrada", f"La carpeta {folder} aún no existe.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la carpeta. Error: {e}")

# ----------------- Procesamiento (Worker) -----------------

def process_worker():
    global TOTAL_SCANS, COUNTS, DATA_CACHE

    load_initial_data()
    app.after(0, app.update_initial_interface)

    while PROCESS_RUNNING or not SCAN_QUEUE.empty():
        try:
            line = SCAN_QUEUE.get(timeout=0.1)
        except queue.Empty:
            continue

        try:
            parts = line.split(',')
            if len(parts) != 2: continue

            store_name = parts[0].strip().upper()
            raw_code = parts[1].strip()
            code = raw_code.replace("'", "-").replace('"', '') 
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if store_name not in DATA_CACHE:
                DATA_CACHE[store_name] = pd.DataFrame(columns=['Timestamp', 'Code', 'Status'])

            df = DATA_CACHE[store_name]

            if code in df['Code'].values:
                status = 'DUP'
            else:
                status = 'OK'
                TOTAL_SCANS += 1
                COUNTS[store_name] = COUNTS.get(store_name, 0) + 1

            new_row = pd.DataFrame([{'Timestamp': timestamp, 'Code': code, 'Status': status}])
            DATA_CACHE[store_name] = pd.concat([df, new_row], ignore_index=True)

            app.after(0, app.update_scan_interface, store_name, code, status)

        except Exception as e:
            print(f"Error al procesar línea: {e}")
        finally:
            SCAN_QUEUE.task_done()

    app.after(0, app.save_button.config, {'state': tk.NORMAL})
    save_current_data()
    app.after(0, app.update_status, "DETENIDO", "Sistema detenido - Excel Actualizado", COLORS['text_secondary'])

# ----------------- Componentes UI Modernos -----------------

class ModernButton(tk.Canvas):
    def __init__(self, parent, text, command, bg_color, hover_color, width=140, height=45):
        super().__init__(parent, width=width, height=height, bg=COLORS['bg_primary'], 
                         highlightthickness=0, cursor='hand2')
        self.text = text
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.enabled = True

        self.rect = self.create_rectangle(0, 0, width, height, fill=bg_color, 
                                          outline='', width=0, tags='button')
        self.text_id = self.create_text(width//2, height//2, text=text, 
                                        fill=COLORS['text_primary'], 
                                        font=('Segoe UI', 11, 'bold'), tags='button')
        
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        self.bind('<Button-1>', self.on_click)

    def on_enter(self, e):
        if self.enabled: self.itemconfig(self.rect, fill=self.hover_color)

    def on_leave(self, e):
        if self.enabled: self.itemconfig(self.rect, fill=self.bg_color)

    def on_click(self, e):
        if self.enabled and self.command: self.command()

    def set_state(self, state):
        self.enabled = (state == 'normal')
        if self.enabled:
            self.itemconfig(self.rect, fill=self.bg_color)
            self.itemconfig(self.text_id, fill=COLORS['text_primary'])
        else:
            self.itemconfig(self.rect, fill=COLORS['bg_secondary'])
            self.itemconfig(self.text_id, fill=COLORS['text_secondary'])

class StoreCard(tk.Frame):
    def __init__(self, parent, store_name, color):
        super().__init__(parent, bg=COLORS['bg_card'], relief=tk.FLAT)
        self.store_name = store_name
        self.color = color
        
        # Borde coloreado
        self.config(highlightbackground=color, highlightthickness=2)
        
        # Configuración grid interna para centrado perfecto
        self.grid_rowconfigure(0, weight=1) # Espacio arriba
        self.grid_rowconfigure(4, weight=1) # Espacio abajo
        self.grid_columnconfigure(0, weight=1)

        # Contenido
        name_label = tk.Label(self, text=store_name, bg=COLORS['bg_card'], 
                             fg=color, font=('Segoe UI', 14, 'bold'))
        name_label.grid(row=1, column=0, pady=(5, 0))
        
        self.count_label = tk.Label(self, text="0", bg=COLORS['bg_card'], 
                                   fg=COLORS['text_primary'], 
                                   font=('Segoe UI', 36, 'bold'))
        self.count_label.grid(row=2, column=0, pady=0)
        
        units_label = tk.Label(self, text="paquetes", bg=COLORS['bg_card'], 
                              fg=COLORS['text_secondary'], font=('Segoe UI', 10))
        units_label.grid(row=3, column=0, pady=(0, 5))

    def update_count(self, count):
        self.count_label.config(text=str(count))
        self.flash_animation()

    def flash_animation(self):
        original_bg = self.color
        self.config(highlightbackground=COLORS['success'])
        self.after(200, lambda: self.config(highlightbackground=original_bg))

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Clasificador DHL Shepora")
        self.configure(bg=COLORS['bg_primary'])
        
        # Iniciar Maximizado (Windows)
        try:
            self.state('zoomed')
        except:
            self.attributes('-fullscreen', True) # Fallback para otros OS
            
        self.fullscreen = False
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", self.exit_fullscreen)

        self.input_buffer = ""
        self.store_cards = {}
        
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind('<FocusIn>', self.handle_focus_in)
        self.bind('<FocusOut>', self.handle_focus_out)
        self.focus_force() 
        self.update_initial_interface() 

    def toggle_fullscreen(self, event=None):
        self.fullscreen = not self.fullscreen
        self.attributes("-fullscreen", self.fullscreen)
        return "break"

    def exit_fullscreen(self, event=None):
        self.fullscreen = False
        self.attributes("-fullscreen", False)
        return "break"

    def create_widgets(self):
        # 1. Cabecera (Fija en la parte superior)
        header = tk.Frame(self, bg=COLORS['bg_secondary'], height=80)
        header.pack(fill='x', side=tk.TOP)
        header.pack_propagate(False) # Mantiene la altura fija
        
        title_label = tk.Label(header, text="CLASIFICADOR DHL SHEPORA", 
                              bg=COLORS['bg_secondary'], fg=COLORS['accent'], 
                              font=('Arial', 24, 'bold'))
        title_label.pack(side=tk.LEFT, padx=30, pady=20)
        
        subtitle_label = tk.Label(header, text="Sistema de Gestión de Paquetería", 
                                 bg=COLORS['bg_secondary'], fg=COLORS['text_secondary'], 
                                 font=('Segoe UI', 11))
        subtitle_label.pack(side=tk.LEFT, padx=0, pady=20)

        # 2. Contenedor Principal (Flexible)
        main_container = tk.Frame(self, bg=COLORS['bg_primary'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)

        # --- Sección TOTAL ---
        stats_frame = tk.Frame(main_container, bg=COLORS['bg_primary'])
        stats_frame.pack(fill='x', pady=(0, 10))
        
        total_card = tk.Frame(stats_frame, bg=COLORS['bg_card'], 
                             highlightbackground=COLORS['accent'], highlightthickness=3)
        total_card.pack(fill='x', pady=5)
        
        total_inner = tk.Frame(total_card, bg=COLORS['bg_card'])
        total_inner.pack(fill='x', padx=20, pady=10)
        
        total_label_text = tk.Label(total_inner, text="TOTAL REGISTROS", 
                                   bg=COLORS['bg_card'], fg=COLORS['text_secondary'], 
                                   font=('Arial', 14, 'bold'))
        total_label_text.pack(side=tk.LEFT)
        
        self.total_scans_label = tk.Label(total_inner, text="0", 
                                         bg=COLORS['bg_card'], fg=COLORS['accent'], 
                                         font=('Arial', 32, 'bold'))
        self.total_scans_label.pack(side=tk.RIGHT)

        # --- Sección TIENDAS (Grid Flexible) ---
        stores_container = tk.Frame(main_container, bg=COLORS['bg_primary'])
        # IMPORTANTE: expand=True permite que ocupe todo el espacio vertical disponible
        stores_container.pack(fill='both', expand=True, pady=(0, 20)) 
        
        # Configuración de pesos para redimensionado automático
        for i in range(2): # 2 Filas
            stores_container.grid_rowconfigure(i, weight=1, uniform='rows')
        for i in range(3): # 3 Columnas
            stores_container.grid_columnconfigure(i, weight=1, uniform='cols')

        for i, store in enumerate(STORES):
            card = StoreCard(stores_container, store, COLORS[store])
            # Sticky nsew hace que la tarjeta se estire en todas direcciones
            card.grid(row=i//3, column=i%3, padx=10, pady=10, sticky='nsew')
            self.store_cards[store] = card
            
        # --- Sección ESTADO Y ESCANEO ---
        bottom_frame = tk.Frame(main_container, bg=COLORS['bg_primary'])
        bottom_frame.pack(fill='x', side=tk.BOTTOM)

        # Status Bar
        status_frame = tk.Frame(bottom_frame, bg=COLORS['bg_card'])
        status_frame.pack(fill='x', pady=(0, 10))
        
        status_inner = tk.Frame(status_frame, bg=COLORS['bg_card'])
        status_inner.pack(fill='x', padx=20, pady=10)
        
        self.status_indicator = tk.Canvas(status_inner, width=16, height=16, 
                                         bg=COLORS['bg_card'], highlightthickness=0)
        self.status_indicator.pack(side=tk.LEFT, padx=(0, 10))
        self.status_dot = self.status_indicator.create_oval(2, 2, 14, 14, 
                                                           fill=COLORS['text_secondary'], outline='')
        
        self.status_label = tk.Label(status_inner, text="LISTO", 
                                    bg=COLORS['bg_card'], fg=COLORS['text_secondary'], 
                                    font=('Segoe UI', 11, 'bold'))
        self.status_label.pack(side=tk.LEFT)
        
        self.status_desc = tk.Label(status_inner, text="Sistema en espera", 
                                   bg=COLORS['bg_card'], fg=COLORS['text_secondary'], 
                                   font=('Segoe UI', 10))
        self.status_desc.pack(side=tk.LEFT, padx=(10, 0))
        
        self.focus_indicator = tk.Label(status_inner, text="SIN FOCO", 
                                       bg=COLORS['bg_card'], fg=COLORS['error'], 
                                       font=('Segoe UI', 10, 'bold'))
        self.focus_indicator.pack(side=tk.RIGHT, padx=(0, 10))

        # Último Escaneo
        scan_frame = tk.Frame(bottom_frame, bg=COLORS['bg_card'])
        scan_frame.pack(fill='x', pady=(0, 15))
        
        scan_inner = tk.Frame(scan_frame, bg=COLORS['bg_card'])
        scan_inner.pack(fill='x', padx=20, pady=10)
        
        self.last_scan_label = tk.Label(scan_inner, text="ESPERANDO ESCANEO...", 
                                       bg=COLORS['bg_card'], fg=COLORS['text_secondary'], 
                                       font=('Arial', 14, 'bold'))
        self.last_scan_label.pack()

        # Botones
        buttons_container = tk.Frame(bottom_frame, bg=COLORS['bg_primary'])
        buttons_container.pack(pady=10)
        
        self.start_button = ModernButton(buttons_container, "INICIAR", self.start_process, 
                                        COLORS['success'], '#00e699')
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        self.stop_button = ModernButton(buttons_container, "DETENER", self.stop_process, 
                                       COLORS['error'], '#ff6b7a')
        self.stop_button.pack(side=tk.LEFT, padx=10)
        self.stop_button.set_state('disabled')
        
        self.save_button = ModernButton(buttons_container, "GUARDAR", self.manual_save, 
                                       COLORS['accent'], COLORS['accent_hover'])
        self.save_button.pack(side=tk.LEFT, padx=10)

        self.folder_button = ModernButton(buttons_container, "ABRIR CARPETA", open_output_folder, 
                                         COLORS['warning'], '#ffb733')
        self.folder_button.pack(side=tk.LEFT, padx=10)

    def handle_key_input(self, event):
        global PROCESS_RUNNING
        if not PROCESS_RUNNING: return

        char = event.char
        if char == '\r' or char == '\n':
            if self.input_buffer:
                SCAN_QUEUE.put(self.input_buffer)
                self.input_buffer = ""
        else:
            self.input_buffer += char

    def handle_focus_in(self, event):
        self.focus_indicator.config(text="CAPTURA ACTIVA", fg=COLORS['success'])
        self.bind('<Key>', self.handle_key_input)

    def handle_focus_out(self, event):
        self.focus_indicator.config(text="SIN FOCO", fg=COLORS['error'])
        self.unbind('<Key>')

    def start_process(self):
        global PROCESS_RUNNING, worker_thread
        if PROCESS_RUNNING: return

        PROCESS_RUNNING = True
        self.input_buffer = ""
        self.update_status("ACTIVO", "Esperando lecturas HID...", COLORS['success'])
        self.start_button.set_state('disabled')
        self.stop_button.set_state('normal')
        
        worker_thread = threading.Thread(target=process_worker)
        worker_thread.daemon = True
        worker_thread.start()
        self.focus_force() 

    def stop_process(self):
        global PROCESS_RUNNING
        if not PROCESS_RUNNING: return

        PROCESS_RUNNING = False
        self.update_status("DETENIENDO", "Guardando datos...", COLORS['warning'])
        self.stop_button.set_state('disabled')
        self.save_button.set_state('disabled')
        self.after(500, self.check_worker_end)

    def check_worker_end(self):
        global worker_thread
        if worker_thread and worker_thread.is_alive():
            self.after(500, self.check_worker_end)
        else:
            self.update_status("DETENIDO", "Sistema detenido", COLORS['text_secondary'])
            self.start_button.set_state('normal')
            self.save_button.set_state('normal')

    def manual_save(self):
        if save_current_data():
             self.update_status("GUARDADO", "Datos guardados exitosamente", COLORS['success'])
             self.after(2000, lambda: self.update_status("LISTO", "Sistema en espera", COLORS['text_secondary']))

    def on_closing(self):
        if PROCESS_RUNNING:
            if messagebox.askyesno("Detener Proceso", "¿Desea detener y salir?"):
                self.stop_process()
                self.after(500, self.destroy)
            else:
                return
        else:
            self.destroy()

    def update_status(self, status_text, desc_text, color):
        self.status_label.config(text=status_text, fg=color)
        self.status_desc.config(text=desc_text)
        self.status_indicator.itemconfig(self.status_dot, fill=color)

    def update_scan_interface(self, store, code, status):
        color = COLORS['success'] if status == 'OK' else COLORS['error']
        status_emoji = "OK" if status == 'OK' else "DUPLICADO"
        self.last_scan_label.config(text=f"{store} - {code} [{status_emoji}]", fg=color)
        self.total_scans_label.config(text=str(TOTAL_SCANS))
        
        for name, count in COUNTS.items():
            if name in self.store_cards:
                self.store_cards[name].update_count(count)

    def update_initial_interface(self):
        self.total_scans_label.config(text=str(TOTAL_SCANS))
        for name, count in COUNTS.items():
            if name in self.store_cards:
                self.store_cards[name].update_count(count)

if __name__ == "__main__":
    app = App()
    worker_thread = None
    app.mainloop()
