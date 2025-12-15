import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import threading
import queue
import os
import subprocess

OUTPUT_FILE_NAME = "REGISTRO_HID.xlsx"
OUTPUT_FOLDER_PATH = r"C:\Users\Omar Zambrano\Desktop\DHL Clasify"
FILE_PATH = os.path.join(OUTPUT_FOLDER_PATH, OUTPUT_FILE_NAME)

SCAN_QUEUE = queue.Queue()
PROCESS_RUNNING = False
STORES = ['DHL', '99MINUTOS', 'FEDEX', 'TERRESTRE', 'BLINK']

DATA_CACHE = {}
COUNTS = {store: 0 for store in STORES}
TOTAL_SCANS = 0

COLORS = {
    'bg_primary': '#1a1a2e',
    'bg_secondary': '#16213e',
    'bg_card': '#0f3460',
    'accent': '#00d4ff',
    'accent_hover': '#00b8e6',
    'success': '#00ff88',
    'warning': '#ffa500',
    'error': '#ff4757',
    'text_primary': '#ffffff',
    'text_secondary': '#b8c1cc',
    'DHL': '#ffcc00',
    '99MINUTOS': '#00d4ff',
    'FEDEX': '#ff6b35',
    'TERRESTRE': '#4ecdc4',
    'BLINK': '#a29bfe'
}

def load_initial_data():
    global DATA_CACHE, COUNTS, TOTAL_SCANS
    DATA_CACHE = {}
    COUNTS = {store: 0 for store in STORES}
    TOTAL_SCANS = 0

    try:
        if not os.path.exists(FILE_PATH):
            print(f"Archivo {OUTPUT_FILE_NAME} no encontrado. Creando nueva caché.")
            return

        xls = pd.ExcelFile(FILE_PATH)
        sheets = xls.sheet_names

        for sheet in sheets:
            df = xls.parse(sheet)
            DATA_CACHE[sheet] = df

            ok_counts = df[df['Status'] == 'OK'].shape[0]
            if sheet in COUNTS:
                COUNTS[sheet] = ok_counts
            TOTAL_SCANS += ok_counts

        print(f"Datos cargados exitosamente desde {OUTPUT_FILE_NAME}.")
    except Exception as e:
        messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo Excel.\nError: {e}")

def save_current_data():
    try:
        os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)

        writer = pd.ExcelWriter(FILE_PATH, engine='xlsxwriter')
        all_sheets = set(STORES).union(DATA_CACHE.keys())

        for sheet in sorted(all_sheets):
            df = DATA_CACHE.get(sheet, pd.DataFrame(columns=['Timestamp', 'Code', 'Status']))
            df.to_excel(writer, sheet_name=sheet, index=False)

        writer.close()
        print(f"Datos guardados exitosamente en {OUTPUT_FILE_NAME}.")
        return True
    except PermissionError:
        messagebox.showerror("Error de Permiso", f"No se puede guardar el archivo. Asegúrese de que no esté abierto.")
        return False
    except Exception as e:
        messagebox.showerror("Error de Escritura", f"Error al guardar: {e}")
        return False

def open_output_folder():
    try:
        folder = OUTPUT_FOLDER_PATH
        if os.path.exists(folder):
            if os.name == 'nt':
                subprocess.Popen(['explorer', folder])
            elif os.uname()[0] == 'Darwin':
                subprocess.Popen(['open', folder])
            else:
                subprocess.Popen(['xdg-open', folder])
        else:
            messagebox.showinfo("Carpeta no encontrada", f"La carpeta {folder} aún no existe.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la carpeta. Error: {e}")

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
            if len(parts) != 2:
                continue

            store_name = parts[0].strip().upper()
            code = parts[1].strip()
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
    app.after(0, app.update_status, "DETENIDO", "Sistema detenido", COLORS['text_secondary'])

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
        if self.enabled:
            self.itemconfig(self.rect, fill=self.hover_color)

    def on_leave(self, e):
        if self.enabled:
            self.itemconfig(self.rect, fill=self.bg_color)

    def on_click(self, e):
        if self.enabled and self.command:
            self.command()

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

        self.config(highlightbackground=color, highlightthickness=2)

        name_label = tk.Label(self, text=store_name, bg=COLORS['bg_card'],
                             fg=color, font=('Segoe UI', 12, 'bold'))
        name_label.pack(pady=(10, 5))

        self.count_label = tk.Label(self, text="0", bg=COLORS['bg_card'],
                                   fg=COLORS['text_primary'],
                                   font=('Segoe UI', 28, 'bold'))
        self.count_label.pack(pady=(0, 5))

        units_label = tk.Label(self, text="paquetes", bg=COLORS['bg_card'],
                              fg=COLORS['text_secondary'], font=('Segoe UI', 9))
        units_label.pack(pady=(0, 10))

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
        self.title("Clasificador HID PRO V6 - Sistema de Gestión de Paquetería")

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 1100
        window_height = 700
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.configure(bg=COLORS['bg_primary'])
        self.input_buffer = ""
        self.store_cards = {}

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind('<FocusIn>', self.handle_focus_in)
        self.bind('<FocusOut>', self.handle_focus_out)
        self.focus_force()
        self.update_initial_interface()

    def create_widgets(self):
        header = tk.Frame(self, bg=COLORS['bg_secondary'], height=80)
        header.pack(fill='x', padx=0, pady=0)
        header.pack_propagate(False)

        title_label = tk.Label(header, text="CLASIFICADOR HID PRO",
                              bg=COLORS['bg_secondary'], fg=COLORS['accent'],
                              font=('Segoe UI', 24, 'bold'))
        title_label.pack(side=tk.LEFT, padx=30, pady=20)

        subtitle_label = tk.Label(header, text="Sistema de Gestión de Paquetería",
                                 bg=COLORS['bg_secondary'], fg=COLORS['text_secondary'],
                                 font=('Segoe UI', 11))
        subtitle_label.pack(side=tk.LEFT, padx=0, pady=20)

        main_container = tk.Frame(self, bg=COLORS['bg_primary'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)

        stats_frame = tk.Frame(main_container, bg=COLORS['bg_primary'])
        stats_frame.pack(fill='x', pady=(0, 20))

        total_card = tk.Frame(stats_frame, bg=COLORS['bg_card'],
                             highlightbackground=COLORS['accent'], highlightthickness=3)
        total_card.pack(fill='x', pady=10)

        total_inner = tk.Frame(total_card, bg=COLORS['bg_card'])
        total_inner.pack(fill='x', padx=20, pady=15)

        total_label_text = tk.Label(total_inner, text="TOTAL DE REGISTROS ÚNICOS",
                                    bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                                    font=('Segoe UI', 12, 'bold'))
        total_label_text.pack(side=tk.LEFT)

        self.total_scans_label = tk.Label(total_inner, text="0",
                                         bg=COLORS['bg_card'], fg=COLORS['accent'],
                                         font=('Segoe UI', 32, 'bold'))
        self.total_scans_label.pack(side=tk.RIGHT)

        stores_container = tk.Frame(main_container, bg=COLORS['bg_primary'])
        stores_container.pack(fill='both', expand=True, pady=(0, 20))

        for i, store in enumerate(STORES):
            card = StoreCard(stores_container, store, COLORS[store])
            card.grid(row=i//3, column=i%3, padx=10, pady=10, sticky='nsew')
            self.store_cards[store] = card

        for i in range(2):
            stores_container.grid_rowconfigure(i, weight=1)
        for i in range(3):
            stores_container.grid_columnconfigure(i, weight=1)

        status_frame = tk.Frame(main_container, bg=COLORS['bg_card'])
        status_frame.pack(fill='x', pady=(0, 20))

        status_inner = tk.Frame(status_frame, bg=COLORS['bg_card'])
        status_inner.pack(fill='x', padx=20, pady=15)

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

        scan_frame = tk.Frame(main_container, bg=COLORS['bg_card'])
        scan_frame.pack(fill='x', pady=(0, 20))

        scan_inner = tk.Frame(scan_frame, bg=COLORS['bg_card'])
        scan_inner.pack(fill='x', padx=20, pady=15)

        scan_label = tk.Label(scan_inner, text="ÚLTIMO ESCANEO:",
                             bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                             font=('Segoe UI', 10))
        scan_label.pack(side=tk.LEFT)

        self.last_scan_label = tk.Label(scan_inner, text="N/A",
                                       bg=COLORS['bg_card'], fg=COLORS['text_primary'],
                                       font=('Segoe UI', 11, 'bold'))
        self.last_scan_label.pack(side=tk.LEFT, padx=(10, 0))

        controls_frame = tk.Frame(main_container, bg=COLORS['bg_primary'])
        controls_frame.pack(fill='x')

        buttons_container = tk.Frame(controls_frame, bg=COLORS['bg_primary'])
        buttons_container.pack()

        self.start_button = ModernButton(buttons_container, "INICIAR",
                                        self.start_process,
                                        COLORS['success'], '#00e699',
                                        width=160, height=50)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ModernButton(buttons_container, "DETENER",
                                       self.stop_process,
                                       COLORS['error'], '#ff6b7a',
                                       width=160, height=50)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        self.stop_button.set_state('disabled')

        self.save_button = ModernButton(buttons_container, "GUARDAR",
                                       self.manual_save,
                                       COLORS['accent'], COLORS['accent_hover'],
                                       width=160, height=50)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.folder_button = ModernButton(buttons_container, "ABRIR CARPETA",
                                         open_output_folder,
                                         COLORS['warning'], '#ffb733',
                                         width=160, height=50)
        self.folder_button.pack(side=tk.LEFT, padx=5)

    def handle_key_input(self, event):
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
            if messagebox.askyesno("Detener Proceso", "El proceso está activo. ¿Desea detenerlo y guardar antes de salir?"):
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
