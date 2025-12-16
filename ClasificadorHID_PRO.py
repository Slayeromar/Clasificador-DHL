import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import threading
import queue
import os
import subprocess

OUTPUT_FILE_NAME = "Clasificacion.xlsx"
OUTPUT_FOLDER_PATH = r"C:\Users\Omar Zambrano\Desktop\Final DHL"
FILE_PATH = os.path.join(OUTPUT_FOLDER_PATH, OUTPUT_FILE_NAME)

SCAN_QUEUE = queue.Queue()
PROCESS_RUNNING = False
STORES = ['DHL', '99MINUTOS', 'FEDEX', 'TERRESTRE', 'BLINK']

DATA_CACHE = {}
COUNTS = {store: 0 for store in STORES}
TOTAL_SCANS = 0

COLORS = {
    'bg_primary': '#f8f9fa',
    'bg_secondary': '#ffcc00',
    'bg_card': '#ffffff',
    'bg_card_inactive': '#e9ecef',
    'accent': '#d40511',
    'accent_hover': '#b8030f',
    'success': '#28a745',
    'success_light': '#d4edda',
    'warning': '#ffc107',
    'error': '#d40511',
    'inactive': '#6c757d',
    'inactive_light': '#adb5bd',
    'text_primary': '#212529',
    'text_secondary': '#6c757d',
    'text_light': '#ffffff',
    'border_light': '#dee2e6',
    'DHL': '#ffcc00',
    '99MINUTOS': '#0066cc',
    'FEDEX': '#ff6600',
    'TERRESTRE': '#4caf50',
    'BLINK': '#8b3f99'
}

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
        messagebox.showerror("Error de Permiso", f"No se puede guardar el archivo '{OUTPUT_FILE_NAME}'.\n\nCIERRE EL EXCEL Y VUELVA A INTENTARLO.")
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
    app.after(0, app.set_inactive_mode)

class ModernButton(tk.Canvas):
    def __init__(self, parent, text, command, bg_color, hover_color, app_ref):
        self.app_ref = app_ref
        self.text = text
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.enabled = True
        self.animation_phase = 0

        super().__init__(parent, bg=COLORS['bg_primary'],
                         highlightthickness=0, cursor='hand2')

        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        self.bind('<Button-1>', self.on_click)

    def draw_button(self):
        self.delete('all')
        width = self.winfo_width() if self.winfo_width() > 1 else 160
        height = self.winfo_height() if self.winfo_height() > 1 else 50

        self.config(width=width, height=height)

        self.rect = self.create_rectangle(3, 3, width-3, height-3, fill=self.bg_color,
                                          outline='', width=0, tags='button')

        font_size = max(8, int(height * 0.3))
        self.text_id = self.create_text(width//2, height//2, text=self.text,
                                        fill='white',
                                        font=('Segoe UI', font_size, 'bold'), tags='button')

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
        self.draw_button()
        if self.enabled:
            self.itemconfig(self.rect, fill=self.bg_color)
            self.itemconfig(self.text_id, fill='white')
        else:
            self.itemconfig(self.rect, fill=COLORS['inactive_light'])
            self.itemconfig(self.text_id, fill=COLORS['text_secondary'])

class StoreCard(tk.Frame):
    def __init__(self, parent, store_name, color, app_ref):
        super().__init__(parent, bg=COLORS['bg_card'], relief=tk.FLAT)
        self.store_name = store_name
        self.color = color
        self.is_active = False
        self.app_ref = app_ref

        self.config(highlightbackground=color, highlightthickness=4)

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.name_label = tk.Label(self, text=store_name, bg=COLORS['bg_card'],
                                   fg=color, font=('Segoe UI', 16, 'bold'))
        self.name_label.grid(row=1, column=0, pady=(15, 5))

        self.count_label = tk.Label(self, text="0", bg=COLORS['bg_card'],
                                   fg=COLORS['text_primary'],
                                   font=('Segoe UI', 48, 'bold'))
        self.count_label.grid(row=2, column=0, pady=5)

        self.units_label = tk.Label(self, text="paquetes", bg=COLORS['bg_card'],
                                    fg=COLORS['text_secondary'], font=('Segoe UI', 11))
        self.units_label.grid(row=3, column=0, pady=(5, 15))

        self.bind('<Configure>', self.on_configure)

    def on_configure(self, event):
        self.update_fonts()

    def update_fonts(self):
        height = self.winfo_height()
        width = self.winfo_width()

        if height > 1:
            name_size = max(10, int(height * 0.15))
            count_size = max(24, int(height * 0.45))
            units_size = max(8, int(height * 0.1))

            self.name_label.config(font=('Segoe UI', name_size, 'bold'))
            self.count_label.config(font=('Segoe UI', count_size, 'bold'))
            self.units_label.config(font=('Segoe UI', units_size))

    def update_count(self, count):
        self.count_label.config(text=str(count))
        if self.is_active:
            self.flash_animation()

    def flash_animation(self):
        self.config(highlightbackground='#28a745', bg='#d4edda')
        self.name_label.config(bg='#d4edda')
        self.count_label.config(bg='#d4edda')
        self.units_label.config(bg='#d4edda')
        self.after(300, self.restore_colors)

    def restore_colors(self):
        self.config(highlightbackground=self.color, bg=COLORS['bg_card'])
        self.name_label.config(bg=COLORS['bg_card'])
        self.count_label.config(bg=COLORS['bg_card'])
        self.units_label.config(bg=COLORS['bg_card'])

    def set_inactive_mode(self):
        self.is_active = False
        self.config(bg=COLORS['bg_card_inactive'], highlightbackground=COLORS['inactive_light'])
        self.name_label.config(bg=COLORS['bg_card_inactive'], fg=COLORS['inactive'])
        self.count_label.config(bg=COLORS['bg_card_inactive'], fg=COLORS['inactive'])
        self.units_label.config(bg=COLORS['bg_card_inactive'])

    def set_active_mode(self):
        self.is_active = True
        self.config(bg=COLORS['bg_card'], highlightbackground=self.color)
        self.name_label.config(bg=COLORS['bg_card'], fg=self.color)
        self.count_label.config(bg=COLORS['bg_card'], fg=COLORS['text_primary'])
        self.units_label.config(bg=COLORS['bg_card'])

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Clasificador DHL Shepora - Sistema Profesional")
        self.configure(bg=COLORS['bg_primary'])

        try:
            self.state('zoomed')
        except:
            self.attributes('-fullscreen', True)

        self.fullscreen = False
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", self.exit_fullscreen)
        self.bind('<Configure>', self.on_window_configure)

        self.input_buffer = ""
        self.store_cards = {}
        self.pulse_animation_id = None
        self.last_width = 0
        self.last_height = 0

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind('<FocusIn>', self.handle_focus_in)
        self.bind('<FocusOut>', self.handle_focus_out)
        self.focus_force()
        self.set_inactive_mode()
        self.update_initial_interface()
        self.after(100, self.update_all_fonts)

    def on_window_configure(self, event):
        width = self.winfo_width()
        height = self.winfo_height()

        if width != self.last_width or height != self.last_height:
            self.last_width = width
            self.last_height = height
            self.after(100, self.update_all_fonts)

    def update_all_fonts(self):
        width = self.winfo_width()
        height = self.winfo_height()

        if width < 100 or height < 100:
            return

        banner_height = max(80, int(height * 0.1))
        banner_font_size = max(18, int(banner_height * 0.35))
        banner_subtitle_size = max(10, int(banner_height * 0.2))

        header_height = max(60, int(height * 0.08))
        header_title_size = max(14, int(header_height * 0.4))
        header_subtitle_size = max(8, int(header_height * 0.25))

        total_font_size = max(32, int(height * 0.08))
        total_label_size = max(10, int(height * 0.02))

        try:
            self.banner_icon.config(font=('Segoe UI', max(24, int(banner_height * 0.5))))
            self.banner_status.config(font=('Segoe UI', banner_font_size, 'bold'))
            self.banner_subtitle.config(font=('Segoe UI', banner_subtitle_size))

            self.title_label.config(font=('Segoe UI', header_title_size, 'bold'))
            self.subtitle_label.config(font=('Segoe UI', header_subtitle_size))

            self.total_label_text.config(font=('Segoe UI', total_label_size, 'bold'))
            self.total_scans_label.config(font=('Segoe UI', total_font_size, 'bold'))

            self.last_scan_label.config(font=('Consolas', max(12, int(height * 0.035)), 'bold'))

            self.focus_status.config(font=('Segoe UI', max(7, int(height * 0.015)), 'bold'))
            self.focus_indicator.config(font=('Segoe UI', max(7, int(height * 0.015))))
            self.system_status.config(font=('Segoe UI', max(7, int(height * 0.015)), 'bold'))
            self.system_indicator.config(font=('Segoe UI', max(7, int(height * 0.015))))

            for card in self.store_cards.values():
                card.update_fonts()

            for button in [self.start_button, self.stop_button, self.save_button, self.folder_button]:
                button.draw_button()

        except:
            pass

    def toggle_fullscreen(self, event=None):
        self.fullscreen = not self.fullscreen
        self.attributes("-fullscreen", self.fullscreen)
        return "break"

    def exit_fullscreen(self, event=None):
        self.fullscreen = False
        self.attributes("-fullscreen", False)
        return "break"

    def create_widgets(self):
        self.status_banner = tk.Frame(self, bg=COLORS['inactive'])
        self.status_banner.pack(fill='x', side=tk.TOP)

        banner_content = tk.Frame(self.status_banner, bg=COLORS['inactive'])
        banner_content.pack(expand=True, padx=20, pady=15)

        self.banner_icon = tk.Label(banner_content, text="⏸", bg=COLORS['inactive'],
                                    fg='white', font=('Segoe UI', 40))
        self.banner_icon.pack(side=tk.LEFT, padx=20)

        banner_text_frame = tk.Frame(banner_content, bg=COLORS['inactive'])
        banner_text_frame.pack(side=tk.LEFT, padx=10, fill='both', expand=True)

        self.banner_status = tk.Label(banner_text_frame, text="SISTEMA INACTIVO",
                                      bg=COLORS['inactive'], fg='white',
                                      font=('Segoe UI', 28, 'bold'), justify=tk.CENTER)
        self.banner_status.pack(anchor='center')

        self.banner_subtitle = tk.Label(banner_text_frame,
                                        text="Presione INICIAR para activar el escaneo",
                                        bg=COLORS['inactive'], fg='white',
                                        font=('Segoe UI', 14), justify=tk.CENTER)
        self.banner_subtitle.pack(anchor='center', pady=(5, 0))

        header = tk.Frame(self, bg=COLORS['bg_secondary'])
        header.pack(fill='x', side=tk.TOP, padx=0, pady=0)

        header_left = tk.Frame(header, bg=COLORS['bg_secondary'])
        header_left.pack(side=tk.LEFT, padx=30, pady=15)

        self.title_label = tk.Label(header_left, text="CLASIFICADOR DHL SHEPORA",
                                    bg=COLORS['bg_secondary'], fg=COLORS['accent'],
                                    font=('Segoe UI', 22, 'bold'), justify=tk.CENTER)
        self.title_label.pack(anchor='center')

        self.subtitle_label = tk.Label(header_left, text="Sistema Profesional de Gestión de Paquetería",
                                       bg=COLORS['bg_secondary'], fg=COLORS['text_secondary'],
                                       font=('Segoe UI', 10), justify=tk.CENTER)
        self.subtitle_label.pack(anchor='center', pady=(3, 0))

        main_container = tk.Frame(self, bg=COLORS['bg_primary'])
        main_container.pack(fill='both', expand=True, padx=20, pady=15)
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        stats_frame = tk.Frame(main_container, bg=COLORS['bg_primary'])
        stats_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))

        total_card = tk.Frame(stats_frame, bg=COLORS['bg_card'],
                             highlightbackground=COLORS['accent'], highlightthickness=3)
        total_card.pack(fill='x')

        total_inner = tk.Frame(total_card, bg=COLORS['bg_card'])
        total_inner.pack(fill='x', padx=30, pady=20)

        total_left = tk.Frame(total_inner, bg=COLORS['bg_card'])
        total_left.pack(side=tk.LEFT)

        self.total_label_text = tk.Label(total_left, text="TOTAL REGISTRADOS",
                                         bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                                         font=('Segoe UI', 12, 'bold'))
        self.total_label_text.pack(anchor='w')

        tk.Label(total_left, text="Paquetes clasificados exitosamente",
                bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                font=('Segoe UI', 9)).pack(anchor='w')

        self.total_scans_label = tk.Label(total_inner, text="0",
                                         bg=COLORS['bg_card'], fg=COLORS['accent'],
                                         font=('Segoe UI', 48, 'bold'))
        self.total_scans_label.pack(side=tk.RIGHT)

        stores_container = tk.Frame(main_container, bg=COLORS['bg_primary'])
        stores_container.grid(row=1, column=0, sticky='nsew', pady=(0, 15))

        for i in range(2):
            stores_container.grid_rowconfigure(i, weight=1, uniform='rows')
        for i in range(3):
            stores_container.grid_columnconfigure(i, weight=1, uniform='cols')

        for i, store in enumerate(STORES):
            card = StoreCard(stores_container, store, COLORS[store], self)
            card.grid(row=i//3, column=i%3, padx=12, pady=12, sticky='nsew')
            self.store_cards[store] = card

        bottom_frame = tk.Frame(main_container, bg=COLORS['bg_primary'])
        bottom_frame.grid(row=2, column=0, sticky='ew')

        scan_display = tk.Frame(bottom_frame, bg=COLORS['bg_card'],
                               highlightbackground=COLORS['border_light'],
                               highlightthickness=2)
        scan_display.pack(fill='x', pady=(0, 12))

        scan_inner = tk.Frame(scan_display, bg=COLORS['bg_card'])
        scan_inner.pack(padx=30, pady=15)

        tk.Label(scan_inner, text="ÚLTIMO ESCANEO",
                bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                font=('Segoe UI', 10, 'bold'), justify=tk.CENTER).pack()

        self.last_scan_label = tk.Label(scan_inner, text="Esperando código...",
                                       bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                                       font=('Consolas', 16, 'bold'), justify=tk.CENTER)
        self.last_scan_label.pack(pady=(5, 0))

        control_frame = tk.Frame(bottom_frame, bg=COLORS['bg_card'],
                                highlightbackground=COLORS['border_light'],
                                highlightthickness=2)
        control_frame.pack(fill='x')

        control_inner = tk.Frame(control_frame, bg=COLORS['bg_card'])
        control_inner.pack(fill='x', padx=30, pady=12)

        status_indicators = tk.Frame(control_inner, bg=COLORS['bg_card'])
        status_indicators.pack(side=tk.LEFT)

        self.focus_status = tk.Label(status_indicators, text="VENTANA",
                                    bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                                    font=('Segoe UI', 9, 'bold'), justify=tk.CENTER)
        self.focus_status.grid(row=0, column=0, sticky='w', padx=(0, 20))

        self.focus_indicator = tk.Label(status_indicators, text="Inactiva",
                                       bg=COLORS['bg_card'], fg=COLORS['error'],
                                       font=('Segoe UI', 9), justify=tk.CENTER)
        self.focus_indicator.grid(row=1, column=0, sticky='w', padx=(0, 20))

        self.system_status = tk.Label(status_indicators, text="SISTEMA",
                                     bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                                     font=('Segoe UI', 9, 'bold'), justify=tk.CENTER)
        self.system_status.grid(row=0, column=1, sticky='w')

        self.system_indicator = tk.Label(status_indicators, text="Detenido",
                                        bg=COLORS['bg_card'], fg=COLORS['inactive'],
                                        font=('Segoe UI', 9), justify=tk.CENTER)
        self.system_indicator.grid(row=1, column=1, sticky='w')

        buttons_container = tk.Frame(control_inner, bg=COLORS['bg_card'])
        buttons_container.pack(side=tk.RIGHT)

        self.start_button = ModernButton(buttons_container, "INICIAR SISTEMA",
                                        self.start_process,
                                        COLORS['success'], '#1e7e34', self)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ModernButton(buttons_container, "DETENER",
                                       self.stop_process,
                                       COLORS['error'], '#bd2130', self)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        self.stop_button.set_state('disabled')

        self.save_button = ModernButton(buttons_container, "GUARDAR",
                                       self.manual_save,
                                       COLORS['accent'], COLORS['accent_hover'], self)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.folder_button = ModernButton(buttons_container, "ABRIR CARPETA",
                                         open_output_folder,
                                         COLORS['warning'], '#e0a800', self)
        self.folder_button.pack(side=tk.LEFT, padx=5)

    def set_inactive_mode(self):
        self.status_banner.config(bg=COLORS['inactive'])
        self.set_bg_recursive(self.status_banner, COLORS['inactive'])

        self.banner_icon.config(text="⏸", fg='white')
        self.banner_status.config(text="SISTEMA INACTIVO", fg='white')
        self.banner_subtitle.config(text="Presione INICIAR para activar el escaneo", fg='white')
        self.system_indicator.config(text="Detenido", fg=COLORS['inactive'])

        if self.pulse_animation_id:
            self.after_cancel(self.pulse_animation_id)
            self.pulse_animation_id = None

        for card in self.store_cards.values():
            card.set_inactive_mode()

    def set_active_mode(self):
        self.status_banner.config(bg=COLORS['success'])
        self.set_bg_recursive(self.status_banner, COLORS['success'])

        self.banner_icon.config(text="▶", fg='white')
        self.banner_status.config(text="SISTEMA ACTIVO - ESCANEO HABILITADO", fg='white')
        self.banner_subtitle.config(text="El sistema está capturando códigos de barras", fg='white')
        self.system_indicator.config(text="Activo y Escaneando", fg=COLORS['success'])

        for card in self.store_cards.values():
            card.set_active_mode()

        self.start_pulse_animation()

    def set_bg_recursive(self, widget, color):
        try:
            widget.config(bg=color)
        except:
            pass
        for child in widget.winfo_children():
            self.set_bg_recursive(child, color)

    def start_pulse_animation(self):
        if not PROCESS_RUNNING:
            return

        current_bg = self.status_banner.cget('bg')
        new_bg = '#2ecc71' if current_bg == COLORS['success'] else COLORS['success']

        self.status_banner.config(bg=new_bg)
        self.set_bg_recursive(self.status_banner, new_bg)

        self.pulse_animation_id = self.after(800, self.start_pulse_animation)

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
        self.focus_indicator.config(text="Activa - Capturando", fg=COLORS['success'])

        if not hasattr(self, '_key_bound') or not self._key_bound:
            self.bind('<Key>', self.handle_key_input)
            self._key_bound = True

    def handle_focus_out(self, event):
        self.focus_indicator.config(text="Inactiva", fg=COLORS['error'])

        if hasattr(self, '_key_bound') and self._key_bound:
            self.unbind('<Key>')
            self._key_bound = False

    def start_process(self):
        global PROCESS_RUNNING, worker_thread
        if PROCESS_RUNNING: return

        PROCESS_RUNNING = True
        self.input_buffer = ""
        self.set_active_mode()
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
        self.stop_button.set_state('disabled')
        self.save_button.set_state('disabled')

        self.status_banner.config(bg=COLORS['warning'])
        self.set_bg_recursive(self.status_banner, COLORS['warning'])
        self.banner_status.config(text="DETENIENDO SISTEMA...")
        self.banner_subtitle.config(text="Guardando datos, por favor espere")

        self.after(500, self.check_worker_end)

    def check_worker_end(self):
        global worker_thread
        if worker_thread and worker_thread.is_alive():
            self.after(500, self.check_worker_end)
        else:
            self.set_inactive_mode()
            self.start_button.set_state('normal')
            self.save_button.set_state('normal')

    def manual_save(self):
        if save_current_data():
            original_text = self.banner_subtitle.cget('text')
            self.banner_subtitle.config(text="Datos guardados exitosamente")
            self.after(2000, lambda: self.banner_subtitle.config(text=original_text))

    def on_closing(self):
        if PROCESS_RUNNING:
            if messagebox.askyesno("Detener Proceso",
                                  "El sistema está activo.\n¿Desea detener y salir?"):
                self.stop_process()
                self.after(1000, self.destroy)
            else:
                return
        else:
            self.destroy()

    def update_scan_interface(self, store, code, status):
        color = COLORS['success'] if status == 'OK' else COLORS['error']
        status_text = "OK" if status == 'OK' else "DUPLICADO"

        self.last_scan_label.config(
            text=f"{store} | {code} | {status_text}",
            fg=color
        )
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
