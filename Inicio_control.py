from pathlib import Path
from datetime import datetime, timedelta
import locale
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import configparser
import sys
import threading

# --- PATCH: Silenciar errores de compatibilidad de Pandas/Dateutil en consola ---
class StderrFilter:
    def __init__(self, original_stderr):
        self.original_stderr = original_stderr
        self.buffer = ""

    def write(self, s):
        self.buffer += s
        if "\n" in self.buffer or len(self.buffer) > 500:
            self.process_buffer()

    def flush(self):
        if self.buffer:
            self.process_buffer(force=True)
        try:
            self.original_stderr.flush()
        except:
            pass
            
    def process_buffer(self, force=False):
        error_signatures = [
            "pandas._libs.tslibs", 
            "total_seconds", 
            "_localize_tso",
            "AttributeError: 'NoneType' object"
        ]
        
        if any(sig in self.buffer for sig in error_signatures):
            self.buffer = ""
            return

        suspicious_starts = [
            "Exception ignored in:", 
            "Traceback (most recent call last):", 
            "AttributeError:"
        ]
        
        is_suspicious = any(start in self.buffer for start in suspicious_starts)
        
        if is_suspicious and len(self.buffer) < 300 and not force:
            return

        try:
            self.original_stderr.write(self.buffer)
        except:
            pass
        finally:
            self.buffer = ""

sys.stderr = StderrFilter(sys.stderr)

# Configuración de español
try:
    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Windows
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Linux
    except locale.Error:
        pass

def obtener_ruta_recurso(nombre_archivo: str) -> Path:
    base_dir = getattr(sys, "_MEIPASS", None)
    if base_dir:
        return Path(base_dir) / nombre_archivo
    return Path(__file__).resolve().parent / nombre_archivo

def aplicar_icono_ventana(ventana: tk.Tk) -> None:
    try:
        icon_path = obtener_ruta_recurso("icon.ico")
        if icon_path.exists():
            if isinstance(ventana, tk.Tk):
                ventana.iconbitmap(str(icon_path))
            else:
                ventana.winfo_toplevel().iconbitmap(str(icon_path))
    except Exception:
        pass


class ConfiguradorRutas:
    """Manejador de configuración de rutas"""

    def __init__(self):
        self.config_file = Path("config_pagos.ini")
        self.config = configparser.ConfigParser()
        
    def cargar_config(self):
        if self.config_file.exists():
            self.config.read(self.config_file, encoding='utf-8')
            return True
        return False
    
    def guardar_config(self, origen, proyecciones, final):
        self.config['RUTAS'] = {
            'archivo_origen': origen,
            'carpeta_proyecciones': proyecciones,
            'archivo_final': final
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
        return True

    def obtener_rutas(self):
        if 'RUTAS' not in self.config:
            return None
        return {
            'origen': Path(self.config['RUTAS']['archivo_origen']),
            'proyecciones': Path(self.config['RUTAS']['carpeta_proyecciones']),
            'final': Path(self.config['RUTAS']['archivo_final'])
        }


class MainApp(tk.Tk):
    """Aplicación principal unificada"""
    
    def __init__(self):
        super().__init__()
        
        # Modern color palette
        self.COLOR_PRIMARIO = "#2C3E50"      # Blue-Gray Dark
        self.COLOR_SECUNDARIO = "#3498DB"    # Blue
        self.COLOR_SEMANAL = "#3498DB"
        self.COLOR_MENSUAL = "#9B59B6"
        self.COLOR_ACENTO = "#27AE60"        # Green
        self.COLOR_FONDO = "#F8F9FA"         # Light Gray
        self.COLOR_SIDEBAR = "#1A252F"       # Darker Blue-Gray
        self.COLOR_CARTA = "#FFFFFF"
        self.COLOR_TEXTO = "#2C3E50"
        self.COLOR_TEXTO_CLARO = "#7F8C8D"
        self.COLOR_BORDE = "#E0E0E0"
        
        self.title("Control de Pagos GCO - Panel Unificado")
        self.geometry("1200x800")
        self.minsize(700, 750)
        aplicar_icono_ventana(self)
        
        # Maximize window
        try:
            self.state('zoomed')
        except:
            pass
            
        self.configurador = ConfiguradorRutas()
        self.rutas = None
        
        self.current_view = None
        self.setup_ui()
        
        # Load configuration or show settings
        if not self.configurador.cargar_config():
            self.show_view("CONFIG")
            messagebox.showinfo("Configuración", "Por favor, configure las rutas de los archivos.")
        else:
            self.rutas = self.configurador.obtener_rutas()
            self.show_view("HOME")

    def setup_ui(self):
        # Main Content Area
        self.content_container = tk.Frame(self, bg=self.COLOR_FONDO)
        self.content_container.pack(fill=tk.BOTH, expand=True)
        
    def show_view(self, view_id):
        self.active_view = view_id
        
        # Clear content area
        for widget in self.content_container.winfo_children():
            widget.destroy()
            
        # Load view
        if view_id == "HOME":
            self.current_view = HomeView(self.content_container, self)
        elif view_id == "SEMANAL":
            self.current_view = WeeklyView(self.content_container, self)
        elif view_id == "MENSUAL":
            self.current_view = MonthlyView(self.content_container, self)
        elif view_id == "CONFIG":
            self.current_view = ConfigView(self.content_container, self)
        elif view_id == "PROGRESS":
            self.current_view = ProgressView(self.content_container, self)
            
        self.current_view.pack(fill=tk.BOTH, expand=True)


class BaseView(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=controller.COLOR_FONDO)
        self.controller = controller
        
    def create_header(self, title, subtitle, icon="📊"):
        header = tk.Frame(self, bg=self.controller.COLOR_CARTA, height=120, bd=0)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        # Bottom border for header
        border = tk.Frame(header, bg=self.controller.COLOR_BORDE, height=1)
        border.pack(side=tk.BOTTOM, fill=tk.X)
        
        content = tk.Frame(header, bg=self.controller.COLOR_CARTA)
        content.place(relx=0.05, rely=0.5, anchor="w")
        
        # Icon
        tk.Label(
            content,
            text=icon,
            font=("Segoe UI", 32),
            bg=self.controller.COLOR_CARTA
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        # Text
        text_frame = tk.Frame(content, bg=self.controller.COLOR_CARTA)
        text_frame.pack(side=tk.LEFT)
        
        tk.Label(
            text_frame,
            text=title,
            font=("Segoe UI", 24, "bold"),
            bg=self.controller.COLOR_CARTA,
            fg=self.controller.COLOR_PRIMARIO
        ).pack(anchor="w")
        
        tk.Label(
            text_frame,
            text=subtitle,
            font=("Segoe UI", 11),
            bg=self.controller.COLOR_CARTA,
            fg=self.controller.COLOR_TEXTO_CLARO
        ).pack(anchor="w")

        # Back button (only if not HOME)
        if self.controller.active_view != "HOME":
            btn_back = tk.Button(
                header,
                text="⬅ Volver al Inicio",
                font=("Segoe UI", 10, "bold"),
                bg=self.controller.COLOR_CARTA,
                fg=self.controller.COLOR_SECUNDARIO,
                activebackground=self.controller.COLOR_FONDO,
                relief=tk.FLAT,
                cursor="hand2",
                padx=20,
                pady=10,
                command=lambda: self.controller.show_view("HOME")
            )
            btn_back.place(relx=0.95, rely=0.5, anchor="e")
            
            def on_enter(e): btn_back.configure(fg=self.controller.COLOR_PRIMARIO)
            def on_leave(e): btn_back.configure(fg=self.controller.COLOR_SECUNDARIO)
            btn_back.bind("<Enter>", on_enter)
            btn_back.bind("<Leave>", on_leave)


class HomeView(BaseView):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_header("Panel de Control", "Seleccione el tipo de proyección que desea realizar", "🏠")
        
        content = tk.Frame(self, bg=self.controller.COLOR_FONDO)
        content.pack(fill=tk.BOTH, expand=True, padx=50, pady=50)
        
        cards_frame = tk.Frame(content, bg=self.controller.COLOR_FONDO)
        cards_frame.place(relx=0.5, rely=0.4, anchor="center")
        
        # Weekly Card
        self.create_card(
            cards_frame, 
            "Proyección Semanal", 
            "Generar reporte de pagos proyectados para la próxima semana.",
            "📅", 
            self.controller.COLOR_SEMANAL,
            lambda: self.controller.show_view("SEMANAL")
        ).grid(row=0, column=0, padx=20)
        
        # Monthly Card
        self.create_card(
            cards_frame, 
            "Proyección Mensual", 
            "Generar reporte consolidado de pagos para el mes completo.",
            "📊", 
            self.controller.COLOR_MENSUAL,
            lambda: self.controller.show_view("MENSUAL")
        ).grid(row=0, column=1, padx=20)

        # Config Button (Smaller, below cards)
        btn_config = tk.Button(
            content,
            text="⚙️ Configuración de Rutas",
            font=("Segoe UI", 10),
            bg=self.controller.COLOR_FONDO,
            fg=self.controller.COLOR_TEXTO_CLARO,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.controller.show_view("CONFIG")
        )
        btn_config.place(relx=0.5, rely=0.85, anchor="center")
        
        def on_enter(e): btn_config.configure(fg=self.controller.COLOR_PRIMARIO)
        def on_leave(e): btn_config.configure(fg=self.controller.COLOR_TEXTO_CLARO)
        btn_config.bind("<Enter>", on_enter)
        btn_config.bind("<Leave>", on_leave)

    def create_card(self, parent, title, desc, icon, color, command):
        card = tk.Frame(parent, bg=self.controller.COLOR_CARTA, width=350, height=250, cursor="hand2")
        card.pack_propagate(False)
        
        # Hover effect
        def on_enter(e): card.configure(bg="#F1F8FF")
        def on_leave(e): card.configure(bg=self.controller.COLOR_CARTA)
        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)
        card.bind("<Button-1>", lambda e: command())
        
        # Border
        card.configure(highlightbackground=self.controller.COLOR_BORDE, highlightthickness=1)
        
        # Top accent bar
        accent = tk.Frame(card, bg=color, height=5)
        accent.pack(fill=tk.X)
        
        tk.Label(card, text=icon, font=("Segoe UI", 48), bg=self.controller.COLOR_CARTA).pack(pady=(30, 10))
        tk.Label(card, text=title, font=("Segoe UI", 16, "bold"), bg=self.controller.COLOR_CARTA, fg=self.controller.COLOR_PRIMARIO).pack()
        
        desc_lbl = tk.Label(
            card, 
            text=desc, 
            font=("Segoe UI", 10), 
            bg=self.controller.COLOR_CARTA, 
            fg=self.controller.COLOR_TEXTO_CLARO,
            wraplength=280,
            justify=tk.CENTER
        )
        desc_lbl.pack(pady=15)
        
        # Bind events to children
        for child in card.winfo_children():
            child.bind("<Button-1>", lambda e: command())
            
        return card


class ConfigView(BaseView):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_header("Configuración", "Gestione las rutas de los archivos de origen y destino", "⚙️")
        
        content = tk.Frame(self, bg=self.controller.COLOR_FONDO)
        content.pack(fill=tk.BOTH, expand=True, padx=50, pady=50)
        
        # Wider form container
        form = tk.Frame(content, bg=self.controller.COLOR_CARTA, padx=40, pady=40, highlightbackground=self.controller.COLOR_BORDE, highlightthickness=1)
        # Use relwidth for responsive wide layout
        form.place(relx=0.5, rely=0.4, anchor="center", relwidth=0.7, height=550) 
        
        # Ensure rutas is not None
        rutas = self.controller.rutas or {'origen': '', 'proyecciones': '', 'final': ''}
        
        self.ruta_origen = tk.StringVar(value=str(rutas.get('origen', "")))
        self.ruta_proyecciones = tk.StringVar(value=str(rutas.get('proyecciones', "")))
        self.ruta_final = tk.StringVar(value=str(rutas.get('final', "")))
        
        self.create_field(form, "Archivo CONTROL DE PAGOS (.xlsm)", self.ruta_origen, True)
        self.create_field(form, "Carpeta de PROYECCIONES", self.ruta_proyecciones, False)
        self.create_field(form, "Archivo CONTROL PAGOS Final (.xlsx)", self.ruta_final, True)
        
        btn_save = tk.Button(
            form,
            text="Guardar Configuración",
            font=("Segoe UI", 11, "bold"),
            bg=self.controller.COLOR_ACENTO,
            fg="white",
            relief=tk.FLAT,
            padx=40,
            pady=12,
            cursor="hand2",
            activebackground="#219150",
            activeforeground="white",
            highlightthickness=0,
            bd=0
        )
        btn_save.pack(pady=(30, 0))
        btn_save.configure(command=self.save_config)

    def create_field(self, parent, label, var, is_file):
        frame = tk.Frame(parent, bg=self.controller.COLOR_CARTA)
        frame.pack(fill=tk.X, pady=12)
        
        tk.Label(frame, text=label, font=("Segoe UI", 9, "bold"), bg=self.controller.COLOR_CARTA, fg=self.controller.COLOR_TEXTO_CLARO).pack(anchor="w", padx=2)
        
        entry_frame = tk.Frame(frame, bg=self.controller.COLOR_CARTA)
        entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        entry = tk.Entry(
            entry_frame, 
            textvariable=var, 
            font=("Segoe UI", 10), 
            bd=1, 
            relief=tk.SOLID,
            highlightthickness=1,
            highlightbackground=self.controller.COLOR_BORDE,
            highlightcolor=self.controller.COLOR_SECUNDARIO
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 10))
        
        btn_browse = tk.Button(
            entry_frame, 
            text="Buscar...", 
            command=lambda: self.browse(var, is_file),
            bg="#E5E7E9",
            fg=self.controller.COLOR_PRIMARIO,
            font=("Segoe UI", 9),
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2",
            activebackground=self.controller.COLOR_BORDE
        )
        btn_browse.pack(side=tk.LEFT)

    def browse(self, var, is_file):
        if is_file:
            path = filedialog.askopenfilename()
        else:
            path = filedialog.askdirectory()
        if path:
            var.set(path)

    def save_config(self):
        if not all([self.ruta_origen.get(), self.ruta_proyecciones.get(), self.ruta_final.get()]):
            messagebox.showerror("Error", "Todas las rutas son obligatorias.")
            return
            
        self.controller.configurador.guardar_config(
            self.ruta_origen.get(),
            self.ruta_proyecciones.get(),
            self.ruta_final.get()
        )
        self.controller.rutas = self.controller.configurador.obtener_rutas()
        messagebox.showinfo("Éxito", "Configuración guardada correctamente.")
        self.controller.show_view("HOME")


class WeeklyView(BaseView):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_header("Proyección Semanal", "Seleccione la fecha de filtrado para el reporte semanal", "📅")
        
        import proceso_semanal
        self.view_impl = proceso_semanal.WeeklyFrame(self, self.controller)
        self.view_impl.pack(fill=tk.BOTH, expand=True)


class MonthlyView(BaseView):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_header("Proyección Mensual", "Seleccione el mes y año para el reporte mensual", "📊")
        
        import proceso_mensual
        self.view_impl = proceso_mensual.MonthlyFrame(self, self.controller)
        self.view_impl.pack(fill=tk.BOTH, expand=True)


class ProgressView(BaseView):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_header("Procesando...", "Espere mientras se genera el reporte", "⚙️")
        
        content = tk.Frame(self, bg=self.controller.COLOR_FONDO)
        content.pack(fill=tk.BOTH, expand=True, padx=50, pady=30)
        
        # Progress Card
        card = tk.Frame(content, bg=self.controller.COLOR_CARTA, padx=30, pady=30, highlightbackground=self.controller.COLOR_BORDE, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)
        
        self.status_label = tk.Label(
            card,
            text="Iniciando proceso...",
            font=("Segoe UI", 14, "bold"),
            bg=self.controller.COLOR_CARTA,
            fg=self.controller.COLOR_PRIMARIO,
            anchor="w"
        )
        self.status_label.pack(fill=tk.X, pady=(0, 20))
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(card, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        self.percent_label = tk.Label(
            card,
            text="0%",
            font=("Segoe UI", 11, "bold"),
            bg=self.controller.COLOR_CARTA,
            fg=self.controller.COLOR_SECUNDARIO
        )
        self.percent_label.pack()
        
        # Log Area
        log_frame = tk.Frame(card, bg="#F5F5F5", pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        tk.Label(log_frame, text="Registro de Actividad", font=("Segoe UI", 10, "bold"), bg="#F5F5F5").pack(anchor="w", padx=10)
        
        self.log_text = tk.Text(
            log_frame,
            font=("Consolas", 10),
            bg="#F5F5F5",
            fg=self.controller.COLOR_TEXTO,
            relief=tk.FLAT,
            state=tk.DISABLED,
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = tk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Action Buttons (Hidden during process)
        self.btn_frame = tk.Frame(card, bg=self.controller.COLOR_CARTA)
        self.btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.btn_back = tk.Button(
            self.btn_frame,
            text="Volver al Inicio",
            font=("Segoe UI", 11),
            bg=self.controller.COLOR_BORDE,
            relief=tk.FLAT,
            padx=25,
            pady=10,
            command=lambda: self.controller.show_view("HOME")
        )
        # Initially hidden

    def log(self, mensaje, tipo="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        iconos = {"INFO": "ℹ️", "OK": "✅", "WARN": "⚠️", "ERROR": "❌", "PROCESS": "🔄"}
        colores = {"INFO": "#3498DB", "OK": "#27AE60", "WARN": "#F39C12", "ERROR": "#E74C3C", "PROCESS": "#9B59B6"}
        
        icono = iconos.get(tipo, "ℹ️")
        color = colores.get(tipo, "#2C3E50")
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_text.insert(tk.END, f"{icono} ", f"icon_{tipo}")
        self.log_text.insert(tk.END, f"{mensaje}\n", "message")
        
        self.log_text.tag_config("timestamp", foreground="#95A5A6")
        self.log_text.tag_config(f"icon_{tipo}", foreground=color)
        self.log_text.tag_config("message", foreground=self.controller.COLOR_TEXTO)
        
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        if tipo == "ERROR":
            self.status_label.config(text=f"❌ Error: {mensaje}", fg="#E74C3C")
        elif tipo == "OK":
            self.status_label.config(text=f"✅ {mensaje}", fg="#27AE60")
        else:
            self.status_label.config(text=mensaje, fg=self.controller.COLOR_PRIMARIO)
        
        self.update_idletasks()

    def set_progress(self, percent, status_text=None):
        self.progress_var.set(percent)
        self.percent_label.config(text=f"{int(percent)}%")
        if status_text:
            self.status_label.config(text=status_text)
        if percent >= 100:
            self.btn_back.pack(side=tk.RIGHT)
        self.update_idletasks()


def main():
    app = MainApp()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("CRASH_LOG.txt", "a", encoding='utf-8') as f:
            f.write(f"\n{'='*50}\nFECHA: {timestamp}\nERROR:\n{error_msg}\n{'='*50}\n")
        messagebox.showerror("Error Fatal", f"Ocurrió un error crítico. Consulte CRASH_LOG.txt")
