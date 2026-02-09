from pathlib import Path
from datetime import datetime, timedelta
import locale
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import configparser
import sys

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
            ventana.iconbitmap(str(icon_path))
    except Exception:
        pass


class ConfiguradorRutas:
    """Manejador de configuración de rutas"""

    def __init__(self):
        self.config_file = Path("config_pagos.ini")
        self.config = configparser.ConfigParser()
        
    def cargar_o_crear_config(self):
        if self.config_file.exists():
            self.config.read(self.config_file, encoding='utf-8')
            return True
        else:
            return self.crear_configuracion_inicial()
    
    def crear_configuracion_inicial(self):
        root = tk.Tk()
        aplicar_icono_ventana(root)
        root.withdraw()
        
        messagebox.showinfo(
            "Primera Configuración",
            "Por favor, seleccione las rutas necesarias para el programa."
        )
        
        messagebox.showinfo("Paso 1", "Seleccione el archivo CONTROL DE PAGOS de comercio")
        archivo_origen = filedialog.askopenfilename(
            title="Seleccionar CONTROL DE PAGOS.xlsm",
            filetypes=[("Excel Macro", "*.xlsm"), ("Todos", "*.*")]
        )
        
        if not archivo_origen:
            messagebox.showerror("Error", "Debe seleccionar el archivo origen.")
            return False
        
        messagebox.showinfo("Paso 2", "Seleccione la carpeta donde se guardarán las PROYECCIONES")
        carpeta_proyecciones = filedialog.askdirectory(
            title="Seleccionar carpeta de PROYECCIONES"
        )
        
        if not carpeta_proyecciones:
            messagebox.showerror("Error", "Debe seleccionar la carpeta de proyecciones.")
            return False
        
        messagebox.showinfo("Paso 3", "Seleccione el archivo CONTROL PAGOS.xlsx - archivo final")
        archivo_final = filedialog.askopenfilename(
            title="Seleccionar CONTROL PAGOS.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialdir=carpeta_proyecciones
        )
        
        if not archivo_final:
            messagebox.showerror("Error", "Debe seleccionar el archivo final.")
            return False
        
        self.config['RUTAS'] = {
            'archivo_origen': archivo_origen,
            'carpeta_proyecciones': carpeta_proyecciones,
            'archivo_final': archivo_final
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
        
        messagebox.showinfo("Configuración Guardada", 
                          f"La configuración se ha guardado en:\n{self.config_file.absolute()}")
        
        root.destroy()
        return True
    
    def obtener_rutas(self):
        return {
            'origen': Path(self.config['RUTAS']['archivo_origen']),
            'proyecciones': Path(self.config['RUTAS']['carpeta_proyecciones']),
            'final': Path(self.config['RUTAS']['archivo_final'])
        }


class VentanaSeleccionTipo:
    """Ventana inicial para seleccionar el tipo de proyección"""
    
    def __init__(self):
        self.tipo_seleccionado = None
        
        self.COLOR_SEMANAL = "#3498DB"
        self.COLOR_MENSUAL = "#3498DB"
        self.COLOR_FONDO = "#ECF0F1"
        self.COLOR_TEXTO = "#2C3E50"
        
    def crear_ventana(self):
        self.root = tk.Tk()
        aplicar_icono_ventana(self.root)
        self.root.title("Control de Pagos GCO - Selector")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        self.root.configure(bg=self.COLOR_FONDO)
        
        self.centrar_ventana()
        
        # Header
        header_frame = tk.Frame(self.root, bg="#2C3E50", height=100)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="CONTROL DE PAGOS GCO",
            font=("Segoe UI", 24, "bold"),
            bg="#2C3E50",
            fg="white"
        )
        title_label.pack(expand=True)
        
        # Contenido
        content_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        label_titulo = tk.Label(
            content_frame,
            text="Seleccione el tipo de proyección:",
            font=("Segoe UI", 14, "bold"),
            bg=self.COLOR_FONDO,
            fg=self.COLOR_TEXTO
        )
        label_titulo.pack(pady=(0, 30))
        
        # Frame para botones
        buttons_frame = tk.Frame(content_frame, bg=self.COLOR_FONDO)
        buttons_frame.pack(fill=tk.BOTH, expand=True)
        
        # Botón Semanal
        btn_semanal = tk.Button(
            buttons_frame,
            text="📅 PROYECCIÓN SEMANAL",
            font=("Segoe UI", 14, "bold"),
            bg=self.COLOR_SEMANAL,
            fg="white",
            activebackground="#2980B9",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.seleccionar_tipo("SEMANAL")
        )
        btn_semanal.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        btn_semanal.bind("<Enter>", lambda e: btn_semanal.configure(bg="#2980B9"))
        btn_semanal.bind("<Leave>", lambda e: btn_semanal.configure(bg=self.COLOR_SEMANAL))
        
        # Botón Mensual
        btn_mensual = tk.Button(
            buttons_frame,
            text="📊 PROYECCIÓN MENSUAL",
            font=("Segoe UI", 14, "bold"),
            bg=self.COLOR_MENSUAL,
            fg="white",
            activebackground="#7D3C98",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.seleccionar_tipo("MENSUAL")
        )
        btn_mensual.pack(fill=tk.BOTH, expand=True)
        
        btn_mensual.bind("<Enter>", lambda e: btn_mensual.configure(bg="#7D3C98"))
        btn_mensual.bind("<Leave>", lambda e: btn_mensual.configure(bg=self.COLOR_MENSUAL))
        
        # Footer con botón cancelar
        footer_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        footer_frame.pack(fill=tk.X, padx=40, pady=(0, 20))
        
        btn_cancelar = tk.Button(
            footer_frame,
            text="Cancelar",
            font=("Segoe UI", 10),
            bg="#95A5A6",
            fg="white",
            activebackground="#7F8C8D",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=self.cancelar
        )
        btn_cancelar.pack(side=tk.RIGHT)
        
        self.root.protocol("WM_DELETE_WINDOW", self.cancelar)
        self.root.mainloop()
    
    def centrar_ventana(self):
        self.root.update_idletasks()
        ancho = self.root.winfo_width()
        alto = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto // 2)
        self.root.geometry(f'{ancho}x{alto}+{x}+{y}')
    
    def seleccionar_tipo(self, tipo):
        self.tipo_seleccionado = tipo
        self.root.destroy()
    
    def cancelar(self):
        self.tipo_seleccionado = None
        self.root.destroy()


class VentanaProgreso:
    """Ventana de progreso durante la ejecución"""
    
    def __init__(self):
        self.root = tk.Tk()
        aplicar_icono_ventana(self.root)
        self.root.title("Procesando...")
        self.root.geometry("600x300")
        self.root.resizable(False, False)
        self.root.configure(bg="#ECF0F1")
        
        self.centrar_ventana()
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg="#ECF0F1")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        titulo = tk.Label(
            main_frame,
            text="⚙️ Procesando Control de Pagos",
            font=("Segoe UI", 16, "bold"),
            bg="#ECF0F1",
            fg="#2C3E50"
        )
        titulo.pack(pady=(0, 20))
        
        # Barra de progreso
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=560
        )
        self.progress.pack(pady=(10, 0))
        self.progress.start(10)
        
        self.root.update()
    
    def centrar_ventana(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def log(self, mensaje, tipo="INFO"):
        """Agrega mensaje al log"""
        self.root.update()
    
    def cerrar(self):
        """Cierra la ventana"""
        self.progress.stop()
        self.root.destroy()


def main():
    """Función principal unificada"""
    # Cargar configuración
    configurador = ConfiguradorRutas()
    if not configurador.cargar_o_crear_config():
        return
    
    rutas = configurador.obtener_rutas()
    
    # Bucle principal
    while True:
        try:
            # 1. Seleccionar tipo de proyección
            selector = VentanaSeleccionTipo()
            selector.crear_ventana()
            
            if selector.tipo_seleccionado is None:
                break
            
            tipo_proyeccion = selector.tipo_seleccionado
            
            # 2. Importar módulo correspondiente
            if tipo_proyeccion == "SEMANAL":
                import proceso_semanal as proceso_semanal
                
                # Ejecutar interfaz semanal
                interfaz = proceso_semanal.InterfazSemanal()
                interfaz.crear_ventana()
                
                if not interfaz.ejecutar_proceso:
                    continue
                
                # Confirmar ejecución
                if not messagebox.askyesno(
                    "Confirmar Ejecución",
                    "Antes de continuar, asegúrese de:\n\n"
                    "   ✓ Haber actualizado el archivo 'CONTROL DE PAGOS.xlsm'\n"
                    "   ✓ Haber guardado todos los cambios\n"
                    "   ✓ Cerrar el archivo si está abierto\n\n"
                    "¿Desea continuar?"
                ):
                    continue
                
                # Ejecutar proceso
                ventana_prog = VentanaProgreso()
                try:
                    procesador = proceso_semanal.ProcesadorSemanal(
                        fecha_filtrado=interfaz.fecha_seleccionada,
                        ventana_progreso=ventana_prog,
                        rutas_config=rutas
                    )
                    resultado = procesador.ejecutar_proceso()
                    ventana_prog.cerrar()
                except Exception as e:
                    ventana_prog.cerrar()
                    messagebox.showerror("Error Fatal", f"Error inesperado:\n\n{str(e)}")
            
            elif tipo_proyeccion == "MENSUAL":
                import proceso_mensual as proceso_mensual
                
                # Ejecutar interfaz mensual
                interfaz = proceso_mensual.InterfazMensual()
                interfaz.crear_ventana()
                
                if not interfaz.ejecutar_proceso:
                    continue
                
                # Ejecutar proceso
                ventana_prog = VentanaProgreso()
                try:
                    procesador = proceso_mensual.ProcesadorMensual(
                        fecha_filtrado=interfaz.fecha_seleccionada,
                        ventana_progreso=ventana_prog,
                        rutas_config=rutas
                    )
                    resultado = procesador.ejecutar_proceso()
                    ventana_prog.cerrar()
                except Exception as e:
                    ventana_prog.cerrar()
                    messagebox.showerror("Error Fatal", f"Error inesperado:\n\n{str(e)}")
            
            # Preguntar si desea procesar otra proyección
            if not messagebox.askyesno(
                "Proceso Completado",
                "¿Desea realizar otra proyección?"
            ):
                break
        
        except Exception as e:
            messagebox.showerror("Error de Configuración", f"Error al iniciar:\n\n{str(e)}")
            if not messagebox.askyesno("Error", "¿Desea intentar nuevamente?"):
                break


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            with open("CRASH_LOG.txt", "a", encoding='utf-8') as f:
                f.write(f"\n{'='*50}\n")
                f.write(f"FECHA: {timestamp}\n")
                f.write(f"ERROR:\n{error_msg}\n")
                f.write(f"{'='*50}\n")
        except:
            pass
        
        try:
            root = tk.Tk()
            aplicar_icono_ventana(root)
            root.withdraw()
            messagebox.showerror("Error Fatal", f"Ocurrió un error crítico:\n\n{str(e)}\n\nConsulte CRASH_LOG.txt")
        except:
            print(f"Error fatal: {e}")
            input("Presione Enter para salir...")