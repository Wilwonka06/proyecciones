"""
AUTOMATIZACIÓN COMPLETA - CONTROL DE PAGOS - VERSIÓN 1.0
Correcciones:
- Soluciona error de encoding con caracteres especiales
- Manejo robusto de errores de codificación
- Rutas configurables desde archivo config
"""

from enum import nonmember
import pandas as pd
import win32com.client
import pythoncom
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
from datetime import datetime, timedelta
import locale
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import configparser
import sys
import time
import traceback

# Configuración de español
try:
    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Windows
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Linux
    except locale.Error:
        pass

class ConfiguradorRutas:
    """Manejador de configuración de rutas"""
    
    def __init__(self):
        self.config_file = Path("config_pagos.ini")
        self.config = configparser.ConfigParser()
        
    def cargar_o_crear_config(self):
        """Carga configuración existente o crea una nueva"""
        if self.config_file.exists():
            self.config.read(self.config_file, encoding='utf-8')
            return True
        else:
            return self.crear_configuracion_inicial()
    
    def crear_configuracion_inicial(self):
        """Crea configuración inicial solicitando rutas al usuario"""
        root = tk.Tk()
        root.withdraw()
        
        messagebox.showinfo(
            "Primera Configuración",
            "Por favor, seleccione las rutas necesarias para el programa."
        )
        
        # Solicitar archivo origen
        messagebox.showinfo("Paso 1", "Seleccione el archivo CONTROL DE PAGOS.xlsm")
        archivo_origen = filedialog.askopenfilename(
            title="Seleccionar CONTROL DE PAGOS.xlsm",
            filetypes=[("Excel Macro", "*.xlsm"), ("Todos", "*.*")]
        )
        
        if not archivo_origen:
            messagebox.showerror("Error", "Debe seleccionar el archivo origen.")
            return False
        
        # Solicitar carpeta de proyecciones
        messagebox.showinfo("Paso 2", "Seleccione la carpeta donde se guardarán las PROYECCIONES")
        carpeta_proyecciones = filedialog.askdirectory(
            title="Seleccionar carpeta de PROYECCIONES"
        )
        
        if not carpeta_proyecciones:
            messagebox.showerror("Error", "Debe seleccionar la carpeta de proyecciones.")
            return False
        
        # Solicitar archivo final
        messagebox.showinfo("Paso 3", "Seleccione el archivo CONTROL PAGOS.xlsx (archivo final)")
        archivo_final = filedialog.askopenfilename(
            title="Seleccionar CONTROL PAGOS.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialdir=carpeta_proyecciones
        )
        
        if not archivo_final:
            messagebox.showerror("Error", "Debe seleccionar el archivo final.")
            return False
        
        # Guardar configuración
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
        """Obtiene las rutas configuradas"""
        return {
            'origen': Path(self.config['RUTAS']['archivo_origen']),
            'proyecciones': Path(self.config['RUTAS']['carpeta_proyecciones']),
            'final': Path(self.config['RUTAS']['archivo_final'])
        }

class InterfazModerna:
    """
    Interfaz gráfica moderna para seleccionar la fecha de filtrado (Proyección)
    """
    def __init__(self):
        self.fecha_seleccionada = None
        self.ejecutar_proceso = False
        
        # Colores del tema
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_ACENTO = "#27AE60"
        self.COLOR_FONDO = "#ECF0F1"
        self.COLOR_TEXTO = "#2C3E50"
        self.COLOR_ERROR = "#E74C3C"
        
    def crear_ventana(self):
        """Crea la ventana de interfaz moderna"""
        self.root = tk.Tk()
        self.root.title("Control de Pagos GCO")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        self.root.configure(bg=self.COLOR_FONDO)
        
        # Centrar ventana
        self.centrar_ventana()
        
        # Configurar estilo
        self.configurar_estilos()
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Header
        self.crear_header(main_frame)
        
        # Contenido principal
        self.crear_contenido(main_frame)
        
        # Footer con botones
        self.crear_footer(main_frame)
        
        self.root.mainloop()
    
    def configurar_estilos(self):
        """Configura los estilos personalizados"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure(
            "Modern.TLabelframe",
            background=self.COLOR_FONDO,
            bordercolor=self.COLOR_SECUNDARIO,
            borderwidth=2
        )
        style.configure(
            "Modern.TLabelframe.Label",
            background=self.COLOR_FONDO,
            foreground=self.COLOR_PRIMARIO,
            font=("Segoe UI", 11, "bold")
        )
    
    def crear_header(self, parent):
        """Crea el header con título"""
        header_frame = tk.Frame(parent, bg=self.COLOR_PRIMARIO, height=140)
        header_frame.pack(fill=tk.X, pady=0)
        header_frame.pack_propagate(False)
        
        content = tk.Frame(header_frame, bg=self.COLOR_PRIMARIO)
        content.place(relx=0.5, rely=0.5, anchor="center")
        
        icon_label = tk.Label(
            content,
            text="📊",
            font=("Segoe UI", 30),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 15))
        
        text_frame = tk.Frame(content, bg=self.COLOR_PRIMARIO)
        text_frame.pack(side=tk.LEFT)
        
        titulo = tk.Label(
            text_frame,
            text="Control de Pagos",
            font=("Segoe UI", 20, "bold"),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        titulo.pack(anchor="w")
        
        subtitulo = tk.Label(
            text_frame,
            text="Sistema de Gestión de Importaciones",
            font=("Segoe UI", 11),
            bg=self.COLOR_PRIMARIO,
            fg="#BDC3C7"
        )
        subtitulo.pack(anchor="w")
    
    def crear_contenido(self, parent):
        """Crea el contenido principal"""
        content_frame = tk.Frame(parent, bg=self.COLOR_FONDO)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        card_frame = tk.Frame(content_frame, bg="white", relief=tk.FLAT, borderwidth=0)
        card_frame.pack(fill=tk.BOTH, expand=True)
        
        self.agregar_sombra(card_frame)
        
        inner_frame = tk.Frame(card_frame, bg="white")
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        section_title = tk.Label(
            inner_frame,
            text="📅 Selección de Fecha de Proyección",
            font=("Segoe UI", 14, "bold"),
            bg="white",
            fg=self.COLOR_PRIMARIO
        )
        section_title.pack(pady=(0, 5))
        
        separator = tk.Frame(inner_frame, height=2, bg=self.COLOR_SECUNDARIO)
        separator.pack(fill=tk.X, pady=(0, 20))
        
        desc_label = tk.Label(
            inner_frame,
            text="Selecciona la fecha para la cual deseas generar la proyección de pagos.\nPor defecto, se sugiere el próximo miércoles.",
            font=("Segoe UI", 10),
            bg="white",
            fg="#7F8C8D",
            justify=tk.CENTER
        )
        desc_label.pack(pady=(0, 25))
        
        cal_frame = tk.Frame(inner_frame, bg="white")
        cal_frame.pack(pady=10)
        
        self.calendario = DateEntry(
            cal_frame,
            width=22,
            background=self.COLOR_SECUNDARIO,
            foreground='white',
            borderwidth=2,
            font=("Segoe UI", 12),
            date_pattern='dd/mm/yyyy',
            locale='es_ES',
            selectbackground=self.COLOR_ACENTO,
            selectforeground='white'
        )
        self.calendario.pack(pady=10)
        
        proximo_miercoles = self.obtener_proximo_miercoles(datetime.now())
        self.calendario.set_date(proximo_miercoles)
        
        fecha_label = tk.Label(
            inner_frame,
            text=f"Fecha sugerida: {proximo_miercoles.strftime('%d/%m/%Y')}",
            font=("Segoe UI", 9, "italic"),
            bg="white",
            fg="#95A5A6"
        )
        fecha_label.pack(pady=(5, 0))
    
    def crear_footer(self, parent):
        """Crea el footer con botones"""
        footer_frame = tk.Frame(parent, bg=self.COLOR_FONDO, height=80)
        footer_frame.pack(fill=tk.X, pady=(0, 20))
        footer_frame.pack_propagate(False)
        
        button_frame = tk.Frame(footer_frame, bg=self.COLOR_FONDO)
        button_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        btn_cancelar = tk.Button(
            button_frame,
            text="Cancelar",
            command=self.cancelar,
            font=("Segoe UI", 11),
            bg="white",
            fg=self.COLOR_TEXTO,
            relief=tk.FLAT,
            borderwidth=2,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        btn_cancelar.pack(side=tk.LEFT, padx=10)
        
        btn_continuar = tk.Button(
            button_frame,
            text="Continuar ➔",
            command=self.continuar,
            font=("Segoe UI", 11, "bold"),
            bg=self.COLOR_ACENTO,
            fg="white",
            relief=tk.FLAT,
            borderwidth=0,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        btn_continuar.pack(side=tk.LEFT, padx=10)
        
        def on_enter_cancelar(e):
            btn_cancelar['bg'] = "#BDC3C7"
        
        def on_leave_cancelar(e):
            btn_cancelar['bg'] = "white"
        
        def on_enter_continuar(e):
            btn_continuar['bg'] = "#229954"
        
        def on_leave_continuar(e):
            btn_continuar['bg'] = self.COLOR_ACENTO
        
        btn_cancelar.bind("<Enter>", on_enter_cancelar)
        btn_cancelar.bind("<Leave>", on_leave_cancelar)
        btn_continuar.bind("<Enter>", on_enter_continuar)
        btn_continuar.bind("<Leave>", on_leave_continuar)
    
    def agregar_sombra(self, widget):
        """Simula sombra con marcos"""
        for i in range(3):
            shade = tk.Frame(
                widget.master,
                bg=f"#{220-i*20:02x}{220-i*20:02x}{220-i*20:02x}"
            )
            shade.place(
                in_=widget,
                x=i+2,
                y=i+2,
                relwidth=1,
                relheight=1
            )
            shade.lower(widget)
    
    def centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def obtener_proximo_miercoles(self, fecha):
        """Calcula el próximo miércoles"""
        dias_hasta_miercoles = (2 - fecha.weekday()) % 7
        if dias_hasta_miercoles == 0:
            dias_hasta_miercoles = 7
        return fecha + timedelta(days=dias_hasta_miercoles)
    
    def continuar(self):
        """Procesa el botón continuar"""
        self.fecha_seleccionada = self.calendario.get_date()
        self.ejecutar_proceso = True
        self.root.destroy()
    
    def cancelar(self):
        """Procesa el botón cancelar"""
        self.ejecutar_proceso = False
        self.root.destroy()

class VentanaProgreso:
    """Ventana de progreso durante la ejecución"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Procesando...")
        self.root.geometry("600x400")
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
        
        # Área de log
        log_frame = tk.Frame(main_frame, bg="white", relief=tk.SUNKEN, borderwidth=2)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Text widget
        self.log_text = tk.Text(
            log_frame,
            font=("Consolas", 9),
            bg="white",
            fg="#2C3E50",
            yscrollcommand=scrollbar.set,
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.log_text.yview)
        
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
        """Centra la ventana"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def log(self, mensaje, tipo="INFO"):
        """Agrega mensaje al log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        colores = {
            "INFO": "#3498DB",
            "OK": "#27AE60",
            "WARN": "#F39C12",
            "ERROR": "#E74C3C"
        }
        
        color = colores.get(tipo, "#2C3E50")
        
        self.log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_text.insert(tk.END, f"[{tipo}] ", tipo)
        self.log_text.insert(tk.END, f"{mensaje}\n")
        
        self.log_text.tag_config("timestamp", foreground="#95A5A6")
        self.log_text.tag_config(tipo, foreground=color, font=("Consolas", 9, "bold"))
        
        self.log_text.see(tk.END)
        self.root.update()
    
    def cerrar(self):
        """Cierra la ventana"""
        self.progress.stop()
        self.root.destroy()

class CopiarArchivo:
    """Clase principal para procesar archivos"""

    marcas_global = ['AMERICANINO', 'ESPRIT', 'CHEVIGNON']
    marcas_unifed = ['NAF NAF', 'RIFLE', 'AEO']
    
    def __init__(self, fecha_filtrado, ventana_progreso, rutas_config):
        self.fecha_filtrado = fecha_filtrado
        self.ventana_progreso = ventana_progreso
        
        # Configurar rutas
        self.ruta_origen = rutas_config['origen']
        self.carpeta_proyecciones = rutas_config['proyecciones']
        self.ruta_destino_final = rutas_config['final']
        
        self.log("Configuración de rutas cargada", "OK")
        self.log(f"Origen: {self.ruta_origen.name}", "INFO")
        self.log(f"Proyecciones: {self.carpeta_proyecciones}", "INFO")
        self.log(f"Final: {self.ruta_destino_final.name}", "INFO")
    
    def log(self, mensaje, tipo="INFO"):
        """Registra mensaje en log"""
        self.ventana_progreso.log(mensaje, tipo)
        try:
            print(f"[{tipo}] {mensaje}")
        except UnicodeEncodeError:
            # Si falla la codificación, intentamos imprimir reemplazando caracteres
            print(f"[{tipo}] {mensaje.encode('ascii', 'replace').decode('ascii')}")
    
    def crear_estructura_carpetas(self, fecha):
        """Crea estructura de carpetas por año/mes/semana"""
        self.log("Creando estructura de carpetas", "INFO")
        
        año = fecha.year
        mes = fecha.strftime('%B').upper()
        
        # Calcular número de semana del mes
        primer_dia_mes = fecha.replace(day=1)
        dias_transcurridos = (fecha - primer_dia_mes).days
        semana = (dias_transcurridos // 7) + 1
        
        carpeta_destino = self.carpeta_proyecciones / f"AÑO {año}" / mes
        carpeta_destino.mkdir(parents=True, exist_ok=True)
        
        self.log(f"Carpeta creada: {carpeta_destino}", "OK")
        return carpeta_destino
    
    def crear_nombre_archivo(self, fecha):
        """Crea nombre del archivo basado en fecha de proyección"""
        dia = fecha.strftime('%d')
        mes = fecha.strftime('%B').upper()
        año = fecha.strftime('%Y')

        nombre = f"{dia} {mes} {año}.xlsx"
        return nombre

    def crear_nombre_segunda_hoja(self, fecha):
        """Crea nombre de segunda hoja: 'MES dia'"""
        mes = fecha.strftime('%B').upper()
        dia = fecha.strftime('%d')

        nombre = f"{mes} {dia}"
        return nombre
    
    def copiar_archivo_base(self, ruta_destino):
        """
        Copia archivo base a destino guardándolo como XLSX y eliminando hojas innecesarias.
        """
        self.log("Copiando archivo completo como .xlsx...", "INFO")
        
        excel = None
        wb = None
        
        try:
            pythoncom.CoInitialize()
            
            # Crear instancia de Excel
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3  # Desactivar macros
            
            # Abrir archivo origen
            self.log(f"Abriendo archivo: {self.ruta_origen.name}", "INFO")
            wb = excel.Workbooks.Open(
                str(self.ruta_origen),
                ReadOnly=True,
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
                Notify=False
            )
            
            # FORZAR visibilidad de todas las hojas ANTES de guardar
            self.log("Haciendo visibles todas las hojas...", "INFO")
            for sheet in wb.Sheets:
                try:
                    sheet.Visible = -1  # xlSheetVisible
                    self.log(f"  + '{sheet.Name}' visible", "INFO")
                except Exception as e:
                    self.log(f"  ! No se pudo hacer visible '{sheet.Name}': {e}", "WARN")
            
            # GUARDAR COMO .XLSX
            ruta_dest_str = str(Path(ruta_destino).resolve())
            self.log(f"Guardando como .xlsx: {Path(ruta_destino).name}", "INFO")
            
            # FileFormat 51 = xlsx (sin macros)
            wb.SaveAs(
                Filename=ruta_dest_str,
                FileFormat=51,
                CreateBackup=False
            )
            
            self.log("Archivo guardado como .xlsx", "OK")
            
            # Cerrar el archivo original
            wb.Close(SaveChanges=False)
            wb = None
            
            # ABRIR EL NUEVO ARCHIVO para limpiar hojas
            self.log("Abriendo archivo nuevo para limpieza...", "INFO")
            wb = excel.Workbooks.Open(ruta_dest_str)
            
            # Identificar hoja de Control de Pagos
            hoja_control = None
            for sheet in wb.Sheets:
                if 'CONTROL' in sheet.Name.upper() and 'PAGOS' in sheet.Name.upper():
                    hoja_control = sheet
                    break
            
            if not hoja_control:
                # Si no encuentra por nombre, intenta la primera visible
                hoja_control = wb.Sheets(1)
                self.log("No se identificó hoja por nombre, usando la primera hoja.", "WARN")
            
            nombre_hoja_control = hoja_control.Name
            self.log(f"Hoja objetivo identificada: '{nombre_hoja_control}'", "OK")
            
            # ELIMINAR todas las hojas EXCEPTO la que necesitamos
            self.log("Eliminando hojas innecesarias...", "INFO")
            
            # Recolectar nombres primero para evitar problemas al iterar y borrar
            hojas_a_eliminar = []
            for sheet in wb.Sheets:
                if sheet.Name != nombre_hoja_control:
                    hojas_a_eliminar.append(sheet.Name)
            
            for nombre_hoja in hojas_a_eliminar:
                try:
                    wb.Sheets(nombre_hoja).Delete()
                    self.log(f"  - Eliminada: '{nombre_hoja}'", "INFO")
                except Exception as e:
                    self.log(f"  ! No se pudo eliminar '{nombre_hoja}': {e}", "WARN")
            
            # Verificar
            if wb.Sheets.Count == 1:
                self.log(f"Archivo limpio. Solo queda: '{wb.Sheets(1).Name}'", "OK")
            else:
                self.log(f"Advertencia: Quedaron {wb.Sheets.Count} hojas", "WARN")
            
            # GUARDAR cambios
            wb.Save()
            self.log("Cambios guardados", "OK")
            
        except Exception as e:
            self.log(f"ERROR al copiar archivo: {e}", "ERROR")
            raise
            
        finally:
            if wb:
                try: 
                    wb.Close(SaveChanges=False) 
                except: pass
            if excel:
                try: 
                    excel.Quit() 
                except: pass
            try: 
                pythoncom.CoUninitialize() 
            except: pass
        
    def leer_datos_control_pagos(self, ruta_archivo):
        """Lee datos del archivo Excel con manejo robusto de encoding"""
        self.log("Leyendo datos del archivo", "INFO")
        
        try:
            # Usar win32com para leer, que maneja mejor los caracteres especiales
            pythoncom.CoInitialize()
            excel = None
            wb = None
            
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                archivo_abs = str(ruta_archivo.absolute())
                self.log(f"Intentando abrir: {archivo_abs}", "INFO")
                
                wb = excel.Workbooks.Open(archivo_abs)
                
                if wb is None:
                    raise Exception("No se pudo abrir el libro (objeto wb es None)")
                
                # Buscar hoja CONTROL DE PAGOS
                # Como ya limpiamos, debería ser la primera
                if wb.Sheets.Count == 0:
                    raise Exception("El libro no tiene hojas")
                    
                ws = wb.Sheets(1)
                
                self.log(f"Hoja encontrada: {ws.Name}", "OK")
                
                # Leer datos
                used_range = ws.UsedRange
                data = used_range.Value
                
                if data:
                    # Convertir tuplas a lista de listas para manejar None
                    data_list = [list(row) if row is not None else [] for row in data]
                    
                    # Asegurarse de que la primera fila tenga encabezados válidos
                    if len(data_list) > 1:
                        headers = data_list[0]
                        # Limpiar encabezados None
                        headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(headers)]
                        
                        df = pd.DataFrame(data_list[1:], columns=headers)
                        self.log(f"Registros leídos: {len(df)}", "OK")
                        return df
                    else:
                        self.log("Archivo sin datos suficientes", "WARN")
                        return None
                else:
                    self.log("No se encontraron datos", "WARN")
                    return None
                    
            finally:
                if wb:
                    wb.Close(SaveChanges=False)
                if excel:
                    excel.Quit()
                pythoncom.CoUninitialize()
                
        except Exception as e:
            self.log(f"Error al leer datos: {str(e)}", "ERROR")
            traceback.print_exc()
            return None
    
    def filtrar_por_fecha(self, df, fecha_filtrado):
        """Filtra registros por fecha de proyección"""
        self.log(f"Filtrando por fecha de proyección: {fecha_filtrado}", "PROCESO")
        
        columnas_normalizadas = []
        for col in df.columns:
            col_normalizado = str(col).strip().upper()
            columnas_normalizadas.append(col_normalizado)
        
        df.columns = columnas_normalizadas
        
        self.log(f"Columnas disponibles: {df.columns.tolist()}", "INFO")
        
        # --- ESTRATEGIA DE FILTRADO ---
        # 1. Primero filtrar por FECHA (FECHA DE VENCIMIENTO o FECHA DE PAGO)
        # 2. Luego filtrar por ESTADO = "PAGAR"
        
        df_procesado = df.copy()
        
        # PASO 1: Filtrar por FECHA
        col_fecha = None
        if 'FECHA DE VENCIMIENTO' in df_procesado.columns:
            col_fecha = 'FECHA DE VENCIMIENTO'
        elif 'FECHA DE PAGO' in df_procesado.columns:
            col_fecha = 'FECHA DE PAGO'
            
        if col_fecha:
            self.log(f"Usando columna de fecha: '{col_fecha}'", "INFO")
            
            # --- CORRECCIÓN FINAL: Limpieza EXTREMA antes de cualquier operación ---
            # El error ocurría al intentar leer sample_vals ANTES de limpiar
            # porque .dropna() o .head() intentan iterar sobre los objetos corruptos
            
            try:
                # 1. Convertir TODO a string primero
                df_procesado[col_fecha] = df_procesado[col_fecha].astype(str)
                
                # 2. Reemplazar valores nulos textuales por NaN real
                vals_nulos = ['None', 'nan', 'NaT', '<NA>', '', 'NaT', 'NoneType']
                df_procesado[col_fecha] = df_procesado[col_fecha].replace(vals_nulos, pd.NA)
                
                # 3. Convertir a datetime PRIMERO (FIX: esto debe ir ANTES de mostrar muestras)
                df_procesado[col_fecha] = pd.to_datetime(df_procesado[col_fecha], errors='coerce')
                
                # 4. AHORA es seguro mostrar muestras
                fechas_validas = df_procesado[col_fecha].dropna()
                if len(fechas_validas) > 0:
                    sample_vals = fechas_validas.head(5).dt.strftime('%Y-%m-%d').tolist()
                    self.log(f"Muestra de fechas (limpias): {sample_vals}", "INFO")
                
            except Exception as e:
                self.log(f"Error crítico limpiando fechas: {e}", "ERROR")
                return pd.DataFrame()
            
            # Calcular rango de semana
            fecha_referencia = fecha_filtrado.date() if isinstance(fecha_filtrado, datetime) else fecha_filtrado
            inicio_semana = fecha_referencia - timedelta(days=fecha_referencia.weekday())
            fin_semana = inicio_semana + timedelta(days=6)
            
            self.log(f"Filtrando fechas en rango: {inicio_semana} al {fin_semana}", "INFO")
            
            # Filtrar por rango
            df_fechas = df_procesado.dropna(subset=[col_fecha])
            
            # --- DEBUG: Ver fechas interpretadas ---
            if not df_fechas.empty:
                fechas_interpretadas = df_fechas[col_fecha].dt.date.unique()[:5]
                self.log(f"Fechas interpretadas (ejemplos): {fechas_interpretadas}", "INFO")
            
            df_fechas = df_fechas[
                (df_fechas[col_fecha].dt.date >= inicio_semana) & 
                (df_fechas[col_fecha].dt.date <= fin_semana)
            ]
            
            registros_fecha = len(df_fechas)
            self.log(f"Registros en fecha ({inicio_semana} - {fin_semana}): {registros_fecha}", "INFO")
            
            if registros_fecha == 0:
                self.log(f"No se encontraron registros en la fecha indicada.", "WARN")
                return pd.DataFrame()

            # PASO 2: Filtrar por ESTADO (sobre los resultados de fecha)
            if 'ESTADO' in df_fechas.columns:
                self.log("Aplicando filtro secundario: ESTADO contiene 'PAGAR'", "INFO")
                # Normalizar estado
                df_fechas['ESTADO_NORM'] = df_fechas['ESTADO'].astype(str).str.upper().str.strip()
                
                # Filtrar
                df_final = df_fechas[df_fechas['ESTADO_NORM'].str.contains('PAGAR', na=False)].copy()
                
                registros_finales = len(df_final)
                self.log(f"Registros finales (Fecha + Estado PAGAR): {registros_finales}", "OK")
                
                if registros_finales == 0:
                    self.log("Advertencia: Hay registros en la fecha pero NINGUNO con estado PAGAR", "WARN")
                    estados_encontrados = df_fechas['ESTADO'].unique()
                    self.log(f"Estados encontrados en esa fecha: {estados_encontrados}", "INFO")
                    # Opcional: retornar vacío o lo que había por fecha
                    return pd.DataFrame()
                
                # Limpiar columna auxiliar
                if 'ESTADO_NORM' in df_final.columns:
                    df_final = df_final.drop(columns=['ESTADO_NORM'])
                    
                return df_final
            else:
                self.log("No se encontró columna ESTADO. Retornando filtro solo por fecha.", "WARN")
                return df_fechas

        else:
            self.log("No se encontró columna de fecha compatible.", "ERROR")
            return pd.DataFrame()
    
    def preparar_datos_segunda_hoja(self, df):
        """Prepara datos para la segunda hoja con todas las columnas necesarias"""
        self.log(f"Preparando datos para proyección...", "PROCESO")
        
        df_resultado = pd.DataFrame()
        
        # Mapeo de columnas
        columnas_necesarias = {
            'IMPORTADOR': 'IMPORTADOR',
            'MARCA': 'MARCA',
            'PROVEEDOR': 'PROVEEDOR',
            'NRO. IMPO': 'NRO. IMPO',
            'VALOR A PAGAR': 'VALOR A PAGAR',
            'MONEDA': 'MONEDA',
            'NOTA CRÉDITO': 'VALOR NOTA CRÉDITO',
        }
        
        for col_dest, col_origen in columnas_necesarias.items():
            if col_origen in df.columns:
                df_resultado[col_dest] = df[col_origen]
            else:
                # Valores por defecto
                if col_dest == 'NOTA CRÉDITO':
                    df_resultado[col_dest] = 0
                else:
                    df_resultado[col_dest] = ''
        
        # Limpiar nombres de marca (quitar prefijo "COMODIN S.A.S - " y normalizar)
        if 'MARCA' in df_resultado.columns:
            def normalizar_marca(valor):
                valor = str(valor).strip()
                # Quitar prefijo
                if 'COMODIN S.A.S - ' in valor:
                    valor = valor.replace('COMODIN S.A.S - ', '')
                
                # Normalizar nombres específicos (completar nombres)
                valor_upper = valor.upper()
                if 'NAF' in valor_upper:
                    return 'NAF NAF'
                elif 'ESPRIT' in valor_upper:
                    return 'ESPRIT'
                elif 'CHEVI' in valor_upper:
                    return 'CHEVIGNON'
                elif 'AMERICANINO' in valor_upper:
                    return 'AMERICANINO'
                elif 'RIFLE' in valor_upper:
                    return 'RIFLE'
                elif 'AEO' in valor_upper:
                    return 'AEO'
                
                return valor

            df_resultado['MARCA'] = df_resultado['MARCA'].apply(normalizar_marca)
        
        
        # Asegurar que VALOR A PAGAR es numérico
        df_resultado['VALOR A PAGAR'] = pd.to_numeric(df_resultado['VALOR A PAGAR'], errors='coerce').fillna(0)
        
        return df_resultado
    
    def agrupar_y_calcular(self, df):
        """Agrupa por PROVEEDOR y prepara filas con subtotales y estilos"""
        self.log(f"Agrupando registros...", "PROCESO")
        
        # Asegurar valores numéricos
        df['VALOR A PAGAR'] = pd.to_numeric(df['VALOR A PAGAR'], errors='coerce').fillna(0)
        df['NOTA CRÉDITO'] = pd.to_numeric(df['NOTA CRÉDITO'], errors='coerce').fillna(0)
        
        # Ordenar por PROVEEDOR y MARCA
        df = df.sort_values(by=['PROVEEDOR', 'MARCA']).reset_index(drop=True)
        
        # Preparar resultado con metadatos para formato
        filas_resultado = []
        grupos = df.groupby('PROVEEDOR', sort=False)

        fila_actual = 2
        letra_columna = 'E'

        for proveedor, grupo in grupos:
            # Agregar filas del grupo
            fila_inicio = fila_actual
            for idx, row in grupo.iterrows():
                fila_dict = row.to_dict()
                fila_dict['_TIPO'] = 'DETALLE'
                filas_resultado.append(fila_dict)
                fila_actual += 1
            
            # Si hay más de una fila, agregar subtotal
            if len(grupo) > 1:
                fila_fin = fila_actual - 1
                # Usar SUBTOTAL(9, ...) para que el Total Final pueda usar SUBTOTAL(9, ...) y evitar duplicados
                suma_valor = f'=SUBTOTAL(9, {letra_columna}{fila_inicio}:{letra_columna}{fila_fin})'
                moneda = grupo['MONEDA'].iloc[0]
                
                fila_subtotal = {col: '' for col in df.columns}
                fila_subtotal['VALOR A PAGAR'] = suma_valor
                fila_subtotal['MONEDA'] = moneda
                fila_subtotal['_TIPO'] = 'SUBTOTAL'
                filas_resultado.append(fila_subtotal)
                fila_actual += 1
                
                # Agregar 1 fila en blanco después del subtotal
                fila_blanco = {col: '' for col in df.columns}
                fila_blanco['_TIPO'] = 'BLANCO'
                filas_resultado.append(fila_blanco)
                fila_actual += 1

            else:
                # Si es un solo registro, marcar como único para formato verde
                # Modificar el último registro agregado (que es el detalle único)
                if filas_resultado:
                    filas_resultado[-1]['_TIPO'] = 'DETALLE_UNICO'
                
                # Agregar 2 filas en blanco para mantener espaciado
                fila_blanco = {col: '' for col in df.columns}
                fila_blanco['_TIPO'] = 'BLANCO'
                filas_resultado.append(fila_blanco)
                filas_resultado.append(fila_blanco.copy())
                fila_actual += 2
        
        return pd.DataFrame(filas_resultado)
    
    def guardar_proyeccion_com(self, ruta_archivo, df_agrupado, nombre_hoja):
        """Guarda proyección con formato y estilos correctos"""
        self.log("Guardando proyección", "INFO")
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            # Usar DispatchEx para nueva instancia segura y evitar conflictos
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Verificar si el archivo existe antes de abrir
            if not ruta_archivo.exists():
                raise FileNotFoundError(f"El archivo no existe: {ruta_archivo}")

            # Intentar abrir el archivo con reintentos por si está bloqueado
            for i in range(3):
                try:
                    wb = excel.Workbooks.Open(str(ruta_archivo.absolute()))
                    break
                except Exception as e:
                    if i == 2: raise e
                    self.log(f"Intento {i+1} fallido al abrir archivo, reintentando...", "WARN")
                    import time
                    time.sleep(1)
            
            if wb is None:
                raise Exception("No se pudo abrir el libro de Excel (wb is None)")

            # Crear nueva hoja
            ws = wb.Sheets.Add()
            ws.Name = nombre_hoja
            
            # ENCABEZADOS
            headers = ['IMPORTADOR', 'MARCA', 'PROVEEDOR', 'NRO. IMPO', 'VALOR A PAGAR', 'MONEDA', 'NOTA CRÉDITO']
            
            for col_idx, header in enumerate(headers, 1):
                cell = ws.Cells(1, col_idx)
                cell.Value = header
                cell.Font.Bold = True
                cell.Font.Color = 0xFFFFFF  # Blanco
                cell.Interior.Color = 0x993366  # Morado
                cell.HorizontalAlignment = -4108  # xlCenter
            
            # ESCRIBIR DATOS CON FORMATO
            fila_actual = 2
            filas_blanco = []
            
            # Columnas para los datos
            columnas_orden = ['IMPORTADOR', 'MARCA', 'PROVEEDOR', 'NRO. IMPO', 'VALOR A PAGAR', 'MONEDA', 'NOTA CRÉDITO']
            
            for idx, row in df_agrupado.iterrows():
                tipo_fila = row.get('_TIPO', 'DETALLE')
                
                if tipo_fila == 'BLANCO':
                    # Fila en blanco
                    filas_blanco.append(fila_actual)
                    fila_actual += 1
                    continue
                
                # Escribir datos según el tipo de fila
                if tipo_fila == 'SUBTOTAL':
                    # FILA DE SUBTOTAL (verde claro)
                    ws.Cells(fila_actual, 5).Formula = row['VALOR A PAGAR']
                    ws.Cells(fila_actual, 6).Value = row['MONEDA']
                    
                    # Formato número
                    ws.Cells(fila_actual, 5).NumberFormat = "$ #,##0.00"
                    
                    # Aplicar formato verde
                    for col in range(5, 7):
                        cell = ws.Cells(fila_actual, col)
                        cell.Interior.Color = 0xCCFFCC  # Verde claro
                        cell.Font.Bold = True
                    
                elif tipo_fila == 'DETALLE_UNICO':
                    # FILA DE DETALLE ÚNICO (Datos + Verde)
                    ws.Cells(fila_actual, 1).Value = str(row.get('IMPORTADOR', ''))
                    ws.Cells(fila_actual, 2).Value = str(row.get('MARCA', ''))
                    ws.Cells(fila_actual, 3).Value = str(row.get('PROVEEDOR', ''))
                    ws.Cells(fila_actual, 4).Value = str(row.get('NRO. IMPO', ''))
                    
                    val_pagar = row.get('VALOR A PAGAR', 0)
                    ws.Cells(fila_actual, 5).Value = float(val_pagar) if val_pagar else 0
                    
                    ws.Cells(fila_actual, 6).Value = str(row.get('MONEDA', ''))
                    
                    val_nota = row.get('NOTA CRÉDITO', 0)
                    ws.Cells(fila_actual, 7).Value = float(val_nota) if val_nota else 0
                    
                    # Formato número (Moneda con 2 decimales)
                    ws.Cells(fila_actual, 5).NumberFormat = "$ #,##0.00"
                    ws.Cells(fila_actual, 7).NumberFormat = "$ #,##0.00"
                    
                    # Aplicar formato verde a columnas de valor
                    for col in range(5, 7):
                        cell = ws.Cells(fila_actual, col)
                        cell.Interior.Color = 0xCCFFCC  # Verde claro
                        cell.Font.Bold = True
                        
                else:
                    # FILA DE DETALLE NORMAL
                    ws.Cells(fila_actual, 1).Value = str(row.get('IMPORTADOR', ''))
                    ws.Cells(fila_actual, 2).Value = str(row.get('MARCA', ''))
                    ws.Cells(fila_actual, 3).Value = str(row.get('PROVEEDOR', ''))
                    ws.Cells(fila_actual, 4).Value = str(row.get('NRO. IMPO', ''))
                    
                    val_pagar = row.get('VALOR A PAGAR', 0)
                    ws.Cells(fila_actual, 5).Value = float(val_pagar) if val_pagar else 0
                    
                    ws.Cells(fila_actual, 6).Value = str(row.get('MONEDA', ''))
                    
                    val_nota = row.get('NOTA CRÉDITO', 0)
                    ws.Cells(fila_actual, 7).Value = float(val_nota) if val_nota else 0
                    
                    # Formato número (Moneda con 2 decimales)
                    ws.Cells(fila_actual, 5).NumberFormat = "$ #,##0.00"
                    ws.Cells(fila_actual, 7).NumberFormat = "$ #,##0.00"
                
                fila_actual += 1
            
            # TOTAL FINAL AL FINAL DE LA LISTA
            # Usamos SUBTOTAL(9, ...) para sumar todo el rango ignorando otros SUBTOTALES
            ultima_fila_datos = fila_actual - 1
            formula_total_final = f'=SUBTOTAL(9, E2:E{ultima_fila_datos})'
            
            ws.Cells(fila_actual, 4).Value = "TOTAL"
            ws.Cells(fila_actual, 5).Formula = formula_total_final
            
            # Formato Total Final
            ws.Cells(fila_actual, 4).Font.Bold = True
            ws.Cells(fila_actual, 5).Font.Bold = True
            ws.Cells(fila_actual, 5).NumberFormat = "$ #,##0.00"
            ws.Cells(fila_actual, 4).Interior.Color = 0xCCFFCC
            ws.Cells(fila_actual, 5).Interior.Color = 0xCCFFCC
            
            # Autoajustar columnas
            ws.Columns.AutoFit()
            
            # Aplicar bordes a todas las celdas con datos (Columnas 1 a 7)
            rango_datos = ws.Range(ws.Cells(1, 1), ws.Cells(fila_actual - 1, 7))
            rango_datos.Borders.LineStyle = 1  # xlContinuous
            rango_datos.Borders.Weight = 2  # xlThin
            
            # Quitar bordes a las filas en blanco
            for fila in filas_blanco:
                ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 7)).Borders.LineStyle = -4142  # xlNone
            
            # -------------------------------------------------------------------------
            # TABLA DE RESUMEN (AL LADO DE LA TABLA)
            # -------------------------------------------------------------------------
            col_resumen = 10  # Columna J
            fila_resumen = 2  # Empezar arriba
            
            # Definir rangos para SUMIF
            rango_criterio = f'B2:B{ultima_fila_datos}'
            rango_suma = f'E2:E{ultima_fila_datos}'
            
            def generar_formula_sumif(marcas):
                partes = []
                for marca in marcas:
                    partes.append(f'SUMIF({rango_criterio}, "{marca}", {rango_suma})')
                if not partes:
                    return "0"
                return f'={"+".join(partes)}'

            # Ajustar marcas si es necesario (CHEVIÑON -> CHEVIGNON)
            marcas_global_norm = [m for m in self.marcas_global]

            # Global
            formula_global = generar_formula_sumif(marcas_global_norm)
            
            # Unified (Excluyendo AEO)
            marcas_unified_clean = [m for m in self.marcas_unifed if m != 'AEO']
            formula_unified = generar_formula_sumif(marcas_unified_clean)
            
            # AEO
            formula_aeo = f'=SUMIF({rango_criterio}, "AEO", {rango_suma})'
            
            # Escribir Tabla Resumen
            # Global
            ws.Cells(fila_resumen, col_resumen).Value = "Global"
            ws.Cells(fila_resumen, col_resumen + 1).Formula = formula_global
            ws.Cells(fila_resumen, col_resumen).Interior.Color = 0xBDD7EE  # Azul claro
            ws.Cells(fila_resumen, col_resumen + 1).Interior.Color = 0xBDD7EE
            
            # Unified
            ws.Cells(fila_resumen + 1, col_resumen).Value = "Unified"
            ws.Cells(fila_resumen + 1, col_resumen + 1).Formula = formula_unified
            ws.Cells(fila_resumen + 1, col_resumen).Interior.Color = 0xFFE699  # Naranja claro
            ws.Cells(fila_resumen + 1, col_resumen + 1).Interior.Color = 0xFFE699
            
            # AEO
            ws.Cells(fila_resumen + 2, col_resumen).Value = "AEO"
            ws.Cells(fila_resumen + 2, col_resumen + 1).Formula = formula_aeo
            ws.Cells(fila_resumen + 2, col_resumen).Interior.Color = 0xCCFFCC  # Verde claro
            ws.Cells(fila_resumen + 2, col_resumen + 1).Interior.Color = 0xCCFFCC
            
            # Gran Total Resumen (Suma de las 3 celdas anteriores)
            letra_col_suma = "K" # Si col_resumen es 10 (J), suma es 11 (K)
            formula_gran_total_resumen = f'=SUM({letra_col_suma}{fila_resumen}:{letra_col_suma}{fila_resumen + 2})'
            
            ws.Cells(fila_resumen + 3, col_resumen + 1).Formula = formula_gran_total_resumen
            ws.Cells(fila_resumen + 3, col_resumen + 1).Font.Bold = True
            ws.Cells(fila_resumen + 3, col_resumen + 1).Borders(8).LineStyle = 1 # xlEdgeTop
            
            # Formato moneda para el resumen
            rango_resumen = ws.Range(ws.Cells(fila_resumen, col_resumen + 1), ws.Cells(fila_resumen + 3, col_resumen + 1))
            rango_resumen.NumberFormat = "$ #,##0.00"

            # Aplicar bordes a la tabla de resumen
            rango_tabla_resumen = ws.Range(ws.Cells(fila_resumen, col_resumen), ws.Cells(fila_resumen + 3, col_resumen + 1))
            rango_tabla_resumen.Borders.LineStyle = 1  # xlContinuous
            rango_tabla_resumen.Borders.Weight = 2  # xlThin

            wb.Save()
            self.log("Proyección guardada exitosamente", "OK")
            
        except Exception as e:
            self.log(f"Error al guardar proyección: {str(e)}", "ERROR")
            raise
            
        finally:
            if wb:
                wb.Close()
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def preparar_df_final(self, df_detalle):
        """Prepara DataFrame para archivo final"""
        self.log("Preparando datos para archivo final", "INFO")
        
        df_final = pd.DataFrame()
        fecha_proyeccion = self.fecha_filtrado
        
        df_final['IMPORTADOR'] = df_detalle['IMPORTADOR']
        df_final['MARCA'] = df_detalle['MARCA']
        df_final['FECHA DE PAGO'] = fecha_proyeccion.strftime('%d/%m/%Y')
        df_final['DIA'] = fecha_proyeccion.day
        df_final['MES'] = fecha_proyeccion.month
        df_final['AÑO'] = fecha_proyeccion.year
        df_final['PROVEEDOR'] = df_detalle['PROVEEDOR']
        df_final['# IMPORTACION'] = df_detalle['NRO. IMPO']
        df_final['VALOR MONEDA ORIGEN'] = df_detalle['VALOR A PAGAR']
        df_final['MONEDA'] = df_detalle['MONEDA']
        
        def calc_valor_usd(row):
            if str(row['MONEDA']).upper() == 'USD':
                return row['VALOR A PAGAR']
            return ''
        
        def calc_factor(row):
            if str(row['MONEDA']).upper() == 'USD':
                return 1
            return ''
        
        df_final['VALOR USD'] = df_detalle.apply(calc_valor_usd, axis=1)
        df_final['FACTOR DE CONVERSION'] = df_detalle.apply(calc_factor, axis=1)
        df_final['DESCUENTO PRONTO PAGO'] = 0
        df_final['FORMA DE PAGO'] = ''
        df_final['TIPO DE PAGO'] = 'CUENTA COMPENSACION'
        df_final['FECHA DE APERTURA CREDITO -UTILIZACION LC'] = 'N/A'
        df_final['FECHA DE VENCIMIENTO'] = 'N/A'
        df_final['# CREDITO'] = 'N/A'
        df_final['# DEUDA EXTERNA'] = 'N/A'
        df_final['NOTA CREDITO'] = 0.00
        df_final['OBSERVACIONES'] = ''
        
        self.log(f"DataFrame final preparado: {len(df_final)} registros", "OK")
        return df_final
    
    def anexar_archivo_final_com(self, df_detalle):
        """Anexa datos al archivo final usando COM con mayor robustez"""
        self.log("Anexando datos al archivo final", "INFO")
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            # Usamos DispatchEx para asegurar una instancia limpia y aislada
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            archivo_path = str(self.ruta_destino_final.absolute())
            wb = excel.Workbooks.Open(archivo_path)
            
            # --- MEJORA: Búsqueda flexible de la hoja ---
            ws = None
            nombres_hojas = [sheet.Name.upper() for sheet in wb.Sheets]
            self.log(f"Hojas en destino: {nombres_hojas}", "INFO")

            # Buscamos coincidencias más amplias
            for sheet in wb.Sheets:
                nombre_u = sheet.Name.upper()
                if 'PAGOS' in nombre_u and 'IMPOR' in nombre_u:
                    ws = sheet
                    break
            
            if not ws:
                ws = wb.Sheets(1)
                self.log(f"No se halló hoja 'Pagos importación', usando: '{ws.Name}'", "WARN")
            else:
                self.log(f"Escribiendo en hoja confirmada: '{ws.Name}'", "OK")

            # Preparar datos (limpieza de nulos para Excel)
            datos = df_detalle.fillna("").values.tolist()
            if not datos:
                self.log("No hay datos para anexar", "WARN")
                return
            
            # --- MEJORA: Encontrar última fila real ---
            # Si la columna A está vacía en la fila 1, empezamos en 1.
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            if last_row == 1 and ws.Cells(1, 1).Value is None:
                start_row = 1
            else:
                start_row = last_row + 1

            self.log(f"Insertando {len(datos)} registros en fila {start_row}", "INFO")
            
            # ESCRIBIR POR RANGO (Mucho más rápido y seguro que el bucle anterior)
            num_filas = len(datos)
            num_cols = len(datos[0])
            
            # Definimos el rango exacto: desde A(start_row) hasta la última columna y fila
            destino = ws.Range(
                ws.Cells(start_row, 1), 
                ws.Cells(start_row + num_filas - 1, num_cols)
            )
            destino.Value = datos
            
            # Forzar guardado
            wb.Save()
            self.log(f"¡ÉXITO! Datos guardados en {self.ruta_destino_final.name}", "OK")
            
        except Exception as e:
            self.log(f"Error crítico al anexar: {str(e)}", "ERROR")
            raise
            
        finally:
            if wb:
                wb.Close(SaveChanges=True)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def agregar_a_archivo_final(self, df_detalle):
        """Proceso completo para agregar a archivo final"""
        try:
            df_final = self.preparar_df_final(df_detalle)
            self.anexar_archivo_final_com(df_final)
        except Exception as e:
            self.log(f"Error en archivo final: {str(e)}", "ERROR")
    
    def ejecutar_proceso(self):
        """Ejecuta el proceso completo"""
        print("\n" + "="*80)
        print("    AUTOMATIZACIÓN DE CONTROL DE PAGOS - VERSIÓN 1.0 ")
        print("="*80 + "\n")
        
        try:
            if not self.ruta_origen.exists():
                self.log(f"No se encuentra el archivo original", "ERROR")
                messagebox.showerror("Error", f"No se encuentra el archivo:\n{self.ruta_origen}")
                return None
            
            fecha_proyeccion = self.fecha_filtrado
            self.log(f"Fecha de proyección: {fecha_proyeccion.strftime('%d/%m/%Y')}", "INFO")
            
            carpeta_destino = self.crear_estructura_carpetas(fecha_proyeccion)
            nombre_archivo = self.crear_nombre_archivo(fecha_proyeccion)
            ruta_archivo_nuevo = carpeta_destino / nombre_archivo
            
            self.copiar_archivo_base(ruta_archivo_nuevo)
            
            # Esperar un momento para liberar el archivo
            time.sleep(2)
            
            df_original = self.leer_datos_control_pagos(ruta_archivo_nuevo)
            if df_original is None:
                return None
            
            df_filtrado = self.filtrar_por_fecha(df_original, fecha_proyeccion)
            
            if len(df_filtrado) == 0:
                self.log("No se encontraron registros", "WARN")
                messagebox.showwarning("Sin registros", "No se encontraron registros para la fecha seleccionada.")
                return None
            
            df_segunda = self.preparar_datos_segunda_hoja(df_filtrado)
            df_agrupado = self.agrupar_y_calcular(df_segunda)
            
            nombre_segunda_hoja = self.crear_nombre_segunda_hoja(fecha_proyeccion)
            
            self.guardar_proyeccion_com(ruta_archivo_nuevo, df_agrupado, nombre_segunda_hoja)
            
            self.agregar_a_archivo_final(df_segunda)
            
            print("\n" + "="*80)
            print("PROCESO COMPLETADO EXITOSAMENTE")
            print("="*80)
            
            messagebox.showinfo(
                "¡Proceso Completado!",
                f"El proceso ha finalizado exitosamente.\n\n"
                f"📁 Proyección guardada en:\n{ruta_archivo_nuevo}\n\n"
                f"📁 Archivo final actualizado:\n{self.ruta_destino_final.name}"
            )
            return str(ruta_archivo_nuevo)
            
        except Exception as e:
            self.log(f"ERROR CRÍTICO: {str(e)}", "ERROR")
            traceback.print_exc()
            messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
            return None

def main():
    """Función principal"""
    try:
        # Configurar rutas
        configurador = ConfiguradorRutas()
        if not configurador.cargar_o_crear_config():
            return
        
        rutas = configurador.obtener_rutas()
        
        # Mostrar ventana de selección de fecha
        interfaz = InterfazModerna()
        interfaz.crear_ventana()
        
        if not interfaz.ejecutar_proceso:
            return
        
        # Confirmar ejecución
        if not messagebox.askyesno(
            "Confirmar Ejecución",
            "Antes de continuar, asegúrese de:\n\n"
            "✓ Haber actualizado el archivo 'CONTROL DE PAGOS.xlsm'\n"
            "✓ Haber guardado todos los cambios\n"
            "✓ Cerrar el archivo si está abierto\n\n"
            "¿Desea continuar?"
        ):
            return
        
        # Crear ventana de progreso
        ventana_prog = VentanaProgreso()
        
        try:
            # Ejecutar proceso
            copiador = CopiarArchivo(
                fecha_filtrado=interfaz.fecha_seleccionada,
                ventana_progreso=ventana_prog,
                rutas_config=rutas
            )
            resultado = copiador.ejecutar_proceso()
            
            # Cerrar ventana de progreso
            ventana_prog.cerrar()
            
        except Exception as e:
            ventana_prog.cerrar()
            messagebox.showerror("Error Fatal", f"Error inesperado:\n\n{str(e)}")
    
    except Exception as e:
        messagebox.showerror("Error de Configuración", f"Error al iniciar:\n\n{str(e)}")

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
            root.withdraw()
            messagebox.showerror("Error Fatal", f"Ocurrió un error crítico:\n\n{str(e)}\n\nConsulte CRASH_LOG.txt")
        except:
            print(f"Error fatal: {e}")
            input("Presione Enter para salir...")