"""
PROCESO MENSUAL - Control de Pagos
Lógica específica para proyecciones mensuales
"""
from pathlib import Path
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
import time
import traceback
import stat
import sys


def obtener_ruta_recurso(nombre_archivo: str) -> Path:
    """Obtiene ruta de recurso en ejecutable o desarrollo"""
    base_dir = getattr(sys, "_MEIPASS", None)
    if base_dir:
        return Path(base_dir) / nombre_archivo
    return Path(__file__).resolve().parent / nombre_archivo


def aplicar_icono_ventana(ventana: tk.Tk) -> None:
    """Aplica ícono a ventana si existe"""
    try:
        icon_path = obtener_ruta_recurso("icon.ico")
        if icon_path.exists():
            ventana.iconbitmap(str(icon_path))
    except Exception:
        pass


class InterfazMensual:
    """
    Interfaz gráfica para seleccionar el MES (a través de una fecha)
    """
    def __init__(self):
        self.fecha_seleccionada = None
        self.ejecutar_proceso = False
        
        self.COLOR_PRIMARIO = "#3498DB" # Morado para diferenciar del semanal
        self.COLOR_SECUNDARIO = "#9B59B6"
        self.COLOR_ACENTO = "#2ECC71"
        self.COLOR_FONDO = "#ECF0F1"
        self.COLOR_TEXTO = "#2C3E50"
        
    def crear_ventana(self):
        self.root = tk.Tk()
        aplicar_icono_ventana(self.root)
        self.root.title("Control de Pagos Mensual")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        self.root.configure(bg=self.COLOR_FONDO)
        
        self.centrar_ventana()
        self.configurar_estilos()
        
        main_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        self.crear_header(main_frame)
        self.crear_contenido(main_frame)
        self.crear_footer(main_frame)
        
        self.root.mainloop()
    
    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
    
    def crear_header(self, parent):
        header_frame = tk.Frame(parent, bg=self.COLOR_PRIMARIO, height=140)
        header_frame.pack(fill=tk.X, pady=0)
        header_frame.pack_propagate(False)
        
        content = tk.Frame(header_frame, bg=self.COLOR_PRIMARIO)
        content.place(relx=0.5, rely=0.5, anchor="center")
        
        icon_label = tk.Label(
            content,
            text="📅",
            font=("Segoe UI", 30),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 15))
        
        text_frame = tk.Frame(content, bg=self.COLOR_PRIMARIO)
        text_frame.pack(side=tk.LEFT)
        
        titulo = tk.Label(
            text_frame,
            text="Proyección Mensual",
            font=("Segoe UI", 20, "bold"),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        titulo.pack(anchor="w")
        
        subtitulo = tk.Label(
            text_frame,
            text="Generación de reporte completo del mes",
            font=("Segoe UI", 11),
            bg=self.COLOR_PRIMARIO,
            fg="#E8DAEF"
        )
        subtitulo.pack(anchor="w")
    
    def crear_contenido(self, parent):
        content_frame = tk.Frame(parent, bg=self.COLOR_FONDO)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        card_frame = tk.Frame(content_frame, bg="white", relief=tk.FLAT, borderwidth=0)
        card_frame.pack(fill=tk.BOTH, expand=True)
        
        self.agregar_sombra(card_frame)
        
        inner_frame = tk.Frame(card_frame, bg="white")
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        section_title = tk.Label(
            inner_frame,
            text="Selección de Mes",
            font=("Segoe UI", 14, "bold"),
            bg="white",
            fg=self.COLOR_PRIMARIO
        )
        section_title.pack(pady=(0, 5))
        
        separator = tk.Frame(inner_frame, height=2, bg=self.COLOR_SECUNDARIO)
        separator.pack(fill=tk.X, pady=(0, 20))
        
        desc_label = tk.Label(
            inner_frame,
            text="Selecciona el Mes y Año que deseas procesar.",
            font=("Segoe UI", 10),
            bg="white",
            fg="#7F8C8D",
            justify=tk.CENTER
        )
        desc_label.pack(pady=(0, 25))
        
        cal_frame = tk.Frame(inner_frame, bg="white")
        cal_frame.pack(pady=10)
        
        # Selectores Mes / Año
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        
        # Frame Mes
        frame_mes = tk.Frame(cal_frame, bg="white")
        frame_mes.pack(side=tk.LEFT, padx=10)
        tk.Label(frame_mes, text="Mes:", bg="white", font=("Segoe UI", 10)).pack(anchor="w")
        self.combo_mes = ttk.Combobox(frame_mes, values=meses, state="readonly", width=15, font=("Segoe UI", 11))
        self.combo_mes.pack(pady=(5,0))
        # Seleccionar mes actual
        self.combo_mes.current(datetime.now().month - 1)

        # Frame Año
        frame_anio = tk.Frame(cal_frame, bg="white")
        frame_anio.pack(side=tk.LEFT, padx=10)
        
        anio_actual = datetime.now().year
        anios = [str(a) for a in range(anio_actual - 2, anio_actual + 5)]
        
        tk.Label(frame_anio, text="Año:", bg="white", font=("Segoe UI", 10)).pack(anchor="w")
        self.combo_anio = ttk.Combobox(frame_anio, values=anios, state="readonly", width=10, font=("Segoe UI", 11))
        self.combo_anio.pack(pady=(5,0))
        self.combo_anio.set(str(anio_actual))
    
    def crear_footer(self, parent):
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
            text="Generar Reporte ➔",
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
    
    def agregar_sombra(self, widget):
        for i in range(3):
            shade = tk.Frame(
                widget.master,
                bg=f"#{220-i*20:02x}{220-i*20:02x}{220-i*20:02x}"
            )
            shade.place(in_=widget, x=i+2, y=i+2, relwidth=1, relheight=1)
            shade.lower(widget)
    
    def centrar_ventana(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def continuar(self):
        mes_idx = self.combo_mes.current() + 1
        anio = int(self.combo_anio.get())
        
        # Crear fecha del primer día del mes seleccionado
        self.fecha_seleccionada = datetime(anio, mes_idx, 1)
        
        self.ejecutar_proceso = True
        self.root.destroy()
    
    def cancelar(self):
        self.ejecutar_proceso = False
        self.root.destroy()

class VentanaProgreso:
    def __init__(self):
        self.root = tk.Tk()
        aplicar_icono_ventana(self.root)
        self.root.title("Procesando...")
        self.root.geometry("600x300")
        self.root.resizable(False, False)
        self.root.configure(bg="#ECF0F1")
        self.centrar_ventana()
        
        main_frame = tk.Frame(self.root, bg="#ECF0F1")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        titulo = tk.Label(
            main_frame,
            text="⚙️ Generando Proyección Mensual",
            font=("Segoe UI", 16, "bold"),
            bg="#ECF0F1",
            fg="#2C3E50"
        )
        titulo.pack(pady=(0, 20))
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=560)
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
        try:
            print(f"[{tipo}] {mensaje}")
        except:
            pass
        self.root.update()
    
    def cerrar(self):
        self.progress.stop()
        self.root.destroy()

class ProcesadorMensual:
    def __init__(self, fecha_filtrado, ventana_progreso, rutas_config):
        self.fecha_filtrado = fecha_filtrado
        self.ventana_progreso = ventana_progreso
        
        global pd, win32com, pythoncom, openpyxl, Font, Alignment, Border, Side
        import pandas as pd
        import win32com.client as win32com
        import pythoncom
        import openpyxl
        from openpyxl.styles import Font, Alignment, Border, Side

        self.ruta_origen = rutas_config['origen']
        self.carpeta_proyecciones = rutas_config['proyecciones']
        self.ruta_destino_final = rutas_config['final']
        
        self.log("Configuración cargada", "INFO")

    def log(self, mensaje, tipo="INFO"):
        self.ventana_progreso.log(mensaje, tipo)

    def crear_estructura_carpetas(self, fecha):
        año = fecha.year
        mes = fecha.strftime('%B').upper()
        carpeta_destino = self.carpeta_proyecciones / f"AÑO {año}" / mes
        carpeta_destino.mkdir(parents=True, exist_ok=True)
        return carpeta_destino
    
    def crear_nombre_archivo(self, fecha):
        # Nombre solo el MES, ej: "FEBRERO.xlsx"
        mes = fecha.strftime('%B').upper()
        nombre = f"{mes}.xlsx"
        return nombre

    def crear_nombre_segunda_hoja(self, fecha):
        return fecha.strftime('%B').upper()
    
    def copiar_archivo_base(self, ruta_destino):
        """Copia archivo base a destino guardándolo como XLSX y eliminando hojas innecesarias."""
        self.log("Copiando archivo completo como .xlsx...", "INFO")
        
        excel = None
        wb = None
        
        try:
            pythoncom.CoInitialize()
            
            # Crear instancia de Excel
            excel = win32com.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3  # Desactivar macros
            
            # Abrir archivo origen
            self.log(f"Abriendo archivo: {self.ruta_origen.name}", "INFO")
            wb = excel.Workbooks.Open(
                str(self.ruta_origen),
                ReadOnly=False,  # Abrir en modo escritura
                UpdateLinks=3,   # Actualizar vínculos
                IgnoreReadOnlyRecommended=True,
                Notify=False
            )
            
            self.log("Actualizando vínculos y recalculando fórmulas...", "INFO")
            try:
                wb.RefreshAll()
                excel.CalculateFull()
                time.sleep(3)
                self.log("Datos actualizados correctamente", "OK")
            except Exception as e:
                self.log(f"Advertencia al actualizar: {e}", "WARN")

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
            
            wb.SaveAs(
                Filename=ruta_dest_str,
                FileFormat=51, # xlOpenXMLWorkbook
                CreateBackup=False
            )
            
            self.log("Archivo guardado como .xlsx", "OK")
            
            # Cerrar el archivo original
            wb.Close(SaveChanges=False)
            wb = None
            
            # Permisos
            try:
                ruta_dest_path = Path(ruta_dest_str)
                ruta_dest_path.chmod(ruta_dest_path.stat().st_mode | stat.S_IWRITE)
                self.log("Permisos de escritura habilitados en el archivo creado.", "OK")
            except Exception as e:
                self.log(f"No se pudieron ajustar permisos de escritura: {e}", "WARN")
            
            # ABRIR EL NUEVO ARCHIVO para limpiar hojas
            self.log("Abriendo archivo nuevo para limpieza...", "INFO")
            wb = excel.Workbooks.Open(
                ruta_dest_str,
                ReadOnly=False,
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
                Notify=False
            )
            
            # Identificar hoja de Control de Pagos
            hoja_control = None
            for sheet in wb.Sheets:
                if 'CONTROL' in sheet.Name.upper() and 'PAGOS' in sheet.Name.upper():
                    hoja_control = sheet
                    break
            
            if not hoja_control:
                hoja_control = wb.Sheets(1)
                self.log("No se identificó hoja por nombre, usando la primera hoja.", "WARN")
            
            nombre_hoja_control = hoja_control.Name
            self.log(f"Hoja objetivo identificada: '{nombre_hoja_control}'", "OK")
            
            # Eliminar hojas extra
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
            
            wb.Save()
            self.log("Archivo base preparado correctamente", "OK")
            
        except Exception as e:
            self.log(f"Error copiando: {e}", "ERROR")
            raise
        finally:
            if wb: 
                try: wb.Close(SaveChanges=False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()

    def leer_datos_proceso_mensual(self, ruta_archivo):
        self.log("Leyendo datos...", "INFO")
        try:
            pythoncom.CoInitialize()
            excel = win32com.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(str(ruta_archivo.absolute()))
            ws = wb.Sheets(1)
            data = ws.UsedRange.Value
            wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()
            
            if data and len(data) > 1:
                data_list = [list(row) if row is not None else [] for row in data]
                headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(data_list[0])]
                df = pd.DataFrame(data_list[1:], columns=headers)
                return df
            return None
        except Exception as e:
            self.log(f"Error leyendo: {e}", "ERROR")
            return None

    def filtrar_por_fecha(self, df, fecha_referencia):
        self.log("Filtrando datos del mes...", "PROCESO")
        
        # Normalizar columnas
        df.columns = [str(col).strip().upper() for col in df.columns]
        
        # Identificar columna de fecha
        col_fecha = None
        if 'FECHA DE VENCIMIENTO' in df.columns:
            col_fecha = 'FECHA DE VENCIMIENTO'
        elif 'FECHA DE PAGO' in df.columns:
            col_fecha = 'FECHA DE PAGO'
            
        if not col_fecha:
            return pd.DataFrame()
            
        # Limpiar y convertir fechas
        df[col_fecha] = df[col_fecha].astype(str)
        vals_nulos = ['None', 'nan', 'NaT', '<NA>', '', 'NaT', 'NoneType']
        df[col_fecha] = df[col_fecha].replace(vals_nulos, pd.NA)
        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
        
        df = df.dropna(subset=[col_fecha])
        
        # Filtrar por MES completo
        fecha_date = fecha_referencia.date() if isinstance(fecha_referencia, datetime) else fecha_referencia
        mes_target = fecha_date.month
        anio_target = fecha_date.year
        
        df_mes = df[
            (df[col_fecha].dt.month == mes_target) & 
            (df[col_fecha].dt.year == anio_target)
        ].copy()
        
        self.log(f"Registros encontrados para el mes: {len(df_mes)}", "INFO")
        
        # NO filtramos por estado (PAGAR/PENDIENTE), se traen todos
        return df_mes

    def preparar_datos_segunda_hoja(self, df):
        self.log("Preparando columnas...", "PROCESO")
        df_resultado = pd.DataFrame()
        
        # Columnas solicitadas
        mapa_columnas = {
            'FECHA INGRESO': 'FECHA INGRESO',
            'IMPORTADOR': 'IMPORTADOR',
            'MARCA': 'MARCA',
            'PROVEEDOR': 'PROVEEDOR',
            'NRO. IMPO': 'NRO. IMPO',
            'MONEDA': 'MONEDA',
            'N ° FACTURA/PROFORMA': ['N ° FACTURA', 'N° FACTURA', 'FACTURA', 'PROFORMA', 'N ° FACTURA/PROFORMA'],
            'FECHA FACTURA / PROFORMA': ['FECHA FACTURA', 'FECHA DE FACTURA', 'FECHA PROFORMA', 'FECHA FACTURA / PROFORMA'],
            'VALOR FACTURA O PROFORMA': ['VALOR FACTURA', 'VALOR PROFORMA', 'VALOR FACTURA O PROFORMA'],
            'VALOR NOTA CRÉDITO': ['VALOR NOTA CRÉDITO', 'NOTA CRÉDITO', 'NOTA DE CRÉDITO'],
            'VALOR A PAGAR': 'VALOR A PAGAR',
            'ESTADO': 'ESTADO',
            'FECHA DE VENCIMIENTO': 'FECHA DE VENCIMIENTO',
            'FECHA DE PAGO': 'FECHA DE PAGO',
            'FECHA DOCUMENTO DE TRANSPORTE': ['FECHA DOCUMENTO DE TRANSPORTE', 'FECHA DOC TRANSPORTE'],
            'TRM MIGO': 'TRM MIGO',
            'TIPO DE PAGO': 'TIPO DE PAGO'
        }
        
        for col_dest, col_origen_posibles in mapa_columnas.items():
            if isinstance(col_origen_posibles, list):
                # Buscar la primera coincidencia
                col_encontrada = None
                for c in col_origen_posibles:
                    if c in df.columns:
                        col_encontrada = c
                        break
                if col_encontrada:
                    df_resultado[col_dest] = df[col_encontrada]
                else:
                    df_resultado[col_dest] = ''
            else:
                if col_origen_posibles in df.columns:
                    df_resultado[col_dest] = df[col_origen_posibles]
                else:
                    df_resultado[col_dest] = ''
        
        # Normalizar nombres de MARCA (Igual que en semanal)
        if 'MARCA' in df_resultado.columns:
            def normalizar_marca(valor):
                valor = str(valor).strip()

                if 'COMODIN S.A.S - ' in valor:
                    valor = valor.replace('COMODIN S.A.S - ', '')
                
                # Normalizar nombres 
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
            
        # Limpiezas numéricas
        cols_numericas = ['VALOR FACTURA O PROFORMA', 'VALOR NOTA CRÉDITO', 'VALOR A PAGAR', 'TRM MIGO']
        for col in cols_numericas:
            if col in df_resultado.columns:
                df_resultado[col] = pd.to_numeric(df_resultado[col], errors='coerce').fillna(0)
                
        # Limpieza de fechas para display
        cols_fechas = ['FECHA INGRESO', 'FECHA FACTURA / PROFORMA', 'FECHA DE VENCIMIENTO', 'FECHA DE PAGO', 'FECHA DOCUMENTO DE TRANSPORTE']
        for col in cols_fechas:
            if col in df_resultado.columns:
                try:
                    # Asegurar conversion a string primero para evitar errores con NoneType
                    df_resultado[col] = df_resultado[col].astype(str)
                    
                    # Limpiar valores nulos textuales
                    vals_nulos = ['None', 'nan', 'NaT', '<NA>', '', 'NaT', 'NoneType']
                    for val in vals_nulos:
                        df_resultado[col] = df_resultado[col].replace(val, pd.NA)
                    
                    # Convertir a datetime
                    df_resultado[col] = pd.to_datetime(df_resultado[col], errors='coerce')
                except Exception as e:
                    self.log(f"Advertencia: No se pudo convertir fecha en {col}: {e}", "WARN")
        
        return df_resultado

    def agrupar_y_calcular(self, df):
        self.log("Agrupando por Semana y Proveedor...", "PROCESO")
        
        # 1. Determinar Semana (ISO)
        if 'FECHA DE VENCIMIENTO' in df.columns and not df['FECHA DE VENCIMIENTO'].isnull().all():
            df['Semana_Abs'] = df['FECHA DE VENCIMIENTO'].dt.isocalendar().week
        else:
            df['Semana_Abs'] = 0
            
        # Mapear a semanas relativas del mes (1, 2, 3...)
        # Esto corrige que aparezcan semanas 6, 7, 8... en febrero
        semanas_unicas = sorted(df['Semana_Abs'].unique())
        mapa_semanas = {s: i+1 for i, s in enumerate(semanas_unicas)}
        df['Semana'] = df['Semana_Abs'].map(mapa_semanas)
        df.drop(columns=['Semana_Abs'], inplace=True)
            
        # Ordenar por Semana, IMPORTADOR y luego Proveedor
        df = df.sort_values(by=['Semana', 'IMPORTADOR', 'PROVEEDOR']).reset_index(drop=True)
        
        filas_resultado = []
        grupos_semana = df.groupby('Semana', sort=False)
        
        fila_actual = 2
        
        # Definir columna de Valor para sumas (index 10 = K, si A=0)
        col_valor_pagar_letra = 'K' # 11va columna
        
        for semana, grupo_sem in grupos_semana:
            # Dentro de la semana, agrupar por IMPORTADOR y PROVEEDOR
            grupos_prov = grupo_sem.groupby(['IMPORTADOR', 'PROVEEDOR'], sort=False)
            
            fila_inicio_semana = fila_actual
            
            for (importador, proveedor), grupo_prov in grupos_prov:
                fila_inicio_prov = fila_actual
                
                # Filas de detalle
                for idx, row in grupo_prov.iterrows():
                    fila_dict = row.to_dict()
                    fila_dict['_TIPO'] = 'DETALLE'
                    filas_resultado.append(fila_dict)
                    fila_actual += 1
                
                # Subtotal Proveedor (si hay más de 1 registro)
                if len(grupo_prov) > 1:
                    fila_fin_prov = fila_actual - 1
                    formula = f'=SUBTOTAL(9, {col_valor_pagar_letra}{fila_inicio_prov}:{col_valor_pagar_letra}{fila_fin_prov})'
                    
                    fila_sub = {c: '' for c in df.columns}
                    # fila_sub['PROVEEDOR'] = f"TOTAL {proveedor}" # Eliminado por solicitud
                    fila_sub['VALOR A PAGAR'] = formula
                    # fila_sub['MONEDA'] = ... # Eliminado por solicitud
                    fila_sub['_TIPO'] = 'SUBTOTAL_PROV'
                    filas_resultado.append(fila_sub)
                    fila_actual += 1
                    
                    # Espacio (1 fila vacía)
                    fila_blanco = {c: '' for c in df.columns}
                    fila_blanco['_TIPO'] = 'BLANCO'
                    filas_resultado.append(fila_blanco)
                    fila_actual += 1
                else:
                    # Detalle único, espacio
                    if filas_resultado:
                        filas_resultado[-1]['_TIPO'] = 'DETALLE_UNICO'
                    
                    # Espacio (1 fila vacía para consistencia)
                    fila_blanco = {c: '' for c in df.columns}
                    fila_blanco['_TIPO'] = 'BLANCO'
                    filas_resultado.append(fila_blanco)
                    fila_actual += 1

            # Total Semana
            fila_fin_semana = fila_actual - 1
            # Usamos SUBTOTAL(9) que ignora otros subtotales(9)
            formula_semana = f'=SUBTOTAL(9, {col_valor_pagar_letra}{fila_inicio_semana}:{col_valor_pagar_letra}{fila_fin_semana})'
            
            fila_total_sem = {c: '' for c in df.columns}
            # Etiqueta en columna anterior a VALOR A PAGAR (VALOR NOTA CRÉDITO)
            fila_total_sem['PROVEEDOR'] = f"TOTAL SEMANA {semana}" 
            fila_total_sem['VALOR A PAGAR'] = formula_semana
            fila_total_sem['_TIPO'] = 'TOTAL_SEMANA'
            filas_resultado.append(fila_total_sem)
            fila_actual += 1
            
            # Espacio grande entre semanas (3 filas vacías)
            fila_blanco = {c: '' for c in df.columns}
            fila_blanco['_TIPO'] = 'BLANCO'
            filas_resultado.append(fila_blanco)
            filas_resultado.append(fila_blanco.copy())
            filas_resultado.append(fila_blanco.copy())
            fila_actual += 3
            
        return pd.DataFrame(filas_resultado), df

    def _convertir_fecha_excel(self, val):
        """Convierte valores a datetime de Python seguros para Excel"""
        if pd.isna(val) or val == "" or val is None:
            return None
        
        try:
            if isinstance(val, (pd.Timestamp, datetime)):
                d = val
            else:
                # Intentar parsear si es string
                d = pd.to_datetime(val, dayfirst=True, errors='coerce')
                if pd.isna(d):
                    return str(val) # Si falla, devolver como texto
            
            # Convertir a python datetime nativo
            py_dt = d.to_pydatetime().replace(tzinfo=None)
            
            # Excel no soporta fechas antes de 1900
            # Devolver como string formateado dd/mm/yyyy
            if py_dt.year < 1900:
                return py_dt.strftime("%d/%m/%Y")
                
            return py_dt
        except:
            # En caso de cualquier error, devolver como string
            return str(val)

    def guardar_proyeccion_com(self, ruta_archivo, df_agrupado, nombre_hoja, df_origen=None):
        self.log("Escribiendo Excel...", "INFO")
        pythoncom.CoInitialize()
        # Usar Dispatch en lugar de DispatchEx para intentar reutilizar instancia si es posible,
        # o mantener DispatchEx si queremos aislamiento. 
        # El error "La llamada fue rechazada" suele ser por sobrecarga de mensajes.
        # Vamos a intentar un enfoque con reintentos para la escritura.
        try:
            excel = win32com.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
        except:
            # Fallback
            excel = win32com.Dispatch("Excel.Application")
            
        wb = None
        
        try:
            wb = excel.Workbooks.Open(str(ruta_archivo.absolute()), ReadOnly=False, UpdateLinks=0)
            ws = wb.Sheets.Add()
            ws.Name = nombre_hoja
            
            # Encabezados
            headers = [
                'FECHA INGRESO', 'IMPORTADOR', 'MARCA', 'PROVEEDOR', 'NRO. IMPO', 'MONEDA', 
                'N ° FACTURA/PROFORMA', 'FECHA FACTURA / PROFORMA', 'VALOR FACTURA O PROFORMA', 
                'VALOR NOTA CRÉDITO', 'VALOR A PAGAR', 'ESTADO', 'FECHA DE VENCIMIENTO', 
                'FECHA DE PAGO', 'FECHA DOCUMENTO DE TRANSPORTE', 'TRM MIGO', 'TIPO DE PAGO'
            ]
            
            # Índices de columnas (1-based)
            # Fechas: 1, 8, 13, 14, 15
            # Numéricas: 9, 10, 11, 16
            cols_fecha = {1, 8, 13, 14, 15}
            cols_num = {9, 10, 11, 16}
            
            # Estilo Encabezados
            for i, h in enumerate(headers, 1):
                cell = ws.Cells(1, i)
                cell.Value = h
                cell.Font.Bold = True
                cell.Font.Color = 0xFFFFFF
                cell.Interior.Color = 0x8E44AD # Morado
                cell.HorizontalAlignment = -4108 # Center
                cell.Borders.LineStyle = 1 # Borde
            
            # Escribir datos con retardo para evitar saturar COM
            fila = 2
            col_mapping = {col: i+1 for i, col in enumerate(headers)}
            
            # Convertir DataFrame a lista de diccionarios para iteración más rápida
            # y pre-procesar datos
            datos_lista = df_agrupado.to_dict('records')
            
            for row in datos_lista:
                tipo = row.get('_TIPO', 'DETALLE')
                
                if tipo == 'BLANCO':
                    fila += 1
                    continue
                
                # Iterar columnas explícitamente y limpiar tipos
                for col_name, col_idx in col_mapping.items():
                    raw_val = row.get(col_name, '')
                    
                    try:
                        # Si es fórmula (empieza por =), escribir directamente
                        if isinstance(raw_val, str) and raw_val.startswith('='):
                            ws.Cells(fila, col_idx).Formula = raw_val
                            continue

                        val_final = None
                        
                        if col_idx in cols_num:
                            # Numéricos
                            try:
                                # Si es string vacío explícito (como en subtotales), dejar None para que quede celda vacía
                                if raw_val == '':
                                    val_final = None
                                elif pd.isna(raw_val):
                                    val_final = 0.0
                                else:
                                    val_final = float(raw_val)
                            except:
                                val_final = 0.0
                            
                        elif col_idx in cols_fecha:
                            # Fechas
                            val_final = self._convertir_fecha_excel(raw_val)
                                
                        else:
                            # Texto
                            if pd.isna(raw_val) or raw_val is None:
                                val_final = ""
                            else:
                                val_final = str(raw_val)

                        if val_final is not None:
                            ws.Cells(fila, col_idx).Value = val_final
                            
                    except Exception as cell_err:
                        # Si falla, intentar una vez más como texto tras breve pausa
                        time.sleep(0.01)
                        try:
                            ws.Cells(fila, col_idx).Value = str(raw_val) if raw_val is not None else ""
                        except:
                            pass # Ignorar si falla segunda vez
                
                # Formatos específicos (solo si no es blanco)
                # Moneda: VALOR FACTURA (9), VALOR NC (10), VALOR A PAGAR (11), TRM (16)
                for c_idx in [9, 10, 11]:
                    ws.Cells(fila, c_idx).NumberFormat = "$ #,##0.00"
                ws.Cells(fila, 16).NumberFormat = "#,##0.00"
                
                # Formato Fechas: 1, 8, 13, 14, 15
                for c_idx in [1, 8, 13, 14, 15]:
                    ws.Cells(fila, c_idx).NumberFormat = "dd/mm/yyyy"

                # Estilos y Bordes por tipo
                if tipo == 'DETALLE':
                    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 17)).Borders.LineStyle = 1
                elif tipo == 'SUBTOTAL_PROV':
                    cell_sub = ws.Cells(fila, 11)
                    cell_sub.Interior.Color = 0xD4EFDF
                    cell_sub.Font.Bold = True
                    cell_sub.Borders.LineStyle = 1
                elif tipo == 'TOTAL_SEMANA':
                    # Ahora la etiqueta está en PROVEEDOR (Col 4) y valor en Col 11
                    # Resaltar la fila en esas columnas
                    ws.Cells(fila, 4).Font.Bold = True
                    ws.Cells(fila, 4).Interior.Color = 0xD7BDE2
                    ws.Cells(fila, 4).Borders.LineStyle = 1
                    
                    cell_val = ws.Cells(fila, 11)
                    cell_val.Interior.Color = 0xD7BDE2
                    cell_val.Font.Bold = True
                    cell_val.Borders.LineStyle = 1
                elif tipo == 'DETALLE_UNICO':
                    ws.Cells(fila, 11).Interior.Color = 0xD4EFDF
                    ws.Cells(fila, 11).Font.Bold = True
                    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 17)).Borders.LineStyle = 1

                fila += 1
                
                # Pausa breve cada 50 filas para dejar respirar al proceso COM
                if fila % 50 == 0:
                    time.sleep(0.05)
            
            ws.Columns.AutoFit()
            
            # Total General al final
            fila_total = fila
            
            # Escribir texto asegurando que se vea
            cell_total_title = ws.Cells(fila_total, 10)
            cell_total_title.Value = "TOTAL GENERAL DEL MES"
            
            # Fórmula
            cell_total_val = ws.Cells(fila_total, 11)
            cell_total_val.Formula = f'=SUBTOTAL(9, K2:K{fila_total-1})'
            
            # Estilos SOLO en celdas 10 y 11
            rango_total = ws.Range(cell_total_title, cell_total_val)
            rango_total.Font.Bold = True
            rango_total.Interior.Color = 0x8E44AD
            rango_total.Font.Color = 0xFFFFFF
            rango_total.Borders.LineStyle = 1
            cell_total_val.NumberFormat = "$ #,##0.00"
            
            # Generar tablas resumen por Semana y Marca (si se proporcionó df_origen)
            if df_origen is not None and 'Semana' in df_origen.columns:
                fila += 4 # Dejar espacio
                
                weeks = sorted(df_origen['Semana'].unique())
                
                for sem in weeks:
                    # Filtrar y Agrupar
                    df_sem = df_origen[df_origen['Semana'] == sem]
                    resumen = df_sem.groupby('MARCA')['VALOR A PAGAR'].sum().reset_index()
                    
                    # Definir columnas para la tabla resumen: B (2) y C (3) o C y D
                    # Vamos a usar Col 3 (Marca) y Col 4 (Valor) para alinear visualmente
                    c_marca = 3
                    c_valor = 4
                    
                    # Encabezado "Semana X"
                    cell_header = ws.Cells(fila, c_marca)
                    cell_header.Value = f"Semana {sem}"
                    cell_header.Font.Bold = True
                    cell_header.Interior.Color = 0xD7BDE2 # Morado suave
                    cell_header.Borders.LineStyle = 1
                    cell_header.HorizontalAlignment = -4108 # Center
                    
                    # Unir celdas de encabezado si se desea, o dejar solo en una
                    ws.Range(ws.Cells(fila, c_marca), ws.Cells(fila, c_valor)).Merge()
                    ws.Range(ws.Cells(fila, c_marca), ws.Cells(fila, c_valor)).Borders.LineStyle = 1
                    
                    fila += 1
                    
                    total_sem_resumen = 0.0
                    
                    for _, r in resumen.iterrows():
                        marca = str(r['MARCA'])
                        try:
                            valor = float(r['VALOR A PAGAR'])
                        except:
                            valor = 0.0
                        
                        total_sem_resumen += valor
                        
                        # Nombre Marca
                        ws.Cells(fila, c_marca).Value = marca
                        ws.Cells(fila, c_marca).Borders.LineStyle = 1
                        
                        # Valor
                        ws.Cells(fila, c_valor).Value = valor
                        ws.Cells(fila, c_valor).NumberFormat = "$ #,##0.00"
                        ws.Cells(fila, c_valor).Borders.LineStyle = 1
                        
                        fila += 1
                        
                    # Total de la tabla
                    ws.Cells(fila, c_marca).Value = "Total"
                    ws.Cells(fila, c_marca).Font.Bold = True
                    ws.Cells(fila, c_marca).Borders.LineStyle = 1
                    
                    ws.Cells(fila, c_valor).Value = total_sem_resumen
                    ws.Cells(fila, c_valor).Font.Bold = True
                    ws.Cells(fila, c_valor).NumberFormat = "$ #,##0.00"
                    ws.Cells(fila, c_valor).Borders.LineStyle = 1
                    ws.Cells(fila, c_valor).Interior.Color = 0xD4EFDF # Verde suave
                    
                    fila += 2 # Espacio entre tablas
            
            wb.Save()
            
        except Exception as e:
            self.log(f"Error guardando excel: {e}", "ERROR")
            raise
        finally:
            if wb: wb.Close()
            excel.Quit()
            pythoncom.CoUninitialize()

    def ejecutar_proceso(self):
        try:
            if not self.ruta_origen.exists():
                messagebox.showerror("Error", "No se encuentra el archivo origen")
                return None
            
            self.log("Iniciando proceso...", "INFO")
            
            carpeta = self.crear_estructura_carpetas(self.fecha_filtrado)
            nombre_arch = self.crear_nombre_archivo(self.fecha_filtrado)
            ruta_destino = carpeta / nombre_arch
            
            self.copiar_archivo_base(ruta_destino)
            time.sleep(1)
            
            df = self.leer_datos_proceso_mensual(ruta_destino)
            if df is None: return
            
            df_mes = self.filtrar_por_fecha(df, self.fecha_filtrado)
            if df_mes.empty:
                self.log("No hay datos para este mes", "WARN")
                messagebox.showwarning("Aviso", "No se encontraron registros en el mes seleccionado.")
                return
            
            df_prep = self.preparar_datos_segunda_hoja(df_mes)
            df_agrup, df_con_semana = self.agrupar_y_calcular(df_prep)
            
            nombre_hoja = self.crear_nombre_segunda_hoja(self.fecha_filtrado)
            self.guardar_proyeccion_com(ruta_destino, df_agrup, nombre_hoja, df_con_semana)
            
            messagebox.showinfo("Éxito", f"Proyección mensual creada en:\n{ruta_destino}")
            return str(ruta_destino)
            
        except Exception as e:
            self.log(f"Error fatal: {e}", "ERROR")
            traceback.print_exc()
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")