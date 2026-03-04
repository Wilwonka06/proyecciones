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
import threading

def obtener_ruta_recurso(nombre_archivo: str) -> Path:
    """Obtiene ruta de recurso en ejecutable o desarrollo"""
    base_dir = getattr(sys, "_MEIPASS", None)
    if base_dir:
        return Path(base_dir) / nombre_archivo
    return Path(__file__).resolve().parent / nombre_archivo

class MonthlyFrame(tk.Frame):
    """Frame de interfaz para la proyección mensual"""
    def __init__(self, parent, controller):
        super().__init__(parent, bg=controller.COLOR_FONDO)
        self.controller = controller
        self.setup_ui()

    def setup_ui(self):
        # Main content container
        content = tk.Frame(self, bg=self.controller.COLOR_FONDO)
        content.pack(expand=True, fill=tk.BOTH, padx=50, pady=30)
        
        # Selection Card
        card = tk.Frame(content, bg=self.controller.COLOR_CARTA, padx=40, pady=40, highlightbackground=self.controller.COLOR_BORDE, highlightthickness=1)
        card.place(relx=0.5, rely=0.4, anchor="center")
        
        tk.Label(
            card, 
            text="Configuración de Proyección Mensual", 
            font=("Segoe UI", 16, "bold"), 
            bg=self.controller.COLOR_CARTA, 
            fg=self.controller.COLOR_PRIMARIO
        ).pack(pady=(0, 20))
        
        tk.Label(
            card, 
            text="Selecciona el mes y año para generar el reporte consolidado.", 
            font=("Segoe UI", 10), 
            bg=self.controller.COLOR_CARTA, 
            fg=self.controller.COLOR_TEXTO_CLARO,
            justify=tk.CENTER
        ).pack(pady=(0, 30))
        
        # Selection container
        sel_container = tk.Frame(card, bg=self.controller.COLOR_CARTA)
        sel_container.pack(pady=10)
        
        # Month
        month_frame = tk.Frame(sel_container, bg="#F5EEF8", padx=15, pady=15, highlightbackground="#E8DAEF", highlightthickness=1)
        month_frame.pack(side=tk.LEFT, padx=10)
        
        tk.Label(month_frame, text="🗓️ Mes", font=("Segoe UI", 10, "bold"), bg="#F5EEF8", fg="#8E44AD").pack(pady=(0, 10))
        
        self.meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        
        self.combo_mes = ttk.Combobox(month_frame, values=self.meses, state="readonly", width=15, font=("Segoe UI", 11))
        self.combo_mes.pack()
        self.combo_mes.current(datetime.now().month - 1)
        
        # Year
        year_frame = tk.Frame(sel_container, bg="#EBF5FB", padx=15, pady=15, highlightbackground="#D6EAF8", highlightthickness=1)
        year_frame.pack(side=tk.LEFT, padx=10)
        
        tk.Label(year_frame, text="📅 Año", font=("Segoe UI", 10, "bold"), bg="#EBF5FB", fg="#3498DB").pack(pady=(0, 10))
        
        anio_actual = datetime.now().year
        anios = [str(a) for a in range(anio_actual - 2, anio_actual + 5)]
        
        self.combo_anio = ttk.Combobox(year_frame, values=anios, state="readonly", width=10, font=("Segoe UI", 11))
        self.combo_anio.pack()
        self.combo_anio.set(str(anio_actual))
        
        # Info label
        self.info_lbl = tk.Label(
            card,
            text=f"💡 Se procesarán todos los registros de {self.meses[datetime.now().month-1]} {anio_actual}",
            font=("Segoe UI", 9, "italic"),
            bg=self.controller.COLOR_CARTA,
            fg="#27AE60"
        )
        self.info_lbl.pack(pady=25)
        
        def update_info(e=None):
            m = self.combo_mes.get()
            a = self.combo_anio.get()
            self.info_lbl.config(text=f"💡 Se procesarán todos los registros de {m} {a}")
            
        self.combo_mes.bind("<<ComboboxSelected>>", update_info)
        self.combo_anio.bind("<<ComboboxSelected>>", update_info)
        
        # Action Button
        self.btn_ejecutar = tk.Button(
            card,
            text="Generar Proyección Mensual",
            font=("Segoe UI", 12, "bold"),
            bg="#8E44AD",
            fg="white",
            relief=tk.FLAT,
            padx=50,
            pady=18,
            cursor="hand2",
            activebackground="#7D3C98",
            activeforeground="white",
            bd=0,
            highlightthickness=0,
            command=self.iniciar_proceso
        )
        self.btn_ejecutar.pack(pady=(10, 0))
        
        def on_enter(e): self.btn_ejecutar.configure(bg="#7D3C98", relief=tk.FLAT)
        def on_leave(e): self.btn_ejecutar.configure(bg="#8E44AD", relief=tk.FLAT)
        self.btn_ejecutar.bind("<Enter>", on_enter)
        self.btn_ejecutar.bind("<Leave>", on_leave)

    def iniciar_proceso(self):
        mes_idx = self.combo_mes.current() + 1
        anio = int(self.combo_anio.get())
        fecha = datetime(anio, mes_idx, 1)
        
        if not messagebox.askyesno(
            "Confirmar Ejecución",
            f"¿Desea iniciar la proyección mensual para {self.meses[mes_idx-1]} {anio}?\n\n"
            "Este proceso puede tardar unos minutos."
        ):
            return
            
        self.controller.show_view("PROGRESS")
        progreso = self.controller.current_view
        
        def run():
            try:
                procesador = ProcesadorMensual(
                    fecha_filtrado=fecha,
                    ventana_progreso=progreso,
                    rutas_config=self.controller.rutas
                )
                procesador.ejecutar_proceso()
                progreso.set_progress(100, "✅ Proceso mensual completado")
                progreso.log("Reporte mensual generado correctamente", "OK")
            except Exception as e:
                progreso.log(f"Error fatal: {str(e)}", "ERROR")
                progreso.set_progress(0, "❌ Error en el proceso")
                messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
        
        threading.Thread(target=run, daemon=True).start()

class ProcesadorMensual:
    """Clase principal para procesar archivos mensuales"""
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
        
    def set_progress(self, percent, status=None):
        self.ventana_progreso.set_progress(percent, status)

    def crear_estructura_carpetas(self, fecha):
        año = fecha.year
        mes = fecha.strftime('%B').upper()
        carpeta_destino = self.carpeta_proyecciones / f"AÑO {año}" / mes
        carpeta_destino.mkdir(parents=True, exist_ok=True)
        return carpeta_destino
    
    def crear_nombre_archivo(self, fecha):
        return f"{fecha.strftime('%B').upper()}.xlsx"

    def crear_nombre_segunda_hoja(self, fecha):
        return fecha.strftime('%B').upper()
    
    def copiar_archivo_base(self, ruta_destino):
        self.log("Copiando archivo completo como .xlsx...", "INFO")
        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            
            wb = excel.Workbooks.Open(str(self.ruta_origen), ReadOnly=False, UpdateLinks=3, IgnoreReadOnlyRecommended=True, Notify=False)
            
            try:
                hoja_control = None
                for sheet in wb.Sheets:
                    if sheet.Name.lower() == "control_pagos":
                        hoja_control = sheet
                        break
                if hoja_control:
                    for qt in hoja_control.QueryTables: qt.Refresh(BackgroundQuery=False)
                    for lo in hoja_control.ListObjects:
                        if lo.QueryTable: lo.QueryTable.Refresh(BackgroundQuery=False)
                    for pt in hoja_control.PivotTables(): pt.PivotCache().Refresh()
            except: pass

            for sheet in wb.Sheets:
                try: sheet.Visible = -1
                except: pass
            
            ruta_dest_str = str(Path(ruta_destino).resolve())
            wb.SaveAs(Filename=ruta_dest_str, FileFormat=51, CreateBackup=False)
            wb.Close(SaveChanges=False)
            wb = None
            
            wb = excel.Workbooks.Open(ruta_dest_str, ReadOnly=False, UpdateLinks=0, IgnoreReadOnlyRecommended=True, Notify=False)
            hoja_control = None
            for sheet in wb.Sheets:
                if 'CONTROL' in sheet.Name.upper() and 'PAGOS' in sheet.Name.upper():
                    hoja_control = sheet
                    break
            if not hoja_control: hoja_control = wb.Sheets(1)
            
            hojas_a_eliminar = [sheet.Name for sheet in wb.Sheets if sheet.Name != hoja_control.Name]
            for nombre in hojas_a_eliminar:
                try: wb.Sheets(nombre).Delete()
                except: pass
            wb.Save()
        finally:
            if wb: wb.Close(SaveChanges=False)
            if excel: excel.Quit()
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
        self.set_progress(10, "Iniciando...")
        try:
            carpeta = self.crear_estructura_carpetas(self.fecha_filtrado)
            ruta = carpeta / self.crear_nombre_archivo(self.fecha_filtrado)
            self.copiar_archivo_base(ruta)

            df = self.leer_datos_proceso_mensual(ruta)
            if df is None or df.empty:
                # Crear hoja vacía con encabezados y total 0
                df_vacio = pd.DataFrame(columns=[
                    'FECHA INGRESO','IMPORTADOR','MARCA','PROVEEDOR','NRO. IMPO','MONEDA',
                    'N ° FACTURA/PROFORMA','FECHA FACTURA / PROFORMA','VALOR FACTURA O PROFORMA',
                    'VALOR NOTA CRÉDITO','VALOR A PAGAR','ESTADO','FECHA DE VENCIMIENTO',
                    'FECHA DE PAGO','FECHA DOCUMENTO DE TRANSPORTE','TRM MIGO','TIPO DE PAGO'
                ])
                self.guardar_proyeccion_com(ruta, df_vacio, self.crear_nombre_segunda_hoja(self.fecha_filtrado), df_origen=df_vacio)
                return True

            df_f = self.filtrar_por_fecha(df, self.fecha_filtrado)
            df_s = self.preparar_datos_segunda_hoja(df_f)
            df_a, df_sem_ref = self.agrupar_y_calcular(df_s)
            self.guardar_proyeccion_com(
                ruta,
                df_a,
                self.crear_nombre_segunda_hoja(self.fecha_filtrado),
                df_origen=df_sem_ref
            )
            return True
        except Exception as e:
            raise e
