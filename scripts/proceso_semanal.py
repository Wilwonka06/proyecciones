"""
PROCESO SEMANAL - Control de Pagos
Lógica específica para proyecciones semanales
"""
from pathlib import Path
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
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

class WeeklyFrame(tk.Frame):
    """Frame de interfaz para la proyección semanal"""
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
            text="Configuración de Proyección", 
            font=("Segoe UI", 16, "bold"), 
            bg=self.controller.COLOR_CARTA, 
            fg=self.controller.COLOR_PRIMARIO
        ).pack(pady=(0, 20))
        
        tk.Label(
            card, 
            text="Selecciona la fecha para generar el reporte.\nSe recomienda el próximo miércoles.", 
            font=("Segoe UI", 10), 
            bg=self.controller.COLOR_CARTA, 
            fg=self.controller.COLOR_TEXTO_CLARO,
            justify=tk.CENTER
        ).pack(pady=(0, 30))
        
        # Calendar container
        cal_frame = tk.Frame(card, bg="#F0F7FF", padx=20, pady=20, highlightbackground="#D1E8FF", highlightthickness=1)
        cal_frame.pack(pady=10)
        
        tk.Label(cal_frame, text="📅 Fecha de Proyección", font=("Segoe UI", 10, "bold"), bg="#F0F7FF", fg=self.controller.COLOR_SECUNDARIO).pack(pady=(0, 10))
        
        self.calendario = DateEntry(
            cal_frame,
            width=20,
            background=self.controller.COLOR_SECUNDARIO,
            foreground='white',
            borderwidth=0,
            font=("Segoe UI", 12),
            date_pattern='dd/mm/yyyy',
            locale='es_ES',
            selectbackground=self.controller.COLOR_ACENTO,
            selectforeground='white',
            todaybackground=self.controller.COLOR_PRIMARIO,
            normalbackground='white'
        )
        self.calendario.pack(padx=10, pady=5)
        
        proximo_miercoles = self.obtener_proximo_miercoles(datetime.now())
        self.calendario.set_date(proximo_miercoles)
        
        # Info suggestion
        self.sugerencia_lbl = tk.Label(
            card,
            text=f"💡 Sugerencia: {proximo_miercoles.strftime('%d de %B de %Y')}",
            font=("Segoe UI", 9, "italic"),
            bg=self.controller.COLOR_CARTA,
            fg="#27AE60"
        )
        self.sugerencia_lbl.pack(pady=15)
        
        # Action Button
        self.btn_ejecutar = tk.Button(
            card,
            text="Generar Proyección Semanal",
            font=("Segoe UI", 12, "bold"),
            bg=self.controller.COLOR_SECUNDARIO,
            fg="white",
            relief=tk.FLAT,
            padx=50,
            pady=18,
            cursor="hand2",
            activebackground="#2980B9",
            activeforeground="white",
            bd=0,
            highlightthickness=0,
            command=self.iniciar_proceso
        )
        self.btn_ejecutar.pack(pady=(20, 0))
        
        def on_enter(e): self.btn_ejecutar.configure(bg="#2980B9", relief=tk.FLAT)
        def on_leave(e): self.btn_ejecutar.configure(bg=self.controller.COLOR_SECUNDARIO, relief=tk.FLAT)
        self.btn_ejecutar.bind("<Enter>", on_enter)
        self.btn_ejecutar.bind("<Leave>", on_leave)

    def obtener_proximo_miercoles(self, fecha):
        dias_hasta_miercoles = (2 - fecha.weekday()) % 7
        if dias_hasta_miercoles == 0:
            dias_hasta_miercoles = 7
        return fecha + timedelta(days=dias_hasta_miercoles)

    def iniciar_proceso(self):
        fecha = self.calendario.get_date()
        
        if not messagebox.askyesno(
            "Confirmar Ejecución",
            f"¿Desea iniciar la proyección para la semana del {fecha.strftime('%d/%m/%Y')}?\n\n"
            "Asegúrese de cerrar los archivos de Excel relacionados."
        ):
            return
            
        # Switch to progress view
        self.controller.show_view("PROGRESS")
        progreso = self.controller.current_view
        
        # Start thread
        def run():
            try:
                procesador = ProcesadorSemanal(
                    fecha_filtrado=fecha,
                    ventana_progreso=progreso,
                    rutas_config=self.controller.rutas
                )
                procesador.ejecutar_proceso()
                progreso.set_progress(100, "✅ Proceso completado exitosamente")
                progreso.log("Reporte generado correctamente", "OK")
            except Exception as e:
                progreso.log(f"Error fatal: {str(e)}", "ERROR")
                progreso.set_progress(0, "❌ Error en el proceso")
                messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
        
        threading.Thread(target=run, daemon=True).start()

class ProcesadorSemanal:
    """Clase principal para procesar archivos"""

    marcas_global = ['AMERICANINO', 'ESPRIT', 'CHEVIGNON']
    marcas_unifed = ['NAF NAF', 'RIFLE', 'AEO']
    
    def __init__(self, fecha_filtrado, ventana_progreso, rutas_config):
        self.fecha_filtrado = fecha_filtrado
        self.ventana_progreso = ventana_progreso

        self.log("Cargando motores de datos (Pandas/Excel)...", "INFO")
        global pd, win32com, pythoncom, openpyxl, Font, Alignment, Border, Side, dataframe_to_rows, load_workbook
        
        import pandas as pd
        import win32com.client as win32com
        import pythoncom
        import openpyxl
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        from openpyxl.utils.dataframe import dataframe_to_rows

        # Configurar rutas
        self.ruta_origen = rutas_config['origen']
        self.carpeta_proyecciones = rutas_config['proyecciones']
        self.ruta_destino_final = rutas_config['final']
        
        self.log("Configuración de rutas cargada", "OK")
        self.log(f"Origen: {self.ruta_origen.name}", "INFO")
        self.log(f"Proyecciones: {self.carpeta_proyecciones}", "INFO")
        self.log(f"Final: {self.ruta_destino_final.name}", "INFO")
    
    def log(self, mensaje, tipo="INFO"):
        self.ventana_progreso.log(mensaje, tipo)
        
    def set_progress(self, percent, status=None):
        self.ventana_progreso.set_progress(percent, status)

    def crear_estructura_carpetas(self, fecha):
        self.log("Creando estructura de carpetas", "INFO")
        año = fecha.year
        mes = fecha.strftime('%B').upper()
        carpeta_destino = self.carpeta_proyecciones / f"AÑO {año}" / mes
        carpeta_destino.mkdir(parents=True, exist_ok=True)
        return carpeta_destino
    
    def crear_nombre_archivo(self, fecha):
        dia = fecha.strftime('%d')
        mes = fecha.strftime('%B').upper()
        año = fecha.strftime('%Y')
        return f"{dia} {mes} {año}.xlsx"

    def crear_nombre_segunda_hoja(self, fecha):
        mes = fecha.strftime('%B').upper()
        dia = fecha.strftime('%d')
        return f"{mes} {dia}"
    
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
        
    def leer_datos_proceso_semanal(self, ruta_archivo):
        self.log("Leyendo datos del archivo", "INFO")
        try:
            pythoncom.CoInitialize()
            excel = win32com.DispatchEx("Excel.Application")
            wb = excel.Workbooks.Open(str(ruta_archivo.absolute()))
            ws = wb.Sheets(1)
            data = ws.UsedRange.Value
            if data:
                data_list = [list(row) if row is not None else [] for row in data]
                headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(data_list[0])]
                df = pd.DataFrame(data_list[1:], columns=headers)
                return df
            return None
        finally:
            if wb: wb.Close(SaveChanges=False)
            if excel: excel.Quit()
            pythoncom.CoUninitialize()
    
    def filtrar_por_fecha(self, df, fecha_filtrado):
        self.log(f"Filtrando por fecha de proyección: {fecha_filtrado}", "PROCESO")
        df.columns = [str(col).strip().upper() for col in df.columns]
        col_fecha = 'FECHA DE VENCIMIENTO' if 'FECHA DE VENCIMIENTO' in df.columns else 'FECHA DE PAGO' if 'FECHA DE PAGO' in df.columns else None
        
        if col_fecha:
            df[col_fecha] = pd.to_datetime(df[col_fecha].astype(str).replace(['None', 'nan', 'NaT', ''], pd.NA), errors='coerce')
            inicio = fecha_filtrado - timedelta(days=fecha_filtrado.weekday())
            fin = inicio + timedelta(days=6)
            df = df.dropna(subset=[col_fecha])
            df = df[(df[col_fecha].dt.date >= inicio) & (df[col_fecha].dt.date <= fin)]
            if 'ESTADO' in df.columns:
                df = df[df['ESTADO'].astype(str).str.upper().str.contains('PAGAR', na=False)]
            return df
        return pd.DataFrame()
    
    def preparar_datos_segunda_hoja(self, df):
        df_res = pd.DataFrame()
        cols = {'IMPORTADOR': 'IMPORTADOR', 'MARCA': 'MARCA', 'PROVEEDOR': 'PROVEEDOR', 'NRO. IMPO': 'NRO. IMPO', 'VALOR A PAGAR': 'VALOR A PAGAR', 'MONEDA': 'MONEDA', 'NOTA CRÉDITO': 'VALOR NOTA CRÉDITO'}
        for d, o in cols.items():
            df_res[d] = df[o] if o in df.columns else (0 if d == 'NOTA CRÉDITO' else '')
        
        def norm_marca(v):
            v = str(v).replace('COMODIN S.A.S - ', '').strip().upper()
            if 'NAF' in v: return 'NAF NAF'
            if 'ESPRIT' in v: return 'ESPRIT'
            if 'CHEVI' in v: return 'CHEVIGNON'
            if 'AMERICANINO' in v: return 'AMERICANINO'
            if 'RIFLE' in v: return 'RIFLE'
            if 'AEO' in v: return 'AEO'
            return v
        df_res['MARCA'] = df_res['MARCA'].apply(norm_marca)
        df_res['VALOR A PAGAR'] = pd.to_numeric(df_res['VALOR A PAGAR'], errors='coerce').fillna(0)
        return df_res
    
    def agrupar_y_calcular(self, df):
        df['VALOR A PAGAR'] = pd.to_numeric(df['VALOR A PAGAR'], errors='coerce').fillna(0)
        df['NOTA CRÉDITO'] = pd.to_numeric(df['NOTA CRÉDITO'], errors='coerce').fillna(0)
        df = df.sort_values(by=['IMPORTADOR', 'PROVEEDOR', 'MARCA']).reset_index(drop=True)
        
        filas = []
        for (imp, prov), gp in df.groupby(['IMPORTADOR', 'PROVEEDOR'], sort=False):
            f_ini = len(filas) + 2
            for _, row in gp.iterrows():
                d = row.to_dict(); d['_TIPO'] = 'DETALLE'; filas.append(d)
            if len(gp) > 1:
                f_fin = len(filas) + 1
                filas.append({'VALOR A PAGAR': f'=SUBTOTAL(9, E{f_ini}:E{f_fin})', 'MONEDA': gp['MONEDA'].iloc[0], '_TIPO': 'SUBTOTAL'})
                filas.append({'_TIPO': 'BLANCO'})
            else:
                if filas: filas[-1]['_TIPO'] = 'DETALLE_UNICO'
                filas.extend([{'_TIPO': 'BLANCO'}, {'_TIPO': 'BLANCO'}])
        return pd.DataFrame(filas)
    
    def guardar_proyeccion_com(self, ruta, df, nombre):
        self.log("Guardando proyección con tabla resumen...", "INFO")
        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = win32com.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(str(ruta.absolute()))
            ws = wb.Sheets.Add()
            ws.Name = nombre
            
            headers = ['IMPORTADOR', 'MARCA', 'PROVEEDOR', 'NRO. IMPO', 'VALOR A PAGAR', 'MONEDA', 'NOTA CRÉDITO']
            for i, h in enumerate(headers, 1):
                c = ws.Cells(1, i)
                c.Value = h
                c.Font.Bold = True
                c.Font.Color = 0xFFFFFF
                c.Interior.Color = 0x993366
                c.HorizontalAlignment = -4108  # xlCenter
                # Borde en encabezado
                c.Borders.LineStyle = 1
                
            f_act = 2
            for _, row in df.iterrows():
                t = row.get('_TIPO', 'DETALLE')

                if t == 'BLANCO':
                    # Fila vacía: no se escribe nada, no se aplican bordes
                    f_act += 1
                    continue

                if t == 'SUBTOTAL':
                    ws.Cells(f_act, 5).Formula = row['VALOR A PAGAR']
                    ws.Cells(f_act, 6).Value = row['MONEDA']
                    ws.Range(ws.Cells(f_act, 5), ws.Cells(f_act, 6)).Interior.Color = 0xCCFFCC
                    ws.Range(ws.Cells(f_act, 5), ws.Cells(f_act, 6)).Font.Bold = True
                    # Borde solo en celdas con contenido del subtotal
                    ws.Range(ws.Cells(f_act, 1), ws.Cells(f_act, 7)).Borders.LineStyle = 1
                else:
                    for i, col in enumerate(headers, 1):
                        val = row.get(col, '')
                        if col in ['VALOR A PAGAR', 'NOTA CRÉDITO']:
                            ws.Cells(f_act, i).Value = float(val) if val and not str(val).startswith('=') else val
                        else:
                            ws.Cells(f_act, i).Value = str(val)
                    if t == 'DETALLE_UNICO':
                        ws.Range(ws.Cells(f_act, 5), ws.Cells(f_act, 6)).Interior.Color = 0xCCFFCC
                        ws.Range(ws.Cells(f_act, 5), ws.Cells(f_act, 6)).Font.Bold = True
                    # Borde para DETALLE y DETALLE_UNICO
                    ws.Range(ws.Cells(f_act, 1), ws.Cells(f_act, 7)).Borders.LineStyle = 1

                ws.Cells(f_act, 5).NumberFormat = "$ #,##0.00"
                ws.Cells(f_act, 7).NumberFormat = "$ #,##0.00"
                f_act += 1

            # --- FILA TOTAL FINAL ---
            ultima_fila_datos = f_act - 1
            ws.Cells(f_act, 4).Value = "TOTAL"
            ws.Cells(f_act, 5).Formula = f'=SUBTOTAL(9, E2:E{ultima_fila_datos})'
            ws.Range(ws.Cells(f_act, 4), ws.Cells(f_act, 5)).Interior.Color = 0xCCFFCC
            ws.Range(ws.Cells(f_act, 4), ws.Cells(f_act, 5)).Font.Bold = True
            ws.Cells(f_act, 5).NumberFormat = "$ #,##0.00"
            ws.Range(ws.Cells(f_act, 4), ws.Cells(f_act, 5)).Borders.LineStyle = 1

            ws.Columns.AutoFit()

            # =========================================================
            # TABLA RESUMEN GLOBAL / UNIFIED / AEO  —  Columna J (col 10)
            # =========================================================
            COL_LABEL = 10   # J
            COL_VALOR = 11   # K
            FILA_INI  = 2

            rango_criterio = f'B2:B{ultima_fila_datos}'
            rango_suma     = f'E2:E{ultima_fila_datos}'

            def formula_sumif(marcas):
                partes = [f'SUMIF({rango_criterio},"{m}",{rango_suma})' for m in marcas]
                return '=' + '+'.join(partes)

            # Datos de cada fila: (etiqueta, fórmula, color_fondo_hex_bgr)
            filas_resumen = [
                ("Global",   formula_sumif(self.marcas_global),                            0xEED7BD),
                ("Unified",  formula_sumif([m for m in self.marcas_unifed if m != 'AEO']), 0x99E6FF),
                ("AEO",      f'=SUMIF({rango_criterio},"AEO",{rango_suma})',               0xCCFFCC),
            ]

            for i, (etiqueta, formula, color) in enumerate(filas_resumen):
                fila = FILA_INI + i
                c_label = ws.Cells(fila, COL_LABEL)
                c_label.Value = etiqueta
                c_label.Interior.Color = color
                c_label.Font.Bold = True
                c_valor = ws.Cells(fila, COL_VALOR)
                c_valor.Formula = formula
                c_valor.Interior.Color = color
                c_valor.NumberFormat = "$ #,##0.00"

            # Gran Total resumen
            fila_total_resumen = FILA_INI + len(filas_resumen)
            c_gt_label = ws.Cells(fila_total_resumen, COL_LABEL)
            c_gt_label.Value = "Total"
            c_gt_label.Font.Bold = True

            c_gt_valor = ws.Cells(fila_total_resumen, COL_VALOR)
            c_gt_valor.Formula = f'=SUM(K{FILA_INI}:K{fila_total_resumen - 1})'
            c_gt_valor.Font.Bold = True
            c_gt_valor.NumberFormat = "$ #,##0.00"
            # Borde superior para separar del resto
            c_gt_valor.Borders(8).LineStyle = 1   # xlEdgeTop
            c_gt_valor.Borders(8).Weight    = 3   # xlMedium
            c_gt_label.Borders(8).LineStyle = 1
            c_gt_label.Borders(8).Weight    = 3

            # Bordes al bloque completo J2:K(total)
            rango_tabla = ws.Range(
                ws.Cells(FILA_INI, COL_LABEL),
                ws.Cells(fila_total_resumen, COL_VALOR)
            )
            rango_tabla.Borders.LineStyle = 1
            rango_tabla.Borders.Weight    = 2

            # Autoajustar columnas J y K
            ws.Columns(COL_LABEL).AutoFit()
            ws.Columns(COL_VALOR).AutoFit()

            self.log("Tabla resumen Global/Unified/AEO escrita en columna J", "OK")
            # =========================================================

            wb.Save()
            self.log("Proyección guardada exitosamente", "OK")

        except Exception as e:
            self.log(f"Error al guardar proyección: {str(e)}", "ERROR")
            raise
        finally:
            if wb: wb.Close()
            if excel: excel.Quit()
            pythoncom.CoUninitialize()

    def preparar_df_final(self, df_detalle):
        """Prepara DataFrame para archivo final con columnas del formato destino"""
        self.log("Preparando datos para archivo final", "INFO")

        df_final = pd.DataFrame()
        fecha_proyeccion = self.fecha_filtrado

        df_final['IMPORTADOR']          = df_detalle['IMPORTADOR']
        df_final['MARCA']               = df_detalle['MARCA']
        df_final['FECHA DE PAGO']       = fecha_proyeccion.strftime('%m/%d/%Y')
        df_final['DIA']                 = fecha_proyeccion.day
        df_final['MES']                 = fecha_proyeccion.month
        df_final['AÑO']                 = fecha_proyeccion.year
        df_final['PROVEEDOR']           = df_detalle['PROVEEDOR']
        df_final['# IMPORTACION']       = df_detalle['NRO. IMPO']
        df_final['VALOR MONEDA ORIGEN'] = df_detalle['VALOR A PAGAR']
        df_final['MONEDA']              = df_detalle['MONEDA']

        def calc_valor_usd(row):
            if str(row['MONEDA']).upper() == 'USD':
                return row['VALOR A PAGAR']
            return ''

        def calc_factor(row):
            if str(row['MONEDA']).upper() == 'USD':
                return 1
            return ''

        df_final['VALOR USD']                                   = df_detalle.apply(calc_valor_usd, axis=1)
        df_final['FACTOR DE CONVERSION']                        = df_detalle.apply(calc_factor, axis=1)
        df_final['DESCUENTO PRONTO PAGO']                       = 0
        df_final['FORMA DE PAGO']                               = ''
        df_final['TIPO DE PAGO']                                = 'CUENTA COMPENSACION'
        df_final['FECHA DE APERTURA CREDITO -UTILIZACION LC']   = 'N/A'
        df_final['FECHA DE VENCIMIENTO']                        = 'N/A'
        df_final['# CREDITO']                                   = 'N/A'
        df_final['# DEUDA EXTERNA']                             = 'N/A'
        df_final['NOTA CREDITO']                                = 0.00
        df_final['OBSERVACIONES']                               = ''

        self.log(f"DataFrame final preparado: {len(df_final)} registros", "OK")
        return df_final

    def anexar_archivo_final_com(self, df_detalle):
        """Anexa o reemplaza datos en el archivo final usando COM con validación de duplicados"""
        self.log("Actualizando archivo final (Verificando duplicados)...", "INFO")

        pythoncom.CoInitialize()
        excel = None
        wb = None

        try:
            excel = win32com.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            archivo_path = str(self.ruta_destino_final.absolute())
            wb = excel.Workbooks.Open(archivo_path)

            # Búsqueda flexible de la hoja destino
            ws = None
            for sheet in wb.Sheets:
                nombre_u = sheet.Name.upper()
                if 'PAGOS' in nombre_u and 'IMPOR' in nombre_u:
                    ws = sheet
                    break

            if not ws:
                ws = wb.Sheets(1)
                self.log(f"No se halló hoja 'Pagos importación', usando: '{ws.Name}'", "WARN")
            else:
                self.log(f"Escribiendo en hoja: '{ws.Name}'", "OK")

            used_range   = ws.UsedRange
            filas_totales = used_range.Rows.Count

            start_row       = filas_totales + 1
            datos_a_escribir = []
            escribir_todo   = False
            headers_finales = []

            def normalizar_clave(val):
                if val is None: return ""
                s = str(val).strip()
                try:
                    if isinstance(val, datetime):
                        return val.strftime('%m/%d/%Y')
                    if '/' in s or '-' in s:
                        ts = pd.to_datetime(s, errors='coerce')
                        if not pd.isna(ts):
                            return ts.strftime('%m/%d/%Y')
                except:
                    pass
                return s

            def ordenar_por_importador(df):
                cols_orden = [c for c in ['IMPORTADOR', 'PROVEEDOR', 'MARCA', '# IMPORTACION'] if c in df.columns]
                if not cols_orden:
                    return df.fillna("")
                return df.fillna("").sort_values(by=cols_orden, kind='mergesort', ignore_index=True)

            if filas_totales > 1:
                self.log("Leyendo registros existentes...", "INFO")
                raw_data = list(used_range.Value)
                headers  = [str(h).strip().upper() if h is not None else f"COL_{i}"
                            for i, h in enumerate(raw_data[0])]

                cols_clave    = ['FECHA DE PAGO', 'PROVEEDOR', 'IMPORTADOR', 'MARCA', '# IMPORTACION']
                indices_clave = {col: headers.index(col) for col in cols_clave if col in headers}

                if len(indices_clave) == 5:
                    data_rows = [list(row) for row in raw_data[1:]]
                    self.log(f"Registros previos: {len(data_rows)}", "INFO")

                    claves_nuevas = set()
                    for _, row in df_detalle.iterrows():
                        key = (
                            normalizar_clave(row.get('FECHA DE PAGO')),
                            normalizar_clave(row.get('PROVEEDOR')),
                            normalizar_clave(row.get('IMPORTADOR')),
                            normalizar_clave(row.get('MARCA')),
                            normalizar_clave(row.get('# IMPORTACION'))
                        )
                        claves_nuevas.add(key)

                    idx_fecha = indices_clave['FECHA DE PAGO']
                    idx_prov  = indices_clave['PROVEEDOR']
                    idx_imp   = indices_clave['IMPORTADOR']
                    idx_marca = indices_clave['MARCA']
                    idx_nro   = indices_clave['# IMPORTACION']
                    max_idx   = max(idx_fecha, idx_prov, idx_imp, idx_marca, idx_nro)

                    rows_a_conservar     = []
                    duplicados_encontrados = 0

                    for row in data_rows:
                        try:
                            if len(row) <= max_idx:
                                rows_a_conservar.append(row)
                                continue
                            key_existente = (
                                normalizar_clave(row[idx_fecha]),
                                normalizar_clave(row[idx_prov]),
                                normalizar_clave(row[idx_imp]),
                                normalizar_clave(row[idx_marca]),
                                normalizar_clave(row[idx_nro])
                            )
                            if key_existente in claves_nuevas:
                                duplicados_encontrados += 1
                            else:
                                rows_a_conservar.append(row)
                        except Exception:
                            rows_a_conservar.append(row)

                    if duplicados_encontrados > 0:
                        self.log(f"Reemplazando {duplicados_encontrados} registros duplicados.", "OK")

                        df_conservado       = pd.DataFrame(rows_a_conservar, columns=headers, dtype=object)
                        df_detalle_obj      = df_detalle.astype(object)
                        df_final_combinado  = pd.concat([df_conservado, df_detalle_obj], ignore_index=True)

                        cols_finales = list(headers)
                        for col in df_final_combinado.columns:
                            if col not in cols_finales:
                                cols_finales.append(col)

                        df_final_combinado  = df_final_combinado.reindex(columns=cols_finales)
                        df_final_ordenado   = ordenar_por_importador(df_final_combinado)

                        datos_a_escribir = df_final_ordenado.values.tolist()
                        headers_finales  = cols_finales
                        escribir_todo    = True
                        start_row        = 2
                    else:
                        self.log("Sin duplicados. Agregando al final.", "INFO")
                        df_tmp = ordenar_por_importador(df_detalle)
                        datos_a_escribir = df_tmp.values.tolist()
                else:
                    self.log(f"Faltan columnas clave. Agregando al final.", "WARN")
                    df_tmp = ordenar_por_importador(df_detalle)
                    datos_a_escribir = df_tmp.values.tolist()
            else:
                self.log("Archivo destino vacío. Escribiendo desde fila 2.", "INFO")
                df_tmp = ordenar_por_importador(df_detalle)
                datos_a_escribir = df_tmp.values.tolist()
                start_row = 2

            if not datos_a_escribir:
                self.log("No hay datos para escribir.", "WARN")
                return

            if escribir_todo:
                ws.Range(
                    ws.Cells(2, 1),
                    ws.Cells(filas_totales + 1000, len(headers_finales) + 10)
                ).ClearContents()
                if len(headers_finales) > len(headers):
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers_finales))).Value = headers_finales

            self.log(f"Escribiendo {len(datos_a_escribir)} registros desde fila {start_row}", "INFO")

            num_filas = len(datos_a_escribir)
            num_cols  = len(datos_a_escribir[0])
            destino   = ws.Range(
                ws.Cells(start_row, 1),
                ws.Cells(start_row + num_filas - 1, num_cols)
            )
            destino.Value = datos_a_escribir

            wb.Save()
            self.log(f"Datos guardados en {self.ruta_destino_final.name}", "OK")

        except Exception as e:
            self.log(f"Error crítico al anexar: {str(e)}", "ERROR")
            raise

        finally:
            if wb:
                try: wb.Close(SaveChanges=True)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()

    def agregar_a_archivo_final(self, df_detalle):
        """Orquesta la preparación y escritura en el archivo final"""
        try:
            df_final = self.preparar_df_final(df_detalle)
            self.anexar_archivo_final_com(df_final)
        except Exception as e:
            self.log(f"Error en archivo final: {str(e)}", "ERROR")
            raise

    def ejecutar_proceso(self):
        self.set_progress(10, "Iniciando...")
        try:
            carpeta = self.crear_estructura_carpetas(self.fecha_filtrado)
            ruta = carpeta / self.crear_nombre_archivo(self.fecha_filtrado)
            self.copiar_archivo_base(ruta)
            self.set_progress(40, "Leyendo datos...")
            df = self.leer_datos_proceso_semanal(ruta)
            self.set_progress(55, "Filtrando registros...")
            df_f = self.filtrar_por_fecha(df, self.fecha_filtrado)
            self.set_progress(70, "Preparando proyección...")
            df_s = self.preparar_datos_segunda_hoja(df_f)
            df_a = self.agrupar_y_calcular(df_s)
            self.set_progress(85, "Guardando archivo de proyección...")
            self.guardar_proyeccion_com(ruta, df_a, self.crear_nombre_segunda_hoja(self.fecha_filtrado))
            self.set_progress(93, "Actualizando archivo final...")
            self.agregar_a_archivo_final(df_s)
            self.set_progress(100, "Completado")
            return True
        except Exception as e:
            raise e