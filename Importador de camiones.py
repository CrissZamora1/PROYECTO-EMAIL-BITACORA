import os
import re
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog, ttk, scrolledtext

try:
    from tkcalendar import Calendar
    TKCALENDAR_DISPONIBLE = True
except ImportError:
    TKCALENDAR_DISPONIBLE = False

try:
    import win32clipboard
    import win32com.client
    PYWIN32_DISPONIBLE = True
except ImportError:
    PYWIN32_DISPONIBLE = False

try:
    from bs4 import BeautifulSoup
    BS4_DISPONIBLE = True
except ImportError:
    BS4_DISPONIBLE = False

class AdaptableInteligenteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consolidador - Pegado de Tabla [DEBUG v6 - ventana HTML]")
        self.root.geometry("1150x750")
        self.root.config(bg="#f4f6f9")

        # --- TÍTULO ---
        titulo = tk.Label(root, text="Consolidador de Tabla (pega y exporta)", font=("Arial", 14, "bold"), bg="#1F4E78", fg="white", pady=10)
        titulo.pack(fill=tk.X)

        # --- 1. DATOS GENERALES ---
        frame_datos = tk.LabelFrame(root, text=" 1. Datos Generales (se aplican al PRÓXIMO bloque que agregues) ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_datos.pack(fill=tk.X, padx=15, pady=5)

        tk.Label(frame_datos, text="FLOTA:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", padx=5)
        self.entry_flota = tk.Entry(frame_datos, font=("Arial", 10), width=15)
        self.entry_flota.grid(row=0, column=1, padx=5, sticky="w")

        tk.Label(frame_datos, text="FECHA:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky="w", padx=5)
        self.entry_fecha = tk.Entry(frame_datos, font=("Arial", 10), width=12)
        self.entry_fecha.grid(row=0, column=3, padx=5, sticky="w")
        self.entry_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))

        if TKCALENDAR_DISPONIBLE:
            tk.Button(frame_datos, text="📅", bg="#4A90E2", fg="white", command=self.abrir_calendario).grid(row=0, column=4, padx=2)

        tk.Label(frame_datos, text="HORA:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=5, sticky="w", padx=5)
        self.entry_hora = tk.Entry(frame_datos, font=("Arial", 10), width=10)
        self.entry_hora.grid(row=0, column=6, padx=5, sticky="w")
        self.entry_hora.insert(0, "7:15")
        tk.Label(frame_datos, text="(ej: 2:30 -> se guarda 2:30:00)", bg="#f4f6f9", fg="#555555", font=("Arial", 8, "italic")).grid(row=0, column=7, sticky="w", padx=5)

        # --- 2. CONFIGURACIÓN DE COLUMNAS ---
        frame_cols = tk.LabelFrame(root, text=" 2. Columnas Oficiales (en el mismo orden en que vienen en la tabla que pegas) ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_cols.pack(fill=tk.X, padx=15, pady=5)

        self.entry_columnas = tk.Entry(frame_cols, font=("Arial", 10))
        self.entry_columnas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        self.entry_columnas.insert(0, "NO., NOMBRE, LICENCIA, PLACA, TELEFONO, CHASIS, TCCHASIS, GENSET, CONTENEDOR, UBICACIÓN, STATUS, DESTINO")

        tk.Button(frame_cols, text="🔄 Actualizar Tabla", font=("Arial", 9, "bold"), bg="#4A90E2", fg="white", command=self.actualizar_columnas_tabla).pack(side=tk.RIGHT, padx=5)

        # --- 2B. CONFIGURACIÓN DE COLUMNAS PARA REPORTES "DOLLYS" ---
        # Estos correos traen columnas extra (DOLLY, CHASIS2, TC-CHASIS2,
        # GENSET2, CONTENEDOR2) además de las normales, así que se importan
        # a una pestaña/tabla aparte con su propio set de columnas.
        frame_cols_dollys = tk.LabelFrame(root, text=" 2B. Columnas para reportes DOLLYS (se detectan porque su encabezado trae la columna DOLLY) ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_cols_dollys.pack(fill=tk.X, padx=15, pady=5)

        self.entry_columnas_dollys = tk.Entry(frame_cols_dollys, font=("Arial", 10))
        self.entry_columnas_dollys.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        self.entry_columnas_dollys.insert(0, "NO., NOMBRE, LICENCIA, PLACA, TELEFONO, CHASIS, TCCHASIS, GENSET, CONTENEDOR, DOLLY, CHASIS2, TCCHASIS2, GENSET2, CONTENEDOR2, UBICACIÓN, STATUS, DESTINO")

        tk.Button(frame_cols_dollys, text="🔄 Actualizar Tabla Dollys", font=("Arial", 9, "bold"), bg="#4A90E2", fg="white", command=self.actualizar_columnas_tabla_dollys).pack(side=tk.RIGHT, padx=5)

        # --- 3. ÁREA DE PEGADO DE LA TABLA ---
        frame_pegado = tk.LabelFrame(root, text=" 3. Pega aquí el texto vertical del correo (no importa si el espaciado es irregular) ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_pegado.pack(fill=tk.BOTH, padx=15, pady=5)

        self.text_area = scrolledtext.ScrolledText(frame_pegado, wrap=tk.WORD, font=("Consolas", 9), height=6)
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        frame_opciones = tk.Frame(frame_pegado, bg="#f4f6f9")
        frame_opciones.pack(fill=tk.X, pady=5)

        tk.Button(frame_opciones, text="📋 Pegar Tabla con Formato (Outlook/Excel)", font=("Arial", 10, "bold"), bg="#2E7D32", fg="white", padx=15, pady=3, command=self.pegar_html_portapapeles).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_opciones, text="⚡ Procesar Texto del Cuadro", font=("Arial", 10, "bold"), bg="#D9822B", fg="white", padx=15, pady=3, command=self.procesar_tabla_pegada).pack(side=tk.RIGHT, padx=5)

        # --- 3B. IMPORTAR DESDE OUTLOOK POR RANGO DE FECHAS ---
        frame_outlook = tk.LabelFrame(root, text=" 3B. Importar varios correos de Outlook por rango de fechas ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_outlook.pack(fill=tk.X, padx=15, pady=5)

        tk.Label(frame_outlook, text="Carpeta:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", padx=5)
        self.entry_carpeta_outlook = tk.Entry(frame_outlook, font=("Arial", 10), width=16)
        self.entry_carpeta_outlook.grid(row=0, column=1, padx=5, sticky="w")
        self.entry_carpeta_outlook.insert(0, "Ubicaciones")

        tk.Label(frame_outlook, text="Desde:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky="w", padx=5)
        self.entry_fecha_desde = tk.Entry(frame_outlook, font=("Arial", 10), width=12)
        self.entry_fecha_desde.grid(row=0, column=3, padx=5, sticky="w")
        self.entry_fecha_desde.insert(0, datetime.now().strftime("%d/%m/%Y"))
        if TKCALENDAR_DISPONIBLE:
            tk.Button(frame_outlook, text="📅", bg="#4A90E2", fg="white", command=lambda: self.abrir_calendario_generico(self.entry_fecha_desde)).grid(row=0, column=4, padx=2)

        tk.Label(frame_outlook, text="Hasta:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=5, sticky="w", padx=5)
        self.entry_fecha_hasta = tk.Entry(frame_outlook, font=("Arial", 10), width=12)
        self.entry_fecha_hasta.grid(row=0, column=6, padx=5, sticky="w")
        self.entry_fecha_hasta.insert(0, datetime.now().strftime("%d/%m/%Y"))
        if TKCALENDAR_DISPONIBLE:
            tk.Button(frame_outlook, text="📅", bg="#4A90E2", fg="white", command=lambda: self.abrir_calendario_generico(self.entry_fecha_hasta)).grid(row=0, column=7, padx=2)

        tk.Button(frame_outlook, text="📥 Buscar y Cargar Correos", font=("Arial", 10, "bold"), bg="#6A1B9A", fg="white", padx=15, pady=3, command=self.importar_correos_outlook).grid(row=0, column=8, padx=15)

        tk.Label(frame_outlook, text="Mapeo de flota (palabra clave del bloque \"MONITOREO ...\" = código):", bg="#f4f6f9", font=("Arial", 8, "italic")).grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=(4, 0))
        self.entry_mapeo_flota = tk.Entry(frame_outlook, font=("Arial", 9), width=45)
        self.entry_mapeo_flota.grid(row=1, column=4, columnspan=5, sticky="w", padx=5, pady=(4, 0))
        self.entry_mapeo_flota.insert(0, "SALVADOR=SV, GUATEMALA=GT, ROSMAR=RR")

        # NOTA: antes había una casilla para ignorar correos "RE:/FW:/..."
        # por el asunto, pero se quitó — un correo de respuesta puede traer
        # datos NUEVOS legítimos (alguien contesta el hilo en vez de escribir
        # uno nuevo), y filtrarlos a ciegas por el asunto los perdía. Ahora
        # los duplicados se detectan por CONTENIDO (ver _insertar_filas):
        # solo se omite una fila si su PLACA+CHASIS+FECHA+HORA+FLOTA ya se
        # importó antes en esta misma corrida, sin importar el asunto.
        self.var_excluir_respuestas = tk.BooleanVar(value=False)

        # --- 4. TABLA VISUAL (dos pestañas: Camiones normales y Dollys) ---
        frame_tabla = tk.LabelFrame(root, text=" 4. Resultado (Doble clic para editar celdas si lo requieres) ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_tabla.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        self.notebook_resultado = ttk.Notebook(frame_tabla)
        self.notebook_resultado.pack(fill=tk.BOTH, expand=True)

        tab_camiones = tk.Frame(self.notebook_resultado, bg="#f4f6f9")
        tab_dollys = tk.Frame(self.notebook_resultado, bg="#f4f6f9")
        self.notebook_resultado.add(tab_camiones, text="Camiones")
        self.notebook_resultado.add(tab_dollys, text="Dollys")

        self.tree = ttk.Treeview(tab_camiones, show="headings", height=8)
        scrollbar_y = ttk.Scrollbar(tab_camiones, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(tab_camiones, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X, before=self.tree)
        self.tree.bind("<Double-1>", self.editar_celda)

        self.tree_dollys = ttk.Treeview(tab_dollys, show="headings", height=8)
        scrollbar_y_d = ttk.Scrollbar(tab_dollys, orient=tk.VERTICAL, command=self.tree_dollys.yview)
        scrollbar_x_d = ttk.Scrollbar(tab_dollys, orient=tk.HORIZONTAL, command=self.tree_dollys.xview)
        self.tree_dollys.configure(yscrollcommand=scrollbar_y_d.set, xscrollcommand=scrollbar_x_d.set)
        self.tree_dollys.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y_d.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x_d.pack(side=tk.BOTTOM, fill=tk.X, before=self.tree_dollys)
        self.tree_dollys.bind("<Double-1>", self.editar_celda)

        self.columnas_actuales = []
        self.columnas_actuales_dollys = []
        self.actualizar_columnas_tabla()
        self.actualizar_columnas_tabla_dollys()

        # --- PANEL INFERIOR ---
        frame_inferior = tk.Frame(root, bg="#f4f6f9", pady=10, padx=10)
        frame_inferior.pack(fill=tk.X)

        tk.Button(frame_inferior, text="🗑️ Limpiar Tabla (pestaña activa)", font=("Arial", 9, "bold"), bg="#C62828", fg="white", padx=10, command=self.limpiar_tabla_activa).pack(side=tk.LEFT)
        tk.Button(frame_inferior, text="💾 Generar y Guardar Excel", font=("Arial", 11, "bold"), bg="#1F4E78", fg="white", padx=20, pady=8, command=self.guardar_excel).pack(side=tk.RIGHT)

    def abrir_calendario(self):
        self.abrir_calendario_generico(self.entry_fecha)

    def abrir_calendario_generico(self, entry_destino):
        top = tk.Toplevel(self.root)
        top.title("Seleccionar Fecha")
        top.geometry("320x280")
        top.grab_set()
        cal = Calendar(top, selectmode='day', date_pattern='dd/mm/yyyy')
        cal.pack(pady=15, fill=tk.BOTH, expand=True)
        tk.Button(top, text="Aceptar", font=("Arial", 10, "bold"), bg="#2E7D32", fg="white", command=lambda: [entry_destino.delete(0, tk.END), entry_destino.insert(0, cal.get_date()), top.destroy()]).pack(pady=5)

    def actualizar_columnas_tabla(self):
        texto_cols = self.entry_columnas.get()
        self.columnas_actuales = [c.strip() for c in texto_cols.split(",") if c.strip()]
        # La tabla visible incluye, además de las columnas oficiales, FLOTA/FECHA/HORA
        # al final — así cada bloque que agregues queda identificado.
        self.columnas_tabla = self.columnas_actuales + ["FLOTA", "FECHA", "HORA"]
        self.tree["columns"] = self.columnas_tabla
        for col in self.columnas_tabla:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="w")
        self.limpiar_tabla(self.tree)

    def actualizar_columnas_tabla_dollys(self):
        texto_cols = self.entry_columnas_dollys.get()
        self.columnas_actuales_dollys = [c.strip() for c in texto_cols.split(",") if c.strip()]
        # Los reportes de Dollys no manejan el concepto de FLOTA, así que
        # esa tabla NO lleva esa columna (a diferencia de Camiones).
        self.columnas_tabla_dollys = self.columnas_actuales_dollys + ["FECHA", "HORA"]
        self.tree_dollys["columns"] = self.columnas_tabla_dollys
        for col in self.columnas_tabla_dollys:
            self.tree_dollys.heading(col, text=col)
            self.tree_dollys.column(col, width=100, anchor="w")
        self.limpiar_tabla(self.tree_dollys)

    def limpiar_tabla(self, tree):
        for item in tree.get_children():
            tree.delete(item)

    def limpiar_tabla_activa(self):
        """Limpia la tabla de la pestaña actualmente visible (Camiones o Dollys)."""
        pestana_actual = self.notebook_resultado.index(self.notebook_resultado.select())
        tree = self.tree if pestana_actual == 0 else self.tree_dollys
        self.limpiar_tabla(tree)

    def editar_celda(self, event):
        tree = event.widget
        columnas = self.columnas_tabla if tree is self.tree else self.columnas_tabla_dollys
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        col_idx = int(col.replace('#', '')) - 1
        col_nombre = columnas[col_idx]
        val_actual = tree.item(item, "values")[col_idx]

        top = tk.Toplevel(self.root)
        top.title(f"Editar {col_nombre}")
        top.geometry("350x130")
        top.grab_set()

        tk.Label(top, text=f"Editar [{col_nombre}]:", font=("Arial", 9, "bold")).pack(pady=10)
        entry = tk.Entry(top, font=("Arial", 11), width=35)
        entry.pack(pady=5)
        entry.insert(0, val_actual)
        entry.focus()

        def guardar():
            valores = list(tree.item(item, "values"))
            valores[col_idx] = entry.get()
            tree.item(item, values=valores)
            top.destroy()

        tk.Button(top, text="Guardar", font=("Arial", 9, "bold"), bg="#2E7D32", fg="white", command=guardar).pack(pady=5)

    def _datos_bloque_actual(self):
        """Lee FLOTA/FECHA/HORA de los campos superiores en el momento en que
        se agrega un bloque, valida la fecha, y devuelve (flota, fecha, hora)
        o None si la fecha no es válida (y ya mostró el error)."""
        flota = self.entry_flota.get().strip()
        fecha_str = self.entry_fecha.get().strip()
        try:
            fecha = datetime.strptime(fecha_str, "%d/%m/%Y").strftime("%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido en '2. Datos Generales'. Usa dd/mm/yyyy.")
            return None
        hora = self._formatear_hora(self.entry_hora.get())
        return flota, fecha, hora

    def _agregar_filas_al_arbol(self, filas_datos, tree=None):
        """Agrega (sin borrar lo existente) una lista de filas de datos al
        árbol indicado (self.tree por defecto), etiquetadas con la
        FLOTA/FECHA/HORA actuales de los campos de arriba."""
        if tree is None:
            tree = self.tree
        datos_bloque = self._datos_bloque_actual()
        if datos_bloque is None:
            return 0
        flota, fecha, hora = datos_bloque
        return self._insertar_filas(tree, filas_datos, flota, fecha, hora)

    def _insertar_filas(self, tree, filas_datos, flota, fecha, hora, columnas_destino=None, claves_vistas=None):
        """Inserta filas de datos en el árbol indicado, agregando
        FECHA/HORA (y FLOTA, salvo para la tabla de Dollys, que no maneja
        ese concepto). Si se pasan columnas_destino (la lista de columnas
        de esa tabla) y claves_vistas (un set mutable), omite silenciosamente
        cualquier fila cuya combinación FECHA+HORA+PLACA+CHASIS (más FLOTA
        cuando aplica) ya se haya visto antes en ese set — así un correo
        duplicado (por ejemplo una respuesta/reenvío que repite el mismo
        reporte ya enviado) no se agrega dos veces, sin depender de adivinar
        por el asunto del correo si es o no una respuesta."""
        es_dollys = (tree is self.tree_dollys)
        extras = [fecha, hora] if es_dollys else [flota, fecha, hora]

        idx_placa = idx_chasis = None
        if columnas_destino is not None:
            idx_placa = self._indice_columna_en(columnas_destino, "PLACA")
            idx_chasis = self._indice_columna_en(columnas_destino, "CHASIS")

        insertadas = 0
        for fila in filas_datos:
            if claves_vistas is not None and idx_placa is not None:
                placa_val = self._normalizar(fila[idx_placa]) if idx_placa < len(fila) else ""
                chasis_val = self._normalizar(fila[idx_chasis]) if (idx_chasis is not None and idx_chasis < len(fila)) else ""
                if es_dollys:
                    clave = (self._normalizar(fecha), self._normalizar(hora), placa_val, chasis_val)
                else:
                    clave = (self._normalizar(flota), self._normalizar(fecha), self._normalizar(hora), placa_val, chasis_val)
                if clave in claves_vistas:
                    continue  # Duplicado exacto ya importado antes; se omite
                claves_vistas.add(clave)
            tree.insert("", tk.END, values=list(fila) + extras)
            insertadas += 1
        return insertadas

    def _parsear_mapeo_flota(self):
        """Convierte el texto 'SALVADOR=SV, GUATEMALA=GT, ROSMAR=RR' en una
        lista de tuplas (palabra_clave_normalizada, codigo), en el orden en
        que el usuario las escribió (importa el orden por si una palabra
        clave está contenida dentro de otra)."""
        texto = self.entry_mapeo_flota.get().strip()
        pares = []
        for parte in texto.split(","):
            if "=" not in parte:
                continue
            clave, codigo = parte.split("=", 1)
            clave = self._normalizar(clave)
            codigo = codigo.strip()
            if clave and codigo:
                pares.append((clave, codigo))
        return pares

    def _detectar_flota_en_texto(self, texto_monitoreo, mapeo):
        """Busca, en el texto del bloque 'MONITOREO ...', alguna de las
        palabras clave configuradas y devuelve su código. Si no reconoce
        ninguna, devuelve el texto crudo prefijado con 'REVISAR:' para que
        quede visible en la tabla y se pueda corregir a mano."""
        texto_norm = self._normalizar(texto_monitoreo)
        for clave, codigo in mapeo:
            if clave in texto_norm:
                return codigo
        crudo = texto_monitoreo.strip()
        return f"REVISAR: {crudo}" if crudo else "REVISAR: (sin identificar)"

    def _extraer_texto_monitoreo(self, soup):
        """Busca en el correo (ya parseado con BeautifulSoup) la línea que
        contiene la palabra 'MONITOREO' y devuelve lo que sigue a esa
        palabra (por ejemplo 'ROSMAR GROUP', 'EL SALVADOR', 'CHQ GUATEMALA').
        Devuelve '' si no encuentra esa línea."""
        texto_completo = soup.get_text("\n")
        for linea in texto_completo.split("\n"):
            linea = linea.strip()
            m = re.search(r'MONITOREO\s+(.+)', linea, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return ""

    def _extraer_hora_de_asunto(self, asunto):
        """Busca un patrón HH:MM al final del asunto del correo (que es
        donde siempre viene la hora, según el formato de estos correos)."""
        m = re.search(r'(\d{1,2}:\d{2})\s*$', asunto.strip())
        return self._formatear_hora(m.group(1)) if m else ""

    def _es_asunto_respuesta_o_reenvio(self, asunto):
        """Detecta si el asunto corresponde a una respuesta o reenvío
        (RE:, FW:, FWD:, RV:, y sus variantes con [EXTERNAL MAIL] u otros
        prefijos de seguridad de por medio), para poder ignorarlos y no
        duplicar los camiones del correo original al que responden/citan."""
        asunto_norm = asunto.strip().upper()
        # Quita etiquetas de seguridad tipo "[EXTERNAL MAIL]" antes de
        # revisar el prefijo real (RE:/FW:/etc.).
        asunto_norm = re.sub(r'^\[[^\]]*\]\s*', '', asunto_norm)
        return bool(re.match(r'^(RE|FW|FWD|RV)\s*:', asunto_norm))

    # ---- Extracción de tablas por REGEX (igual que el Office Script que ya
    # funciona con estos correos) en vez de un parser de DOM. El HTML real
    # de estos correos viene mal formado/anidado de una forma que
    # BeautifulSoup interpreta mal (pierde la tabla real de camiones), por
    # eso se abandona el parser y se replica exactamente la lógica probada:
    # regex sobre <table>, luego <tr>, luego <td|th>. ----

    def _decodificar_entidades_html(self, s):
        """Reemplaza las entidades HTML más comunes por su carácter real."""
        return (s.replace('&nbsp;', ' ').replace('&amp;', '&')
                 .replace('&lt;', '<').replace('&gt;', '>')
                 .replace('&#39;', "'").replace('&quot;', '"'))

    def _texto_celda_regex(self, html_celda):
        """Quita las etiquetas HTML de una celda <td>/<th> y decodifica
        entidades, dejando el texto plano tal como lo vería el usuario."""
        t = re.sub(r'<[^>]*>', ' ', html_celda)
        t = self._decodificar_entidades_html(t)
        return re.sub(r'\s+', ' ', t).strip()

    def _tablas_html_regex(self, html):
        """Extrae TODAS las tablas <table>...</table> del correo usando
        regex (no un parser de DOM), devolviendo una lista de tablas, cada
        una como lista de filas, y cada fila como lista de textos de celda."""
        tablas_html = re.findall(r'<table[\s\S]*?</table>', html, re.IGNORECASE)
        tablas = []
        for tabla_html in tablas_html:
            filas_html = re.findall(r'<tr[\s\S]*?</tr>', tabla_html, re.IGNORECASE)
            filas = []
            for fila_html in filas_html:
                celdas_html = re.findall(r'<t[dh][\s\S]*?</t[dh]>', fila_html, re.IGNORECASE)
                if not celdas_html:
                    continue
                filas.append([self._texto_celda_regex(c) for c in celdas_html])
            tablas.append(filas)
        return tablas

    def _fila_es_encabezado_camiones(self, fila):
        """Una fila es el encabezado real de la tabla de camiones si alguna
        de sus celdas normaliza EXACTAMENTE a 'PLACA' y otra a 'CHASIS'
        (igual que el Office Script), sin exigir que sea la primera fila
        de la tabla — el encabezado puede venir después de filas de título."""
        fila_norm = [self._normalizar(c) for c in fila]
        return "PLACA" in fila_norm and "CHASIS" in fila_norm

    def _es_encabezado_dolly(self, fila_encabezado):
        """Detecta si el encabezado ya identificado como 'de camiones'
        corresponde en realidad al reporte de 'Dollys' (trae también una
        columna DOLLY), que se importa aparte por tener columnas extra."""
        fila_norm = [self._normalizar(c) for c in fila_encabezado]
        return "DOLLY" in fila_norm

    def _parsear_tabla_camiones(self, html):
        """
        Revisa TODAS las tablas del correo (por regex) y, dentro de cada
        una, TODAS sus filas (no solo la primera) buscando la fila que sea
        el encabezado real (PLACA y CHASIS). Devuelve esa fila de encabezado
        y todas las siguientes (el llamador descarta la primera, que es el
        encabezado). Devuelve None si ninguna tabla del correo es un reporte
        de camiones.
        """
        for filas in self._tablas_html_regex(html):
            for i, fila in enumerate(filas):
                if self._fila_es_encabezado_camiones(fila):
                    return filas[i:]
        return None

    def _encabezados_tablas_html(self, html):
        """DIAGNÓSTICO: devuelve, para cada tabla encontrada en el correo,
        las primeras filas tal cual vienen (sin exigir que coincidan con
        PLACA/CHASIS), para comparar contra lo que el código espera cuando
        ninguna tabla es reconocida."""
        resultado = []
        for filas in self._tablas_html_regex(html):
            resultado.append(filas[:3])  # primeras filas de cada tabla, de muestra
        return resultado

    def _buscar_carpeta_outlook(self, carpeta, nombre_buscado):
        """Busca recursivamente, dentro de una carpeta de Outlook y sus
        subcarpetas, una carpeta cuyo nombre coincida (sin importar
        mayúsculas/acentos) con nombre_buscado."""
        if self._normalizar(carpeta.Name) == self._normalizar(nombre_buscado):
            return carpeta
        for sub in carpeta.Folders:
            encontrada = self._buscar_carpeta_outlook(sub, nombre_buscado)
            if encontrada:
                return encontrada
        return None

    def importar_correos_outlook(self):
        if not PYWIN32_DISPONIBLE or not BS4_DISPONIBLE:
            faltan = []
            if not PYWIN32_DISPONIBLE:
                faltan.append("pywin32")
            if not BS4_DISPONIBLE:
                faltan.append("beautifulsoup4")
            messagebox.showerror(
                "Falta instalar librerías",
                "Para importar desde Outlook primero instala:\n\n"
                f"pip install {' '.join(faltan)}\n\n"
                "y vuelve a abrir el programa."
            )
            return

        n_cols = len(self.columnas_actuales)
        n_cols_dollys = len(self.columnas_actuales_dollys)
        if n_cols == 0:
            messagebox.showwarning("Aviso", "No hay columnas configuradas. Actualiza la tabla primero.")
            return

        nombre_carpeta = self.entry_carpeta_outlook.get().strip()
        if not nombre_carpeta:
            messagebox.showwarning("Aviso", "Escribe el nombre de la carpeta de Outlook a revisar.")
            return

        try:
            fecha_desde = datetime.strptime(self.entry_fecha_desde.get().strip(), "%d/%m/%Y")
            fecha_hasta = datetime.strptime(self.entry_fecha_hasta.get().strip(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido en 'Desde'/'Hasta'. Usa dd/mm/yyyy.")
            return

        if fecha_hasta < fecha_desde:
            messagebox.showerror("Error", "La fecha 'Hasta' no puede ser anterior a la fecha 'Desde'.")
            return

        mapeo_flota = self._parsear_mapeo_flota()

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo conectar con Outlook (¿está abierto?):\n{e}")
            return

        carpeta = None
        for store_root in ns.Folders:
            carpeta = self._buscar_carpeta_outlook(store_root, nombre_carpeta)
            if carpeta:
                break

        if carpeta is None:
            messagebox.showerror("Error", f"No se encontró ninguna carpeta llamada '{nombre_carpeta}' en Outlook.")
            return

        total_correos = 0
        correos_omitidos_respuesta = 0
        correos_validos = 0
        correos_validos_dollys = 0
        correos_sin_reconocer = 0  # correos con tabla html pero SIN encabezado PLACA/CHASIS reconocido
        total_filas = 0
        total_filas_dollys = 0
        total_filas_duplicadas = 0  # filas omitidas por ser duplicado exacto de una ya importada
        flotas_a_revisar = set()
        muestras_encabezados = []  # DIAGNÓSTICO: encabezados reales de correos sin match
        html_diagnostico_crudo = None  # DIAGNÓSTICO: HTML crudo del primer correo sin match
        claves_vistas_camiones = set()  # para deduplicar por PLACA+CHASIS+FECHA+HORA+FLOTA
        claves_vistas_dollys = set()

        try:
            # Primero se recolectan los correos candidatos (dentro del rango
            # de fechas) junto con su fecha/hora exacta de recepción, y se
            # ordenan aquí mismo en Python de forma ASCENDENTE (más viejo
            # primero). No se depende de carpeta.Items.Sort(...) de Outlook
            # porque su comportamiento varía según el tipo de cuenta/backend
            # (Exchange, IMAP, Nuevo Outlook, etc.) y no siempre respeta el
            # orden pedido.
            candidatos = []
            for item in list(carpeta.Items):
                if getattr(item, "Class", None) != 43:  # 43 = olMail
                    continue
                recibido = item.ReceivedTime
                fecha_recibido = datetime(recibido.year, recibido.month, recibido.day)
                if not (fecha_desde <= fecha_recibido <= fecha_hasta):
                    continue
                candidatos.append((recibido, fecha_recibido, item))

            candidatos.sort(key=lambda t: t[0])  # ascendente: más viejo primero

            for recibido, fecha_recibido, item in candidatos:
                total_correos += 1

                if self.var_excluir_respuestas.get() and self._es_asunto_respuesta_o_reenvio(item.Subject or ""):
                    correos_omitidos_respuesta += 1
                    continue  # Es una respuesta/reenvío (RE:/FW:/etc.); se ignora para no duplicar

                html = item.HTMLBody
                if not html:
                    # Correo sin HTML (por ejemplo, es texto plano, o el
                    # cuerpo no cargó): también cuenta como "sin reconocer",
                    # con la hora de recepción para poder identificarlo.
                    correos_sin_reconocer += 1
                    if len(muestras_encabezados) < 3:
                        muestras_encabezados.append(
                            (item.Subject or "(sin asunto)", recibido, [])
                        )
                    continue

                filas_html = self._parsear_tabla_camiones(html)
                if not filas_html:
                    # DIAGNÓSTICO: guarda los encabezados reales de hasta 3
                    # correos para poder mostrarlos si al final no hubo
                    # ningún correo válido, y guarda el HTML CRUDO de hasta 1
                    # correo para poder mostrarlo en una ventana copiable
                    # (evita depender de rutas de archivo, que pueden fallar
                    # o ser confusas si el programa corre como .exe).
                    correos_sin_reconocer += 1
                    if len(muestras_encabezados) < 3:
                        muestras_encabezados.append(
                            (item.Subject or "(sin asunto)", recibido, self._encabezados_tablas_html(html))
                        )
                    if not html_diagnostico_crudo:
                        html_diagnostico_crudo = html
                    continue  # No es un correo con tabla de camiones; se ignora

                es_dolly = self._es_encabezado_dolly(filas_html[0])
                tree_destino = self.tree_dollys if es_dolly else self.tree
                n_cols_destino = n_cols_dollys if es_dolly else n_cols
                columnas_destino = self.columnas_actuales_dollys if es_dolly else self.columnas_actuales
                claves_vistas_destino = claves_vistas_dollys if es_dolly else claves_vistas_camiones

                if es_dolly:
                    correos_validos_dollys += 1
                else:
                    correos_validos += 1
                cuerpo = filas_html[1:]

                soup = BeautifulSoup(html, "html.parser")
                # Los reportes de Dollys no manejan FLOTA, así que no tiene
                # sentido intentar detectarla ni advertir "REVISAR" para
                # ellos — flota se deja vacía y sin marcar en ese caso.
                if es_dolly:
                    flota = ""
                else:
                    texto_monitoreo = self._extraer_texto_monitoreo(soup)
                    flota = self._detectar_flota_en_texto(texto_monitoreo, mapeo_flota)
                    if flota.startswith("REVISAR:"):
                        flotas_a_revisar.add(flota)

                hora = self._formatear_hora(recibido.strftime("%H:%M"))
                fecha = fecha_recibido.strftime("%d/%m/%Y")

                filas_normalizadas = []
                for fila in cuerpo:
                    if len(fila) < n_cols_destino:
                        fila = fila + [""] * (n_cols_destino - len(fila))
                    elif len(fila) > n_cols_destino:
                        fila = fila[:n_cols_destino]
                    filas_normalizadas.append(fila)

                # La deduplicación es por CONTENIDO (FLOTA+FECHA+HORA+PLACA+
                # CHASIS), no por el asunto del correo — así una respuesta o
                # reenvío que trae datos NUEVOS sí se importa igual, y solo
                # se omite si de verdad repite un camión ya importado antes.
                filas_insertadas = self._insertar_filas(
                    tree_destino, filas_normalizadas, flota, fecha, hora,
                    columnas_destino=columnas_destino, claves_vistas=claves_vistas_destino
                )
                total_filas_duplicadas += len(filas_normalizadas) - filas_insertadas
                if es_dolly:
                    total_filas_dollys += filas_insertadas
                else:
                    total_filas += filas_insertadas
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error leyendo los correos: {e}")
            return

        mensaje = (
            f"Correos revisados en el rango: {total_correos}\n"
            f"Respuestas/reenvíos ignorados: {correos_omitidos_respuesta}\n"
            f"Correos con tabla de camiones (normal): {correos_validos}\n"
            f"Filas agregadas (Camiones): {total_filas}\n"
            f"Correos con tabla de Dollys: {correos_validos_dollys}\n"
            f"Filas agregadas (Dollys): {total_filas_dollys}\n"
            f"⚠️ Correos SIN tabla reconocida: {correos_sin_reconocer}"
        )
        if flotas_a_revisar:
            mensaje += "\n\n⚠️ Algunas filas quedaron marcadas como 'REVISAR' en la columna FLOTA porque no se reconoció el texto de monitoreo:\n" + "\n".join(sorted(flotas_a_revisar))
            mensaje += "\n\nPuedes corregirlas con doble clic en la tabla, o ajustar el 'Mapeo de flota' arriba."

        messagebox.showinfo("Importación desde Outlook", mensaje)

        # DIAGNÓSTICO: si hubo AL MENOS UN correo sin tabla reconocida
        # (aunque otros sí se hayan importado bien), muestra los encabezados
        # reales encontrados para poder revisar por qué. Antes esto solo
        # aparecía si TODOS fallaban, dejando pasar fallas parciales sin
        # avisar.
        if correos_sin_reconocer > 0 and muestras_encabezados:
            detalle = ""
            for asunto, recibido_muestra, tablas in muestras_encabezados:
                detalle += f"\n📧 [{recibido_muestra.strftime('%d/%m/%Y %H:%M')}] {asunto}\n"
                if not tablas:
                    detalle += "   (no se encontró ninguna tabla <table> en este correo)\n"
                for i, filas_muestra in enumerate(tablas):
                    detalle += f"   Tabla {i + 1} ({len(filas_muestra)} fila(s) de muestra):\n"
                    for fila in filas_muestra:
                        detalle += f"      {fila}\n"
            nota_archivos = ""
            if html_diagnostico_crudo:
                nota_archivos = (
                    "\n\nAdemás, se abrirá otra ventana con el HTML crudo "
                    "completo de uno de esos correos, con un botón para "
                    "copiarlo — pégamelo en el chat para revisar la "
                    "estructura exacta y corregir la detección con certeza."
                )
            messagebox.showwarning(
                f"Diagnóstico: {correos_sin_reconocer} correo(s) sin tabla reconocida",
                f"{correos_sin_reconocer} correo(s) de este rango NO tuvieron ninguna "
                "tabla cuyo encabezado contenga 'PLACA' y 'CHASIS' (aunque otros correos "
                "del mismo rango sí se hayan importado bien). Estos son los encabezados "
                "REALES que sí se encontraron en algunos de esos correos, para comparar:\n"
                + detalle +
                "\nSi ahí ves las columnas pero con otro nombre/formato, o si el "
                "correo no trae ninguna tabla <table> (por ejemplo si es una imagen "
                "o un PDF adjunto), cuéntame qué aparece para ajustar la detección."
                + nota_archivos
            )
            if html_diagnostico_crudo:
                self._mostrar_ventana_html_crudo(html_diagnostico_crudo)

    def _mostrar_ventana_html_crudo(self, html):
        """DIAGNÓSTICO: abre una ventana con el HTML crudo de un correo en
        un cuadro de texto seleccionable, con un botón para copiarlo al
        portapapeles. Evita depender de rutas de archivo (que pueden fallar
        o ser confusas, sobre todo si el programa corre como .exe)."""
        top = tk.Toplevel(self.root)
        top.title("Diagnóstico: HTML crudo del correo (para copiar y compartir)")
        top.geometry("800x600")

        tk.Label(
            top,
            text="Copia todo este texto y pégalo en el chat:",
            font=("Arial", 10, "bold"),
        ).pack(pady=5)

        caja = scrolledtext.ScrolledText(top, wrap=tk.WORD, font=("Consolas", 8))
        caja.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        caja.insert("1.0", html)

        def copiar_todo():
            top.clipboard_clear()
            top.clipboard_append(html)
            messagebox.showinfo("Copiado", "El HTML se copió al portapapeles.")

        tk.Button(
            top, text="📋 Copiar todo", font=("Arial", 10, "bold"),
            bg="#2E7D32", fg="white", command=copiar_todo
        ).pack(pady=5)

    def _leer_html_portapapeles(self):
        """
        Lee el fragmento HTML que Windows guarda en el portapapeles cuando se
        copia una tabla desde Outlook, Word o Excel (formato registrado
        'HTML Format'). Ese formato trae un pequeño encabezado de texto con
        los offsets StartHTML/EndHTML antes del HTML real.
        Devuelve el HTML como texto, o None si no hay ese formato disponible.
        """
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(cf_html):
                return None
            datos = win32clipboard.GetClipboardData(cf_html)
        finally:
            win32clipboard.CloseClipboard()

        if isinstance(datos, bytes):
            datos = datos.decode('utf-8', errors='ignore')

        m_start = re.search(r'StartHTML:(\d+)', datos)
        m_end = re.search(r'EndHTML:(\d+)', datos)
        if m_start and m_end:
            inicio = int(m_start.group(1))
            fin = int(m_end.group(1))
            return datos[inicio:fin]
        return datos  # Por si no trae encabezado, se devuelve tal cual

    def _parsear_tabla_html(self, html):
        """Devuelve una lista de filas (cada una lista de textos de celda,
        preservando las celdas vacías) a partir de la PRIMERA tabla <table>
        encontrada en el HTML."""
        soup = BeautifulSoup(html, "html.parser")
        tabla = soup.find("table")
        if not tabla:
            return None
        filas = []
        for tr in tabla.find_all("tr"):
            celdas = tr.find_all(["td", "th"])
            if not celdas:
                continue
            fila = [c.get_text(separator=" ", strip=True) for c in celdas]
            filas.append(fila)
        return filas

    def pegar_html_portapapeles(self):
        if not PYWIN32_DISPONIBLE or not BS4_DISPONIBLE:
            faltan = []
            if not PYWIN32_DISPONIBLE:
                faltan.append("pywin32")
            if not BS4_DISPONIBLE:
                faltan.append("beautifulsoup4")
            messagebox.showerror(
                "Falta instalar librerías",
                "Para usar 'Pegar Tabla con Formato' primero instala:\n\n"
                f"pip install {' '.join(faltan)}\n\n"
                "y vuelve a abrir el programa."
            )
            return

        n_cols = len(self.columnas_actuales)
        if n_cols == 0:
            messagebox.showwarning("Aviso", "No hay columnas configuradas. Actualiza la tabla primero.")
            return

        try:
            html = self._leer_html_portapapeles()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el portapapeles: {e}")
            return

        if not html:
            messagebox.showwarning(
                "Aviso",
                "El portapapeles no contiene una tabla con formato.\n\n"
                "Copia la tabla directamente desde Outlook o Excel (Ctrl+C) "
                "e intenta de nuevo."
            )
            return

        filas_html = self._parsear_tabla_camiones(html)
        if not filas_html:
            # DIAGNÓSTICO: muestra los encabezados reales encontrados en lo
            # copiado, para comparar contra lo que se esperaba (PLACA/CHASIS).
            tablas = self._encabezados_tablas_html(html)
            if tablas:
                partes = []
                for i, filas_muestra in enumerate(tablas):
                    partes.append(f"Tabla {i+1}:\n" + "\n".join(f"   {fila}" for fila in filas_muestra))
                detalle = "\n".join(partes)
            else:
                detalle = "(no se encontró ninguna tabla <table> en el portapapeles)"
            messagebox.showwarning(
                "Aviso",
                "No se encontró ninguna tabla de camiones (con columnas como PLACA/CHASIS) en lo copiado.\n\n"
                "Encabezados reales encontrados:\n" + detalle + "\n\n"
                "Asegúrate de seleccionar y copiar la tabla completa (no solo el texto)."
            )
            return

        # Detecta si la tabla copiada es del tipo 'Dollys' (trae columna
        # DOLLY en su encabezado) para usar su propio set de columnas y
        # su propia tabla, en vez de mezclarla con la de camiones normales.
        es_dolly = self._es_encabezado_dolly(filas_html[0])
        tree_destino = self.tree_dollys if es_dolly else self.tree
        columnas_destino = self.columnas_actuales_dollys if es_dolly else self.columnas_actuales
        n_cols_destino = len(columnas_destino)

        # Saltar la fila de encabezados si la primera celda coincide con el
        # nombre de la primera columna configurada.
        encabezado_esperado = columnas_destino[0].strip().upper().rstrip('.')
        if filas_html and filas_html[0][0].strip().upper().rstrip('.') in (encabezado_esperado, "NO", "NUMERO", "#"):
            filas_html = filas_html[1:]

        filas_finales = []
        for fila in filas_html:
            if len(fila) < n_cols_destino:
                fila = fila + [""] * (n_cols_destino - len(fila))
            elif len(fila) > n_cols_destino:
                fila = fila[:n_cols_destino]
            filas_finales.append(fila)

        agregadas = self._agregar_filas_al_arbol(filas_finales, tree_destino)
        if agregadas:
            destino_txt = "Dollys" if es_dolly else "Camiones"
            messagebox.showinfo("Éxito", f"Se agregaron {agregadas} filas (bloque FLOTA/HORA actual) a la tabla '{destino_txt}', sin borrar lo que ya había.")

    def _dividir_linea_tabulada(self, linea):
        """Divide una línea con tabulaciones en celdas, preservando las vacías
        (dos tabs seguidos = una celda vacía)."""
        return [p.strip() for p in linea.split('\t')]

    def _normalizar(self, texto):
        """Deja solo letras/números en mayúscula, sin acentos ni espacios,
        para poder comparar nombres de columna sin importar tildes/guiones."""
        texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
        return re.sub(r'[^A-Z0-9]', '', texto.upper())

    def _indice_columna_en(self, columnas, *variantes):
        """Igual que _indice_columna, pero recibe explícitamente la lista de
        columnas en la que buscar (para poder usarla tanto con
        self.columnas_actuales como con self.columnas_actuales_dollys)."""
        for i, col in enumerate(columnas):
            if self._normalizar(col) in variantes:
                return i
        return None

    def _indice_columna(self, *variantes):
        """Devuelve el índice (dentro de self.columnas_actuales) de la primera
        columna cuyo nombre normalizado coincida con alguna de las variantes,
        o None si no existe esa columna en la configuración actual."""
        return self._indice_columna_en(self.columnas_actuales, *variantes)

    def procesar_tabla_pegada(self):
        contenido_original = self.text_area.get("1.0", tk.END)
        if not contenido_original.strip():
            messagebox.showwarning("Aviso", "Pega los datos en el cuadro.")
            return

        n_cols = len(self.columnas_actuales)
        if n_cols == 0:
            messagebox.showwarning("Aviso", "No hay columnas configuradas. Actualiza la tabla primero.")
            return

        contiene_tabs = '\t' in contenido_original

        if contiene_tabs:
            # --- Formato TABLA (una fila completa por línea, celdas separadas por TAB) ---
            lineas = [l for l in contenido_original.splitlines() if l.strip() != ""]
            if not lineas:
                messagebox.showwarning("Aviso", "No se encontraron filas para procesar.")
                return

            primera_celda = self._dividir_linea_tabulada(lineas[0])[0].strip().upper().rstrip('.')
            encabezado_esperado = self.columnas_actuales[0].strip().upper().rstrip('.')
            if primera_celda == encabezado_esperado or primera_celda in ("NO", "NUMERO", "#"):
                lineas = lineas[1:]

            filas_finales = []
            for linea in lineas:
                celdas = self._dividir_linea_tabulada(linea)
                if len(celdas) < n_cols:
                    celdas = celdas + [""] * (n_cols - len(celdas))
                elif len(celdas) > n_cols:
                    celdas = celdas[:n_cols]
                filas_finales.append(celdas)
        else:
            # --- Formato VERTICAL con reconocimiento por patrones ---
            # Ignora por completo las líneas en blanco (no dependemos de contarlas):
            # cada campo se identifica por su formato característico, no por su posición.
            lineas = [l.strip() for l in contenido_original.splitlines() if l.strip() != ""]
            if not lineas:
                messagebox.showwarning("Aviso", "No se encontraron datos para procesar.")
                return

            # Un nuevo camión inicia cuando aparece un número de orden (1-999)
            # y ya llevamos varios campos acumulados del camión anterior.
            camiones = []
            c_actual = []
            for linea in lineas:
                if linea.isdigit() and 1 <= int(linea) <= 999 and len(c_actual) >= 5:
                    camiones.append(c_actual)
                    c_actual = [linea]
                else:
                    c_actual.append(linea)
            if c_actual:
                camiones.append(c_actual)

            idx_no = self._indice_columna("NO")
            idx_nombre = self._indice_columna("NOMBRE")
            idx_licencia = self._indice_columna("LICENCIA")
            idx_placa = self._indice_columna("PLACA")
            idx_telefono = self._indice_columna("TELEFONO")
            idx_chasis = self._indice_columna("CHASIS")
            idx_tcchasis = self._indice_columna("TCCHASIS")
            idx_genset = self._indice_columna("GENSET")
            idx_contenedor = self._indice_columna("CONTENEDOR")
            idx_ubicacion = self._indice_columna("UBICACION")
            idx_status = self._indice_columna("STATUS", "ESTADO")
            idx_destino = self._indice_columna("DESTINO")

            estados_posibles = ["CARGADO", "VACIO", "SOLO CHASIS", "CARGANDO", "EXPORTACION"]

            filas_finales = []
            for c in list(camiones):
                c = list(c)
                fila = [""] * n_cols

                if idx_no is not None and c and c[0].isdigit():
                    fila[idx_no] = c.pop(0)

                if idx_nombre is not None and c and not any(re.search(pat, c[0]) for pat in [r'^\d{4}', r'^[A-Z0-9]{5,}', r'^TC-', r'^TJG-', r'CARGADO|VACIO|SOLO CHASIS|CARGANDO|EXPORTACION']):
                    fila[idx_nombre] = c.pop(0)

                if idx_licencia is not None:
                    for i, val in enumerate(c):
                        if re.search(r'\d{4}\s+\d+', val):
                            fila[idx_licencia] = val
                            c.pop(i)
                            break

                if idx_placa is not None:
                    for i, val in enumerate(c):
                        if re.search(r'^[0-9]{3,4}[A-Z]{2,3}$', val) or (len(val) <= 7 and re.search(r'[A-Z]', val) and re.search(r'\d', val) and ' ' not in val and '-' not in val):
                            fila[idx_placa] = val
                            c.pop(i)
                            break

                if idx_telefono is not None:
                    for i, val in enumerate(c):
                        if ('-' in val or '/' in val) and 'TMGS' not in val and 'TEMU' not in val and 'TC-' not in val:
                            fila[idx_telefono] = val
                            c.pop(i)
                            break

                if idx_chasis is not None:
                    for i, val in enumerate(c):
                        if 'TMGS' in val or 'NAV' in val or 'TGZ' in val:
                            fila[idx_chasis] = val
                            c.pop(i)
                            break

                if idx_tcchasis is not None:
                    for i, val in enumerate(c):
                        if val.startswith('TC-'):
                            fila[idx_tcchasis] = val
                            c.pop(i)
                            break

                if idx_genset is not None:
                    for i, val in enumerate(c):
                        if 'TJG-' in val:
                            fila[idx_genset] = val
                            c.pop(i)
                            break

                if idx_contenedor is not None:
                    for i, val in enumerate(c):
                        if re.search(r'^[A-Z]{4}\d{6,}', val):
                            fila[idx_contenedor] = val
                            c.pop(i)
                            break

                if idx_status is not None:
                    for i, val in enumerate(c):
                        if val.upper() in estados_posibles:
                            fila[idx_status] = val
                            c.pop(i)
                            break

                # Lo que sobre se reparte, en orden, en los huecos vacíos restantes
                # (típicamente UBICACIÓN y DESTINO).
                huecos_vacios = [i for i, x in enumerate(fila) if x == ""]
                while c and huecos_vacios:
                    fila[huecos_vacios.pop(0)] = c.pop(0)

                filas_finales.append(fila)

        agregadas = self._agregar_filas_al_arbol(filas_finales)
        if agregadas:
            self.text_area.delete("1.0", tk.END)
            messagebox.showinfo("Éxito", f"Se agregaron {agregadas} camiones (bloque FLOTA/HORA actual) a la tabla, sin borrar lo que ya había.")

    def _formatear_hora(self, hora_str):
        """Convierte '2:30' en '2:30:00'. Si ya trae segundos ('2:30:00') o
        viene vacía/():no válida, se devuelve tal cual (o vacía)."""
        hora_str = hora_str.strip()
        if not hora_str:
            return ""
        partes = hora_str.split(":")
        try:
            h = int(partes[0])
            m = int(partes[1]) if len(partes) > 1 else 0
            s = int(partes[2]) if len(partes) > 2 else 0
            return f"{h}:{m:02d}:{s:02d}"
        except (ValueError, IndexError):
            # No se pudo interpretar como hora; se devuelve el texto original
            return hora_str

    def guardar_excel(self):
        filas_camiones = []
        for item in self.tree.get_children():
            valores = self.tree.item(item, "values")
            if any(str(v).strip() for v in valores):
                filas_camiones.append(valores)

        filas_dollys = []
        for item in self.tree_dollys.get_children():
            valores = self.tree_dollys.item(item, "values")
            if any(str(v).strip() for v in valores):
                filas_dollys.append(valores)

        if not filas_camiones and not filas_dollys:
            messagebox.showwarning("Aviso", "No hay datos en ninguna de las dos tablas para guardar.")
            return

        carpeta = filedialog.askdirectory(title="Seleccione la carpeta de destino")
        if not carpeta:
            return

        archivo = os.path.join(carpeta, f"Registro_Camiones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

        hf = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        hfont = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        dfont = Font(name="Arial", size=10)
        border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

        def escribir_hoja(ws, columnas, filas):
            for c_idx, h in enumerate(columnas, 1):
                cell = ws.cell(row=1, column=c_idx, value=h)
                cell.fill = hf
                cell.font = hfont
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for r_idx, r_vals in enumerate(filas, 2):
                for c_idx, val in enumerate(r_vals, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = dfont
                    cell.border = border
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            for col in ws.columns:
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = 18

            ws.freeze_panes = 'A2'

        try:
            wb = openpyxl.Workbook()
            primera_hoja_usada = False

            # Cada fila ya trae su propia FLOTA/FECHA/HORA (asignadas al momento
            # en que se agregó ese bloque), así que se exportan tal cual están.
            if filas_camiones:
                ws = wb.active
                ws.title = "Datos"
                escribir_hoja(ws, self.columnas_tabla, filas_camiones)
                primera_hoja_usada = True

            if filas_dollys:
                ws_d = wb.active if not primera_hoja_usada else wb.create_sheet("Dollys")
                if not primera_hoja_usada:
                    ws_d.title = "Dollys"
                escribir_hoja(ws_d, self.columnas_tabla_dollys, filas_dollys)

            wb.save(archivo)
            messagebox.showinfo("¡Guardado Exitoso!", "Archivo Excel generado con éxito.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AdaptableInteligenteApp(root)
    root.mainloop()
